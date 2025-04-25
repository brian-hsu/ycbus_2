from typing import Dict
import argparse
from ycbus_v2 import BusBookingSystem, BookingData
from utils.config_loader import load_config
from utils.notification import LineNotifier
from utils.email_notification import EmailNotifier
from read_google_sheet import ReadGSheet
from utils.captcha_handler import CaptchaHandler
import time
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import urllib.request
import os
from selenium.webdriver.firefox.options import Options
from PIL import Image
import requests
import random
import base64
from collections import Counter
from utils.gmail_sender import GmailSender
from config import BASE_URL


def parse_arguments():
    parser = argparse.ArgumentParser(description="YC Bus Booking System")
    parser.add_argument(
        "--mode",
        choices=["desktop", "server", "debug"],
        default="server",
        help="運行模式",
    )
    parser.add_argument("--headless", action="store_true", help="是否使用無頭模式")
    return parser.parse_args()


def load_data_from_txt():
    """從 data.txt 讀取基本配置"""
    with open("data.txt", "r", encoding="utf-8") as file:
        lines = file.readlines()

    result = {}
    for line in lines:
        if ':' in line:  # 確保行包含冒號
            key, value = line.strip().split(":", 1)  # 只分割第一個冒號
            result[key.strip()] = value.strip()
    return result


def load_data_from_gsheet():
    """從 Google Sheet 讀取資料"""
    data = load_data_from_txt()
    read_gc = ReadGSheet()
    gc_dict = read_gc.gsheet_cover(data["gsheet_cover"])

    booking_data = {
        "name": data["name"],
        "num": data["ycbus_password"],
        "date": gc_dict["ride_date"],
        "go_time": gc_dict["goto_time"],
        "back_time": gc_dict["return_time"],
        "goto_pickup_area": gc_dict["goto_pickup_area"],
        "goto_dropoff_area": gc_dict["goto_dropoff_area"],
        "goto_pickup_address": gc_dict["goto_pickup_address"],
        "goto_dropoff_address": gc_dict["goto_dropoff_address"],
        "return_pickup_area": gc_dict["return_pickup_area"],
        "return_dropoff_area": gc_dict["return_dropoff_area"],
        "return_pickup_address": gc_dict["return_pickup_address"],
        "return_dropoff_address": gc_dict["return_dropoff_address"],
        "Message": gc_dict["note_message"],
    }

    # 處理回程地址的特殊情況
    if booking_data["return_pickup_area"] == "same_goto_dropoff":
        booking_data["return_pickup_area"] = booking_data["goto_dropoff_area"]

    if booking_data["return_pickup_address"] == "same_goto_dropoff":
        booking_data["return_pickup_address"] = booking_data["goto_dropoff_address"]

    if booking_data["return_dropoff_area"] == "same_pickup":
        booking_data["return_dropoff_area"] = booking_data["goto_pickup_area"]

    if booking_data["return_dropoff_address"] == "same_pickup":
        booking_data["return_dropoff_address"] = booking_data["goto_pickup_address"]

    # 返回预约数据和通知相关信息
    notification_data = {
        "line_token": data.get("line_token", ""),
        "gmail_sender": data.get("gmail_sender", ""),
        "gmail_password": data.get("gmail_password", ""),
        "recipient_emails": [email.strip() for email in data.get("recipient_emails", "").split(",") if email.strip()]
    }

    # 檢查必要的電子郵件設定
    if not notification_data["gmail_sender"] or not notification_data["gmail_password"] or not notification_data["recipient_emails"]:
        print("警告：電子郵件設定不完整")
        print(f"gmail_sender: {notification_data['gmail_sender']}")
        print(f"gmail_password: {notification_data['gmail_password']}")
        print(f"recipient_emails: {notification_data['recipient_emails']}")
        raise ValueError("請在 data.txt 中配置完整的電子郵件設定")

    return booking_data, notification_data


def handle_login_process(system):
    """處理登入流程，包含驗證碼處理"""
    max_attempts = 5  # 增加嘗試次數
    for attempt in range(max_attempts):
        try:
            print(f"開始第 {attempt + 1} 次登入嘗試")

            # 確保頁面已經加載
            system.navigate_to_login_page()

            # 等待驗證碼圖片載入，嘗試多種定位方式
            captcha_img = None
            locators = [
                (By.ID, "captchaImage"),
                (By.CSS_SELECTOR, "img[src*='captcha']"),
                (By.CSS_SELECTOR, "img[alt*='captcha']"),
                (By.XPATH, "//img[contains(@src, 'captcha')]"),
                (By.XPATH, "//img[contains(@alt, 'captcha')]"),
            ]

            for locator in locators:
                try:
                    captcha_img = WebDriverWait(system.driver, 5).until(
                        EC.presence_of_element_located(locator)
                    )
                    if captcha_img:
                        print(f"成功使用定位器 {locator} 找到驗證碼圖片")
                        break
                except:
                    continue

            if not captcha_img:
                print("無法找到驗證碼圖片，嘗試截取整個頁面")
                # 截取整個頁面以便調試
                debug_screenshot = os.path.join(
                    "temp_captcha", f"page_screenshot_{attempt}.png"
                )
                system.driver.save_screenshot(debug_screenshot)
                print(f"已保存頁面截圖至: {debug_screenshot}")

                # 嘗試重新加載頁面
                system.driver.refresh()
                time.sleep(3)
                continue

            # 確保圖片已完全載入
            time.sleep(2)

            # 直接使用截圖方式獲取驗證碼圖片（優先使用此方法）
            debug_img_path = os.path.join("temp_captcha", f"captcha_{attempt}.png")
            try:
                # 獲取驗證碼元素的位置和大小
                location = captcha_img.location
                size = captcha_img.size

                # 截取整個頁面
                system.driver.save_screenshot(debug_img_path)

                # 從截圖中裁剪出驗證碼部分
                full_img = Image.open(debug_img_path)
                left = location["x"]
                top = location["y"]
                right = location["x"] + size["width"]
                bottom = location["y"] + size["height"]

                # 裁剪並保存驗證碼圖片
                captcha_img_cropped = full_img.crop((left, top, right, bottom))
                captcha_img_cropped.save(debug_img_path)
                print(f"已通過截圖方式保存驗證碼圖片至: {debug_img_path}")
            except Exception as crop_error:
                print(f"截圖獲取驗證碼失敗: {str(crop_error)}")

                # 備用方法：嘗試使用JavaScript直接獲取圖片的base64編碼
                try:
                    print("嘗試使用JavaScript獲取驗證碼圖片...")

                    # 使用JavaScript獲取圖片的base64編碼
                    img_base64 = system.driver.execute_script(
                        """
                        var img = arguments[0];
                        var canvas = document.createElement('canvas');
                        canvas.width = img.width;
                        canvas.height = img.height;
                        var ctx = canvas.getContext('2d');
                        ctx.drawImage(img, 0, 0);
                        return canvas.toDataURL('image/png').substring(22);
                    """,
                        captcha_img,
                    )

                    # 將base64轉換為圖片並保存
                    if img_base64:
                        with open(debug_img_path, "wb") as f:
                            f.write(base64.b64decode(img_base64))
                        print(
                            f"已通過JavaScript獲取並保存驗證碼圖片至: {debug_img_path}"
                        )
                    else:
                        raise Exception("無法獲取圖片的base64編碼")
                except Exception as js_error:
                    print(f"JavaScript獲取驗證碼失敗: {str(js_error)}")

                    # 最後嘗試：使用URL下載方式
                    try:
                        # 獲取圖片的完整URL
                        img_url = captcha_img.get_attribute("src")
                        print(f"驗證碼圖片URL: {img_url}")

                        if not img_url or img_url == "":
                            print("無法獲取驗證碼圖片URL，嘗試使用JavaScript獲取")
                            img_url = system.driver.execute_script(
                                "return arguments[0].src;", captcha_img
                            )
                            print(f"使用JavaScript獲取的URL: {img_url}")

                        if not img_url or img_url == "":
                            print("仍然無法獲取驗證碼圖片URL，跳過此次嘗試")
                            system.driver.refresh()
                            time.sleep(3)
                            continue

                        # 修正URL處理
                        if not img_url.startswith("http"):
                            # 檢查是否是相對於根目錄的路徑
                            if img_url.startswith("/"):
                                # 獲取域名部分
                                domain_parts = system.driver.current_url.split("/")
                                base_url = f"{domain_parts[0]}//{domain_parts[2]}"
                                img_url = f"{base_url}{img_url}"
                            else:
                                # 獲取當前頁面的基礎URL
                                base_url = system.driver.current_url.rsplit("/", 1)[0]
                                img_url = f"{base_url}/{img_url}"

                        # 添加隨機參數避免快取
                        img_url = f"{img_url}{'&' if '?' in img_url else '?'}random={random.randint(1, 100000)}"
                        print(f"處理後的圖片URL: {img_url}")

                        # 使用requests庫下載圖片
                        response = requests.get(img_url, stream=True, timeout=10)
                        if response.status_code == 200:
                            with open(debug_img_path, "wb") as f:
                                f.write(response.content)
                            print(f"已通過URL下載並保存驗證碼圖片至: {debug_img_path}")
                        else:
                            raise Exception(
                                f"下載圖片失敗，狀態碼: {response.status_code}"
                            )
                    except Exception as url_error:
                        print(f"URL下載驗證碼失敗: {str(url_error)}")
                        system.driver.refresh()
                        time.sleep(3)
                        continue

            # 使用本地圖片路徑進行驗證碼識別
            captcha_handler = CaptchaHandler(system.driver)
            captcha_code = captcha_handler.recognize_captcha(debug_img_path)

            if not captcha_code:
                print("驗證碼識別失敗，重試中...")
                system.driver.refresh()
                time.sleep(3)
                continue

            print(f"識別出的驗證碼: {captcha_code}")

            # 嘗試登入
            try:
                # 使用更靈活的方式填寫登入表單
                try:
                    # 嘗試填寫帳戶名稱
                    username_field = None
                    username_locators = [
                        (By.CSS_SELECTOR, "#cusname"),
                        (By.ID, "cusname"),
                        (By.NAME, "cusname"),
                        (By.XPATH, "//input[@placeholder='姓名']"),
                        (By.XPATH, "//input[contains(@id, 'name')]"),
                    ]

                    for locator in username_locators:
                        try:
                            username_field = WebDriverWait(system.driver, 5).until(
                                EC.presence_of_element_located(locator)
                            )
                            if username_field:
                                break
                        except:
                            continue

                    if not username_field:
                        print("無法找到用戶名輸入框")
                        system.driver.refresh()
                        time.sleep(3)
                        continue

                    username_field.clear()
                    username_field.send_keys(system.booking_data.name)
                    print("已填寫帳戶名稱")

                    # 嘗試填寫乘客編號
                    password_field = None
                    password_locators = [
                        (By.CSS_SELECTOR, "#idcode"),
                        (By.ID, "idcode"),
                        (By.NAME, "idcode"),
                        (By.XPATH, "//input[@placeholder='乘客編號']"),
                        (By.XPATH, "//input[contains(@id, 'code')]"),
                    ]

                    for locator in password_locators:
                        try:
                            password_field = WebDriverWait(system.driver, 5).until(
                                EC.presence_of_element_located(locator)
                            )
                            if password_field:
                                break
                        except:
                            continue

                    if not password_field:
                        print("無法找到乘客編號輸入框")
                        system.driver.refresh()
                        time.sleep(3)
                        continue

                    password_field.clear()
                    password_field.send_keys(system.booking_data.num)
                    print("已填寫乘客編號")

                    # 嘗試填寫驗證碼
                    captcha_field = None
                    captcha_locators = [
                        (By.CSS_SELECTOR, "#captcha"),
                        (By.ID, "captcha"),
                        (By.NAME, "captcha"),
                        (By.XPATH, "//input[@placeholder='驗證碼']"),
                        (By.XPATH, "//input[contains(@id, 'captcha')]"),
                    ]

                    for locator in captcha_locators:
                        try:
                            captcha_field = WebDriverWait(system.driver, 5).until(
                                EC.presence_of_element_located(locator)
                            )
                            if captcha_field:
                                break
                        except:
                            continue

                    if not captcha_field:
                        print("無法找到驗證碼輸入框")
                        system.driver.refresh()
                        time.sleep(3)
                        continue

                    captcha_field.clear()
                    captcha_field.send_keys(captcha_code)
                    print(f"已填寫驗證碼: {captcha_code}")

                    # 嘗試點擊登入按鈕
                    login_button = None
                    login_button_locators = [
                        (By.CSS_SELECTOR, "#btn101"),
                        (By.ID, "btn101"),
                        (By.XPATH, "//input[@type='button' and @value='登入']"),
                        (By.XPATH, "//button[contains(text(), '登入')]"),
                        (By.XPATH, "//input[contains(@value, '登入')]"),
                    ]

                    for locator in login_button_locators:
                        try:
                            login_button = WebDriverWait(system.driver, 5).until(
                                EC.element_to_be_clickable(locator)
                            )
                            if login_button:
                                break
                        except:
                            continue

                    if not login_button:
                        print("無法找到登入按鈕")
                        system.driver.refresh()
                        time.sleep(3)
                        continue

                    login_button.click()
                    print("已點擊登入按鈕")

                    # 等待頁面反應
                    time.sleep(3)

                    # 檢查是否登入成功 - 改為檢查特定按鈕是否存在
                    try:
                        # 尋找特定按鈕元素
                        success_buttons = [
                            "查看預約趟",
                            "查今日車趟_車號(含臨時車)",
                            "查明日車趟_車號",
                            "預約訂車",
                            "查詢預約",
                            "取消預約"
                        ]

                        # 等待頁面完全加載
                        time.sleep(2)
                        
                        # 檢查當前URL
                        current_url = system.driver.current_url
                        if "netbook/book.php" in current_url or "book.php" in current_url:
                            print("通過URL檢查確認登入成功！")
                            return True
                            
                        # 檢查是否有任何一個按鈕存在
                        for button_text in success_buttons:
                            try:
                                button = system.driver.find_element(
                                    By.XPATH, f"//input[@value='{button_text}']"
                                )
                                if button and button.is_displayed():
                                    print(f"找到按鈕: {button_text}，登入成功！")
                                    return True
                            except:
                                continue
                                
                        # 檢查是否有錯誤訊息
                        try:
                            error_elements = system.driver.find_elements(
                                By.CSS_SELECTOR, ".alert-danger, .error, .w3-red"
                            )
                            for error in error_elements:
                                if error.is_displayed():
                                    print(f"登入失敗 - 出現錯誤訊息: {error.text}")
                                    return False
                        except:
                            pass
                            
                        # 如果沒有找到任何按鈕，但URL已改變，也視為登入成功
                        if current_url != BASE_URL:
                            print("URL已改變，登入成功！")
                            return True
                            
                        print("未找到登入成功後的按鈕")
                        return False
                    except Exception as e:
                        print(f"檢查登入按鈕時出錯: {str(e)}")
                        return False

                except Exception as form_error:
                    print(f"填寫表單時發生錯誤: {str(form_error)}")
                    system.driver.refresh()
                    time.sleep(3)

            except Exception as login_error:
                print(f"登入過程中發生錯誤: {str(login_error)}")
                try:
                    alert = system.driver.switch_to.alert
                    alert.accept()
                except:
                    pass
                system.driver.refresh()
                time.sleep(3)

        except Exception as e:
            print(f"登入嘗試 {attempt + 1} 失敗: {str(e)}")
            if attempt < max_attempts - 1:
                print("準備進行下一次嘗試...")
                try:
                    alert = system.driver.switch_to.alert
                    alert.accept()
                except:
                    pass
                system.driver.refresh()
                time.sleep(3)

    print("已達到最大重試次數，登入失敗")
    return False


def main():
    try:
        print("開始執行預約程序...")

        # 創建臨時目錄存放驗證碼圖片
        if not os.path.exists("temp_captcha"):
            os.makedirs("temp_captcha")
            print("已創建臨時目錄: temp_captcha")
        else:
            # 清理舊的驗證碼圖片
            for file in os.listdir("temp_captcha"):
                try:
                    os.remove(os.path.join("temp_captcha", file))
                except:
                    pass
            print("已清理臨時目錄中的舊文件")

        args = parse_arguments()
        args.headless = True
        print(f"運行模式: {args.mode}, 無頭模式: {args.headless}")

        try:
            booking_data_dict, notification_data = load_data_from_gsheet()
            print("成功從 Google Sheet 讀取預約資料")
            print(booking_data_dict)
        except Exception as e:
            print(f"從 Google Sheet 讀取資料失敗: {str(e)}")
            print("嘗試從本地文件讀取資料...")
            data = load_data_from_txt()
            booking_data_dict = {
                "name": data.get("name", ""),
                "num": data.get("ycbus_password", ""),
                "date": data.get("system_booking_date", ""),
                "go_time": data.get("goto_time", ""),
                "back_time": data.get("return_time", ""),
                "goto_pickup_area": data.get("goto_pickup_area", ""),
                "goto_dropoff_area": data.get("goto_dropoff_area", ""),
                "goto_pickup_address": data.get("goto_pickup_address", ""),
                "goto_dropoff_address": data.get("goto_dropoff_address", ""),
                "return_pickup_area": data.get("return_pickup_area", ""),
                "return_dropoff_area": data.get("return_dropoff_area", ""),
                "return_pickup_address": data.get("return_pickup_address", ""),
                "return_dropoff_address": data.get("return_dropoff_address", ""),
                "Message": data.get("note_message", "")
            }
            notification_data = {
                "line_token": data.get("line_token", ""),
                "gmail_sender": data.get("gmail_sender", ""),
                "gmail_password": data.get("gmail_password", ""),
                "recipient_emails": [email.strip() for email in data.get("recipient_emails", "").split(",") if email.strip()]
            }

        booking_data = BookingData(**booking_data_dict)
        
        # 創建通知器 - 電子郵件通知變成可選的
        notifier = None
        if (notification_data.get("gmail_sender") and 
            notification_data.get("gmail_password") and 
            notification_data.get("recipient_emails")):
            print("使用電子郵件進行通知")
            notifier = EmailNotifier(
                sender_email=notification_data["gmail_sender"],
                app_password=notification_data["gmail_password"],
                recipient_emails=notification_data["recipient_emails"],
                sender_name="預約系統通知"
            )
        else:
            print("警告：未配置電子郵件通知所需的設定，將繼續執行但不發送通知")

        # 創建Firefox選項並設定偏好
        firefox_options = Options()
        firefox_options.set_preference("browser.tabs.warnOnRepost", False)
        # 關閉其他可能的對話框
        firefox_options.set_preference("dom.successive_dialog_time_limit", 0)
        firefox_options.set_preference("dom.disable_beforeunload", True)
        # 禁用 PDF 查看器
        firefox_options.set_preference("pdfjs.disabled", True)
        # 禁用緩存
        firefox_options.set_preference("browser.cache.disk.enable", False)
        firefox_options.set_preference("browser.cache.memory.enable", False)
        firefox_options.set_preference("browser.cache.offline.enable", False)
        firefox_options.set_preference("network.http.use-cache", False)
        # 添加處理重複提交警告的設定
        firefox_options.set_preference("browser.formfill.enable", False)
        firefox_options.set_preference("browser.sessionstore.resume_from_crash", False)
        firefox_options.set_preference("browser.sessionstore.resume_session_once", False)
        firefox_options.set_preference("browser.sessionstore.max_resumed_crashes", 0)
        firefox_options.set_preference("browser.sessionstore.warnOnQuit", False)
        firefox_options.set_preference("browser.sessionstore.enabled", False)

        # 如果需要無頭模式
        if args.headless:
            firefox_options.add_argument("--headless")
            print("已啟用無頭模式")

        print("正在初始化瀏覽器...")
        system = BusBookingSystem(
            booking_data=booking_data,
            browser_type="firefox",
            options=firefox_options,  # 使用options而非firefox_profile
        )
        print("瀏覽器初始化完成")

        try:
            # 新增登入處理流程
            print("開始登入流程...")
            if not handle_login_process(system):
                error_msg = "登入失敗，無法完成預約"
                print(error_msg)
                try:
                    if notifier:
                        notifier.send_notification(error_msg)
                except Exception as notify_error:
                    print(f"發送通知失敗: {str(notify_error)}")
                return

            print("登入成功，開始預約流程...")
            success, screenshot_path = system.book_journey()
            if success:
                success_msg = "預約成功"
                print(success_msg)
                
                if notifier:  # 只有在有通知器的情況下才發送通知
                    try:
                        # 準備郵件內容
                        text_content = f"""
預約成功通知
====================
預約日期：{booking_data.date}
去程時間：{booking_data.go_time}
回程時間：{booking_data.back_time}
去程上車地點：{booking_data.goto_pickup_address}
去程下車地點：{booking_data.goto_dropoff_address}
回程上車地點：{booking_data.return_pickup_address}
回程下車地點：{booking_data.return_dropoff_address}
備註訊息：{booking_data.Message}
====================
此為自動發送的通知郵件，請勿直接回覆。
"""
                        
                        html_content = f"""
<html>
<head>
    <style>
        body {{ font-family: Arial, sans-serif; }}
        .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
        .header {{ background-color: #4CAF50; color: white; padding: 10px; text-align: center; }}
        .content {{ padding: 20px; }}
        .footer {{ font-size: 12px; color: #666; text-align: center; margin-top: 20px; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>預約成功通知</h2>
        </div>
        <div class="content">
            <p><strong>預約日期：</strong>{booking_data.date}</p>
            <p><strong>去程時間：</strong>{booking_data.go_time}</p>
            <p><strong>回程時間：</strong>{booking_data.back_time}</p>
            <p><strong>去程上車地點：</strong>{booking_data.goto_pickup_address}</p>
            <p><strong>去程下車地點：</strong>{booking_data.goto_dropoff_address}</p>
            <p><strong>回程上車地點：</strong>{booking_data.return_pickup_address}</p>
            <p><strong>回程下車地點：</strong>{booking_data.return_dropoff_address}</p>
            <p><strong>備註訊息：</strong>{booking_data.Message}</p>
        </div>
        <div class="footer">
            此為自動發送的通知郵件，請勿直接回覆。
        </div>
    </div>
</body>
</html>
"""
                        
                        # 檢查截圖是否存在
                        attachments = []
                        if screenshot_path:
                            try:
                                # 檢查檔案是否存在
                                if not os.path.exists(screenshot_path):
                                    print(f"錯誤：截圖檔案不存在: {screenshot_path}")
                                    raise FileNotFoundError(f"截圖檔案不存在: {screenshot_path}")
                                
                                # 檢查檔案是否可讀
                                if not os.access(screenshot_path, os.R_OK):
                                    print(f"錯誤：截圖檔案無法讀取: {screenshot_path}")
                                    raise PermissionError(f"截圖檔案無法讀取: {screenshot_path}")
                                
                                # 檢查檔案大小
                                file_size = os.path.getsize(screenshot_path)
                                if file_size == 0:
                                    print(f"錯誤：截圖檔案大小為 0: {screenshot_path}")
                                    raise ValueError(f"截圖檔案大小為 0: {screenshot_path}")
                                
                                # 檢查檔案格式
                                try:
                                    with Image.open(screenshot_path) as img:
                                        img.verify()
                                except Exception as e:
                                    print(f"錯誤：截圖檔案格式無效: {screenshot_path}, 錯誤: {str(e)}")
                                    raise ValueError(f"截圖檔案格式無效: {screenshot_path}")
                                
                                attachments.append(screenshot_path)
                                print(f"成功添加截圖附件: {screenshot_path}, 檔案大小: {file_size} bytes")
                                
                            except Exception as e:
                                print(f"處理截圖檔案時發生錯誤: {str(e)}")
                                # 繼續執行，不中斷郵件發送
                        else:
                            print("警告：未提供截圖路徑")
                        
                        # 發送郵件
                        gmail_sender = GmailSender(
                            sender_email=notification_data["gmail_sender"],
                            app_password=notification_data["gmail_password"]
                        )
                        
                        # 確保所有參數都是字串類型
                        subject = str("預約成功通知")
                        text_content = str(text_content)
                        html_content = str(html_content)
                        
                        success = gmail_sender.send_email(
                            recipient_emails=notification_data["recipient_emails"],
                            subject=subject,
                            text_content=text_content,
                            html_content=html_content,
                            image_paths=attachments,
                            sender_name="預約系統通知"
                        )
                        
                        if success:
                            print(f"成功發送郵件通知，附加檔案：{attachments}")
                        else:
                            print("郵件發送失敗")
                        
                        # 清理截圖檔案
                        for path in attachments:
                            try:
                                os.remove(path)
                                print(f"已刪除暫存檔案：{path}")
                            except Exception as e:
                                print(f"刪除檔案失敗：{path}, 錯誤：{str(e)}")
                        
                    except Exception as notify_error:
                        print(f"發送成功通知失敗: {str(notify_error)}")
            else:
                fail_msg = "預約失敗"
                print(fail_msg)
                if notifier:  # 只有在有通知器的情況下才發送通知
                    try:
                        notifier.send_notification(fail_msg)
                    except Exception as notify_error:
                        print(f"發送失敗通知失敗: {str(notify_error)}")
        except Exception as e:
            error_msg = f"預約過程中發生錯誤: {str(e)}"
            print(error_msg)
            if notifier:  # 只有在有通知器的情況下才發送通知
                try:
                    notifier.send_notification(error_msg)
                except Exception as notify_error:
                    print(f"發送錯誤通知失敗: {str(notify_error)}")
        finally:
            if hasattr(system, "driver"):
                try:
                    print("關閉瀏覽器...")
                    system.driver.quit()
                    print("瀏覽器已關閉")
                except Exception as quit_error:
                    print(f"關閉瀏覽器時發生錯誤: {str(quit_error)}")
    except Exception as main_error:
        print(f"主程序發生嚴重錯誤: {str(main_error)}")


if __name__ == "__main__":
    main()
