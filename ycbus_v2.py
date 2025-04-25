import os
import re
import logging
from typing import Optional, Dict, Any
from dataclasses import dataclass
from config import *
from retrying import retry

import time

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from webdriver_manager.firefox import GeckoDriverManager
from datetime import datetime
import datetime
import requests
from read_google_sheet import ReadGSheet
from selenium.webdriver.remote.webelement import WebElement

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

@dataclass
class BookingData:
    """預約資料的資料類別"""
    name: str
    num: str
    date: str
    go_time: str
    back_time: str
    goto_pickup_area: str
    goto_dropoff_area: str
    goto_pickup_address: str
    goto_dropoff_address: str
    return_pickup_area: str
    return_dropoff_area: str
    return_pickup_address: str
    return_dropoff_address: str
    Message: str

class BusBookingSystem:
    def __init__(self, booking_data, browser_type="firefox", options=None):
        """初始化預約系統
        
        Args:
            booking_data: 預約資料物件
            browser_type: 瀏覽器類型，預設為Firefox
            options: 瀏覽器選項設定
        """
        self.logger = logging.getLogger(self.__class__.__name__)
        self.booking_data = booking_data
        self.browser_type = browser_type
        self.options = options
        self.driver = self._initialize_driver()
        self.wait = WebDriverWait(
            self.driver, 
            DEFAULT_TIMEOUT, 
            poll_frequency=DEFAULT_POLL_FREQUENCY
        )
        # 導航到登入頁面
        self.navigate_to_login_page()

    def _initialize_driver(self):
        """初始化並返回WebDriver實例"""
        if self.browser_type.lower() == "firefox":
            if self.options:
                return webdriver.Firefox(options=self.options)
            else:
                return webdriver.Firefox()
        elif self.browser_type.lower() == "chrome":
            return webdriver.Chrome()
        else:
            raise ValueError(f"不支援的瀏覽器類型: {self.browser_type}")

    def wait_for_element(self, selector: str, timeout: int = DEFAULT_TIMEOUT) -> Optional[WebElement]:
        """等待元素出現並返回"""
        try:
            # 檢查是否是直接的 CSS 選擇器
            if selector.startswith('#') or selector.startswith('.') or selector.startswith('[') or selector.find(' ') > 0 or selector.find('>') > 0:
                # 先檢查元素是否存在
                element = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                
                # 然後確保元素可見且可點擊
                try:
                    self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                    )
                    # 如果元素存在但不可見，則嘗試使用JavaScript使其可見
                    if not element.is_displayed():
                        self.logger.warning(f"元素存在但不可見，嘗試使用JavaScript使其可見: {selector}")
                        self.driver.execute_script("arguments[0].style.display = 'block';", element)
                except:
                    self.logger.warning(f"元素不可點擊，可能會影響操作: {selector}")
                
                return element
            else:
                # 嘗試從 CSS_SELECTORS 字典中查找
                try:
                    from config import CSS_SELECTORS
                    css_selector = CSS_SELECTORS.get(selector)
                    if css_selector:
                        # 先檢查元素是否存在
                        element = self.wait.until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, css_selector))
                        )
                        
                        # 然後確保元素可見且可點擊
                        try:
                            self.wait.until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, css_selector))
                            )
                            # 如果元素存在但不可見，則嘗試使用JavaScript使其可見
                            if not element.is_displayed():
                                self.logger.warning(f"元素存在但不可見，嘗試使用JavaScript使其可見: {css_selector}")
                                self.driver.execute_script("arguments[0].style.display = 'block';", element)
                        except:
                            self.logger.warning(f"元素不可點擊，可能會影響操作: {css_selector}")
                        
                        return element
                    else:
                        # 如果在字典中找不到，則直接使用提供的選擇器
                        self.logger.warning(f"在 CSS_SELECTORS 中找不到 {selector}，嘗試直接使用")
                        
                        # 先檢查元素是否存在
                        element = self.wait.until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                        )
                        
                        # 然後確保元素可見且可點擊
                        try:
                            self.wait.until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                            )
                            # 如果元素存在但不可見，則嘗試使用JavaScript使其可見
                            if not element.is_displayed():
                                self.logger.warning(f"元素存在但不可見，嘗試使用JavaScript使其可見: {selector}")
                                self.driver.execute_script("arguments[0].style.display = 'block';", element)
                        except:
                            self.logger.warning(f"元素不可點擊，可能會影響操作: {selector}")
                        
                        return element
                except ImportError:
                    # 如果無法導入 CSS_SELECTORS，則直接使用提供的選擇器
                    self.logger.warning("無法導入 CSS_SELECTORS，直接使用提供的選擇器")
                    
                    # 先檢查元素是否存在
                    element = self.wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    
                    # 然後確保元素可見且可點擊
                    try:
                        self.wait.until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                        )
                        # 如果元素存在但不可見，則嘗試使用JavaScript使其可見
                        if not element.is_displayed():
                            self.logger.warning(f"元素存在但不可見，嘗試使用JavaScript使其可見: {selector}")
                            self.driver.execute_script("arguments[0].style.display = 'block';", element)
                    except:
                        self.logger.warning(f"元素不可點擊，可能會影響操作: {selector}")
                    
                    return element
        except TimeoutException:
            self.logger.warning(f"等待元素超時: {selector}")
            return None

    @retry(stop_max_attempt_number=MAX_RETRIES)
    def login(self, captcha_code: str) -> bool:
        """
        執行登入程序
        Args:
            captcha_code: 驗證碼字串
        Returns:
            bool: 登入是否成功
        """
        try:
            # 等待並填寫帳戶名稱
            username_field = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#cusname"))
            )
            username_field.clear()
            username_field.send_keys(self.booking_data.name)
            self.logger.info("已填寫帳戶名稱")
            
            # 等待並填寫乘客編號
            password_field = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#idcode"))
            )
            password_field.clear()
            password_field.send_keys(self.booking_data.num)
            self.logger.info("已填寫乘客編號")
            
            # 等待並填寫驗證碼
            captcha_field = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#captcha"))
            )
            captcha_field.clear()
            captcha_field.send_keys(captcha_code)
            self.logger.info(f"已填寫驗證碼: {captcha_code}")
            
            # 點擊登入按鈕
            login_button = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#btn101"))
            )
            login_button.click()
            self.logger.info("已點擊登入按鈕")
            
            # 等待並檢查是否登入成功（修改判斷邏輯）
            time.sleep(1)  # 等待頁面反應
            
            # 檢查是否存在錯誤訊息
            try:
                error_message = self.driver.find_element(By.CSS_SELECTOR, ".alert-danger")
                if error_message.is_displayed():
                    self.logger.warning(f"登入失敗 - 出現錯誤訊息: {error_message.text}")
                    return False
            except NoSuchElementException:
                pass
            
            # 檢查是否成功進入系統
            try:
                # 檢查是否存在登出按鈕或其他登入成功後才會出現的元素
                self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div[align='center'].w3-large"))
                )
                current_url = self.driver.current_url
                if "netbook/book.php" in current_url:  # 假設登入成功後會跳轉到這個URL
                    self.logger.info("登入成功")
                    return True
                else:
                    self.logger.warning("登入失敗 - URL未改變")
                    return False
            except TimeoutException:
                self.logger.warning("登入失敗 - 未找到登入後的頁面元素")
                return False
            
        except Exception as e:
            self.logger.error(f"登入過程發生錯誤: {str(e)}")
            return False

    def book_journey(self) -> bool:
        """執行完整預約流程"""
        try:
            # 登入已經在 main.py 中的 handle_login_process 函數中處理
            # 直接進行預約流程
            self.logger.info("開始預約流程...")
            
            # 點擊"預約訂車"按鈕
            try:
                self.logger.info("點擊預約訂車按鈕...")
                book_button = self.wait_for_element("input[value='預約訂車'][onclick*='act.value=\\'netbook\\';snt()']")
                if not book_button:
                    self.logger.error("找不到預約訂車按鈕")
                    return False, None
                book_button.click()
                self.logger.info("已點擊預約訂車按鈕")
                # 等待頁面加載
                time.sleep(2)
            except Exception as e:
                self.logger.error(f"點擊預約訂車按鈕失敗: {str(e)}")
                return False, None
            
            if not self.select_journey_details():
                self.logger.error("選擇行程詳情失敗")
                return False, None
            
            if not self.fill_address_details():
                self.logger.error("填寫地址詳情失敗")
                return False, None
            
            self.logger.info("保存預約...")
            # self.save_booking()
            self.logger.info("預約成功！")
            
            # 在點擊存檔按鈕前截圖
            screenshot_path = None
            try:
                self.logger.info("在點擊存檔按鈕前截取表單畫面...")
                # 獲取視窗大小
                original_size = self.driver.get_window_size()
                
                # 設置視窗大小為頁面大小
                width = self.driver.execute_script("return document.body.parentNode.scrollWidth")
                height = self.driver.execute_script("return document.body.parentNode.scrollHeight")
                self.driver.set_window_size(width, height)
                
                # 等待頁面加載
                time.sleep(2)
                
                # 生成時間戳記檔名
                now_time = datetime.datetime.now()
                date_time = now_time.strftime("%Y_%m%d_%H%M_%S")
                screenshot_path = os.path.abspath(f"form_{date_time}.png")
                
                # 截取表單區域
                form_element = self.driver.find_element(By.CSS_SELECTOR, "#form1")
                form_element.screenshot(screenshot_path)
                
                # 恢復原始視窗大小
                self.driver.set_window_size(original_size["width"], original_size["height"])
                
                self.logger.info(f"存檔前表單畫面已保存至: {screenshot_path}")
            except Exception as e:
                self.logger.warning(f"存檔前截取表單畫面失敗: {str(e)}")
                # 截圖失敗不影響後續操作，繼續執行
            
            # 點擊存檔按鈕
            try:
                self.logger.info("點擊存檔按鈕...")
                save_button = self.wait_for_element("input#btnSave")
                if not save_button:
                    self.logger.error("找不到存檔按鈕")
                    return False, None
                
                # 確保按鈕可見並可點擊
                if not save_button.is_displayed():
                    self.logger.warning("存檔按鈕不可見，嘗試使其可見")
                    self.driver.execute_script("arguments[0].style.display = 'block';", save_button)
                
                try:
                    # save_button.click()
                    self.logger.info("已點擊存檔按鈕")
                except Exception as click_error:
                    self.logger.error(f"點擊存檔按鈕失敗: {str(click_error)}")
                    try:
                        self.logger.info("嘗試使用JavaScript點擊存檔按鈕")
                        self.driver.execute_script("arguments[0].click();", save_button)
                        self.logger.info("已使用JavaScript點擊存檔按鈕")
                    except Exception as js_error:
                        self.logger.error(f"使用JavaScript點擊存檔按鈕失敗: {str(js_error)}")
                        return False, None
            except Exception as e:
                self.logger.error(f"點擊存檔按鈕過程中出錯: {str(e)}")
                return False, None
            
            self.logger.info("地址詳情填寫完成")
            return True, screenshot_path
        except Exception as e:
            self.logger.error(f"預約失敗: {str(e)}")
            return False, None

    def _setup_firefox_driver(self, headless: bool) -> webdriver:
        """設置 Firefox 瀏覽器驅動"""
        try:
            options = FirefoxOptions()
            if headless:
                options.add_argument('--headless')
            
            service = FirefoxService(GeckoDriverManager().install())
            driver = webdriver.Firefox(service=service, options=options)
            driver.maximize_window()
            driver.get(BASE_URL)  # 假設你有定義 BASE_URL
            return driver
        except Exception as e:
            self.logger.error(f"Failed to setup Firefox driver: {str(e)}")
            raise

    def _setup_chrome_driver(self, headless: bool) -> webdriver:
        """設置 Chrome 瀏覽器驅動"""
        try:
            options = webdriver.ChromeOptions()
            if headless:
                options.add_argument('--headless')
            
            service = ChromeService(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
            driver.maximize_window()
            driver.get(BASE_URL)  # 假設你有定義 BASE_URL
            return driver
        except Exception as e:
            self.logger.error(f"Failed to setup Chrome driver: {str(e)}")
            raise

    def select_journey_details(self) -> bool:
        """選擇行程細節"""
        try:
            self.logger.info("選擇行程日期和時間...")
            
            # 點擊日期按鈕 - 使用日期值來定位
            date_value = self.booking_data.date  # 例如 "2023/05/20"
            self.logger.info(f"選擇日期: {date_value}")
            date_button = self.wait_for_element(f"input[value*='{date_value}']")
            if not date_button:
                self.logger.error(f"找不到日期按鈕: {date_value}")
                return False
            date_button.click()
            
            # 點擊去程按鈕
            go_button = self.wait_for_element("input#setgom2")
            if not go_button:
                self.logger.error("找不到去程按鈕")
                return False
            go_button.click()
            
            # 選擇去程時間
            go_time_value = self.booking_data.go_time  # 例如 "08:00"
            self.logger.info(f"選擇去程時間: {go_time_value}")
            
            # 修改選擇器以匹配實際的HTML結構
            # 尋找包含指定時間的單選按鈕，使用 jump.value 而非 go_time
            go_time_button = self.wait_for_element(f"input[type='radio'][onclick*='jump.value'][onclick*='{go_time_value}']")
            if not go_time_button:
                # 嘗試使用更寬鬆的選擇器
                self.logger.warning(f"使用第一種選擇器找不到去程時間按鈕: {go_time_value}，嘗試更寬鬆的選擇器")
                go_time_button = self.wait_for_element(f"input[onclick*='jump.value=\"{go_time_value}\"']")
                
                if not go_time_button:
                    # 嘗試使用ID選擇器，假設ID格式為radioBtn加數字
                    self.logger.warning(f"使用第二種選擇器找不到去程時間按鈕: {go_time_value}，嘗試查找所有時間按鈕")
                    # 獲取所有radio按鈕
                    radio_buttons = self.driver.find_elements(By.CSS_SELECTOR, "input[type='radio'][onclick*='jump.value']")
                    for button in radio_buttons:
                        onclick_attr = button.get_attribute("onclick")
                        if go_time_value in onclick_attr:
                            go_time_button = button
                            break
            
            if not go_time_button:
                self.logger.error(f"找不到去程時間按鈕: {go_time_value}")
                return False
            
            go_time_button.click()
            
            # 點擊回程按鈕
            back_button = self.wait_for_element("input#setgon")
            if not back_button:
                self.logger.error("找不到回程按鈕")
                return False
            back_button.click()
            
            # 選擇回程時間
            back_time_value = self.booking_data.back_time  # 例如 "17:00"
            self.logger.info(f"選擇回程時間: {back_time_value}")
            
            # 修改選擇器以匹配實際的HTML結構
            # 尋找包含指定時間的單選按鈕，使用 jump.value 而非 back_time
            back_time_button = self.wait_for_element(f"input[type='radio'][onclick*='jump.value'][onclick*='{back_time_value}']")
            if not back_time_button:
                # 嘗試使用更寬鬆的選擇器
                self.logger.warning(f"使用第一種選擇器找不到回程時間按鈕: {back_time_value}，嘗試更寬鬆的選擇器")
                back_time_button = self.wait_for_element(f"input[onclick*='jump.value=\"{back_time_value}\"']")
                
                if not back_time_button:
                    # 嘗試使用ID選擇器，假設ID格式為radioBtn加數字
                    self.logger.warning(f"使用第二種選擇器找不到回程時間按鈕: {back_time_value}，嘗試查找所有時間按鈕")
                    # 獲取所有radio按鈕
                    radio_buttons = self.driver.find_elements(By.CSS_SELECTOR, "input[type='radio'][onclick*='jump.value']")
                    for button in radio_buttons:
                        onclick_attr = button.get_attribute("onclick")
                        if back_time_value in onclick_attr:
                            back_time_button = button
                            break
            
            if not back_time_button:
                self.logger.error(f"找不到回程時間按鈕: {back_time_value}")
                return False
            
            back_time_button.click()
            
            # 點擊送出按鈕
            send_button = self.wait_for_element("input#next5")
            if not send_button:
                self.logger.error("找不到送出按鈕")
                return False
            send_button.click()
            
            self.logger.info("行程選擇完成")
            return True
        except Exception as e:
            self.logger.error(f"選擇行程詳情失敗: {str(e)}")
            return False

    def fill_address_details(self) -> bool:
        """填寫地址細節"""
        try:
            self.logger.info("填寫地址詳情...")
            
            # 先確保頁面已完全加載完成
            self.logger.info("等待頁面完全加載...")
            time.sleep(2)  # 短暫等待以確保頁面渲染完成
            
            # 去程上車地點
            self.logger.info("填寫去程上車地點...")
            
            # 確保去程上車區域輸入框可用
            go_on_area = None
            try:
                # 多次嘗試不同的定位方式
                locators = [
                    "input[name='areain']",
                    "#areain",
                    "input#areain",
                    "//input[@name='areain']"
                ]
                
                for loc in locators:
                    self.logger.info(f"嘗試定位去程上車區域: {loc}")
                    if loc.startswith("//"):
                        # 如果是 XPath
                        try:
                            go_on_area = self.wait.until(
                                EC.presence_of_element_located((By.XPATH, loc))
                            )
                            if go_on_area and go_on_area.is_displayed():
                                self.logger.info(f"成功找到去程上車區域: {loc}")
                                break
                        except:
                            continue
                    else:
                        # 如果是 CSS 選擇器
                        go_on_area = self.wait_for_element(loc)
                        if go_on_area:
                            self.logger.info(f"成功找到去程上車區域: {loc}")
                            break
                
                if not go_on_area:
                    self.logger.error("無法找到去程上車區域輸入框，嘗試使用JavaScript定位")
                    go_on_area = self.driver.execute_script("return document.querySelector('input[name=\"areain\"]')")
                    if not go_on_area:
                        self.logger.error("找不到去程上車地區按鈕")
                        return False
            except Exception as e:
                self.logger.error(f"定位去程上車區域時發生錯誤: {str(e)}")
                return False
            
            # 確保元素可見並可點擊
            if not go_on_area.is_displayed():
                self.logger.warning("去程上車區域輸入框不可見，嘗試使其可見")
                self.driver.execute_script("arguments[0].style.display = 'block';", go_on_area)
            
            # 點擊前確保元素可點擊
            try:
                go_on_area.click()
                self.logger.info("已成功點擊去程上車區域")
            except Exception as e:
                self.logger.error(f"點擊去程上車區域失敗: {str(e)}")
                try:
                    self.logger.info("嘗試使用JavaScript點擊")
                    self.driver.execute_script("arguments[0].click();", go_on_area)
                except Exception as js_error:
                    self.logger.error(f"使用JavaScript點擊失敗: {str(js_error)}")
                    return False
                
            # 從字段名稱中提取城市和地區，增加更健壯的處理
            try:
                if '_' not in self.booking_data.goto_pickup_area:
                    self.logger.warning(f"去程上車區域格式不正確: {self.booking_data.goto_pickup_area}，應該是'城市_地區'格式")
                    # 嘗試使用預設值或其他方式處理
                    goto_pickup_city = "a"  # 預設使用新北市
                    goto_pickup_area = self.booking_data.goto_pickup_area
                else:
                    goto_pickup_parts = self.booking_data.goto_pickup_area.split('_')
                    goto_pickup_city = goto_pickup_parts[0]
                    goto_pickup_area = goto_pickup_parts[1]
            except Exception as e:
                self.logger.error(f"處理去程上車區域數據時出錯: {str(e)}")
                goto_pickup_city = "a"  # 預設使用新北市
                goto_pickup_area = "三芝"  # 預設區域
            
            # 選擇去程上車城市和地區
            try:
                # 等待城市選擇框出現
                self.logger.info("等待城市選擇框出現")
                go_on_city = None
                try:
                    go_on_city = self.wait_for_element("select[name='city']")
                    if not go_on_city:
                        go_on_city = self.driver.find_element(By.NAME, "city")
                except:
                    try:
                        go_on_city = self.driver.find_element(By.NAME, "city")
                    except:
                        self.logger.error("找不到城市選擇框")
                        return False
                
                # 確保選擇框可見
                if not go_on_city.is_displayed():
                    self.logger.warning("城市選擇框不可見，嘗試使其可見")
                    self.driver.execute_script("arguments[0].style.display = 'block';", go_on_city)
                
                self.logger.info(f"選擇城市: {goto_pickup_city}")
                go_on_city_select = Select(go_on_city)
                go_on_city_select.select_by_value(goto_pickup_city)
                
                # 等待加載地區選擇框
                time.sleep(1)
                
                # 選擇地區
                self.logger.info(f"選擇地區: {goto_pickup_area}")
                go_on_area_select = None
                try:
                    go_on_area_select = self.wait_for_element("select[name='areain_u']")
                    if not go_on_area_select:
                        go_on_area_select = self.driver.find_element(By.NAME, "areain_u")
                except:
                    try:
                        go_on_area_select = self.driver.find_element(By.NAME, "areain_u")
                    except:
                        self.logger.error("找不到地區選擇框")
                        return False
                
                # 確保選擇框可見
                if not go_on_area_select.is_displayed():
                    self.logger.warning("地區選擇框不可見，嘗試使其可見")
                    self.driver.execute_script("arguments[0].style.display = 'block';", go_on_area_select)
                
                go_on_area_select_obj = Select(go_on_area_select)
                
                # 檢查選擇框是否有該地區選項
                options_text = [o.text for o in go_on_area_select_obj.options]
                if goto_pickup_area in options_text:
                    go_on_area_select_obj.select_by_visible_text(goto_pickup_area)
                elif options_text:
                    self.logger.warning(f"找不到指定地區 {goto_pickup_area}，使用第一個可用選項: {options_text[0]}")
                    go_on_area_select_obj.select_by_visible_text(options_text[0])
                else:
                    self.logger.error("地區選擇框沒有選項")
                    return False
            except Exception as e:
                self.logger.error(f"選擇城市和地區時出錯: {str(e)}")
                # 試圖直接設置 areain 的值
                try:
                    self.logger.info("嘗試直接設置去程上車區域")
                    go_on_area = self.wait_for_element("input[name='areain']")
                    if go_on_area:
                        go_on_area.clear()
                        go_on_area.send_keys(goto_pickup_area)
                    else:
                        self.logger.error("無法找到去程上車區域輸入框")
                        return False
                except Exception as direct_error:
                    self.logger.error(f"直接設置去程上車區域失敗: {str(direct_error)}")
                    return False
            
            # 填寫去程上車地址
            try:
                go_on_address = self.wait_for_element("input[name='pointin']")
                if not go_on_address:
                    self.logger.error("找不到去程上車地址輸入框")
                    return False
                
                go_on_address.clear()
                go_on_address.send_keys(self.booking_data.goto_pickup_address)
                self.logger.info(f"已填寫去程上車地址: {self.booking_data.goto_pickup_address}")
            except Exception as e:
                self.logger.error(f"填寫去程上車地址失敗: {str(e)}")
                return False
            
            # 去程下車地點
            self.logger.info("填寫去程下車地點...")
            # ... 與去程上車地點類似的實現方式 ...
            
            # 確保去程下車區域輸入框可用
            go_off_area = None
            try:
                # 多次嘗試不同的定位方式
                locators = [
                    "input[name='areaoff']",
                    "#areaoff",
                    "input#areaoff",
                    "//input[@name='areaoff']"
                ]
                
                for loc in locators:
                    self.logger.info(f"嘗試定位去程下車區域: {loc}")
                    if loc.startswith("//"):
                        # 如果是 XPath
                        try:
                            go_off_area = self.wait.until(
                                EC.presence_of_element_located((By.XPATH, loc))
                            )
                            if go_off_area and go_off_area.is_displayed():
                                self.logger.info(f"成功找到去程下車區域: {loc}")
                                break
                        except:
                            continue
                    else:
                        # 如果是 CSS 選擇器
                        go_off_area = self.wait_for_element(loc)
                        if go_off_area:
                            self.logger.info(f"成功找到去程下車區域: {loc}")
                            break
                
                if not go_off_area:
                    self.logger.error("無法找到去程下車區域輸入框，嘗試使用JavaScript定位")
                    go_off_area = self.driver.execute_script("return document.querySelector('input[name=\"areaoff\"]')")
                    if not go_off_area:
                        self.logger.error("找不到去程下車地區按鈕")
                        return False
            except Exception as e:
                self.logger.error(f"定位去程下車區域時發生錯誤: {str(e)}")
                return False
            
            # 確保元素可見並可點擊
            if not go_off_area.is_displayed():
                self.logger.warning("去程下車區域輸入框不可見，嘗試使其可見")
                self.driver.execute_script("arguments[0].style.display = 'block';", go_off_area)
            
            # 點擊前確保元素可點擊
            try:
                go_off_area.click()
                self.logger.info("已成功點擊去程下車區域")
            except Exception as e:
                self.logger.error(f"點擊去程下車區域失敗: {str(e)}")
                try:
                    self.logger.info("嘗試使用JavaScript點擊")
                    self.driver.execute_script("arguments[0].click();", go_off_area)
                except Exception as js_error:
                    self.logger.error(f"使用JavaScript點擊失敗: {str(js_error)}")
                    return False
            
            # 從字段名稱中提取城市和地區，增加更健壯的處理
            try:
                if '_' not in self.booking_data.goto_dropoff_area:
                    self.logger.warning(f"去程下車區域格式不正確: {self.booking_data.goto_dropoff_area}，應該是'城市_地區'格式")
                    # 嘗試使用預設值或其他方式處理
                    goto_dropoff_city = "a"  # 預設使用新北市
                    goto_dropoff_area = self.booking_data.goto_dropoff_area
                else:
                    goto_dropoff_parts = self.booking_data.goto_dropoff_area.split('_')
                    goto_dropoff_city = goto_dropoff_parts[0]
                    goto_dropoff_area = goto_dropoff_parts[1]
            except Exception as e:
                self.logger.error(f"處理去程下車區域數據時出錯: {str(e)}")
                goto_dropoff_city = "a"  # 預設使用新北市
                goto_dropoff_area = "淡水"  # 預設區域
            
            # 選擇去程下車城市和地區
            try:
                # 等待城市選擇框出現
                self.logger.info("等待去程下車城市選擇框出現")
                go_off_city = None
                try:
                    go_off_city = self.wait_for_element("select[name='citya']")
                    if not go_off_city:
                        go_off_city = self.driver.find_element(By.NAME, "citya")
                except:
                    try:
                        go_off_city = self.driver.find_element(By.NAME, "citya")
                    except:
                        self.logger.error("找不到去程下車城市選擇框")
                        return False
                
                # 確保選擇框可見
                if not go_off_city.is_displayed():
                    self.logger.warning("去程下車城市選擇框不可見，嘗試使其可見")
                    self.driver.execute_script("arguments[0].style.display = 'block';", go_off_city)
                
                self.logger.info(f"選擇去程下車城市: {goto_dropoff_city}")
                go_off_city_select = Select(go_off_city)
                go_off_city_select.select_by_value(goto_dropoff_city)
                
                # 等待加載地區選擇框
                time.sleep(1)
                
                # 選擇地區
                self.logger.info(f"選擇去程下車地區: {goto_dropoff_area}")
                go_off_area_select = None
                try:
                    go_off_area_select = self.wait_for_element("select[name='areaoff_u']")
                    if not go_off_area_select:
                        go_off_area_select = self.driver.find_element(By.NAME, "areaoff_u")
                except:
                    try:
                        go_off_area_select = self.driver.find_element(By.NAME, "areaoff_u")
                    except:
                        self.logger.error("找不到去程下車地區選擇框")
                        return False
                
                # 確保選擇框可見
                if not go_off_area_select.is_displayed():
                    self.logger.warning("去程下車地區選擇框不可見，嘗試使其可見")
                    self.driver.execute_script("arguments[0].style.display = 'block';", go_off_area_select)
                
                go_off_area_select_obj = Select(go_off_area_select)
                
                # 檢查選擇框是否有該地區選項
                options_text = [o.text for o in go_off_area_select_obj.options]
                if goto_dropoff_area in options_text:
                    go_off_area_select_obj.select_by_visible_text(goto_dropoff_area)
                elif options_text:
                    self.logger.warning(f"找不到指定地區 {goto_dropoff_area}，使用第一個可用選項: {options_text[0]}")
                    go_off_area_select_obj.select_by_visible_text(options_text[0])
                else:
                    self.logger.error("去程下車地區選擇框沒有選項")
                    return False
            except Exception as e:
                self.logger.error(f"選擇去程下車城市和地區時出錯: {str(e)}")
                # 試圖直接設置 areaoff 的值
                try:
                    self.logger.info("嘗試直接設置去程下車區域")
                    go_off_area = self.wait_for_element("input[name='areaoff']")
                    if go_off_area:
                        go_off_area.clear()
                        go_off_area.send_keys(goto_dropoff_area)
                    else:
                        self.logger.error("無法找到去程下車區域輸入框")
                        return False
                except Exception as direct_error:
                    self.logger.error(f"直接設置去程下車區域失敗: {str(direct_error)}")
                    return False
            
            # 填寫去程下車地址
            try:
                go_off_address = self.wait_for_element("input[name='pointoff']")
                if not go_off_address:
                    self.logger.error("找不到去程下車地址輸入框")
                    return False
                
                go_off_address.clear()
                go_off_address.send_keys(self.booking_data.goto_dropoff_address)
                self.logger.info(f"已填寫去程下車地址: {self.booking_data.goto_dropoff_address}")
            except Exception as e:
                self.logger.error(f"填寫去程下車地址失敗: {str(e)}")
                return False
            
            # 填寫留言給客服（如果有）
            try:
                if hasattr(self.booking_data, 'Message') and self.booking_data.Message:
                    message_box = self.wait_for_element("textarea[name='pmark']")
                    if message_box:
                        message_box.clear()
                        message_box.send_keys(self.booking_data.Message)
                        self.logger.info("已填寫留言給客服")
            except Exception as e:
                self.logger.warning(f"填寫留言給客服失敗: {str(e)}")
                # 這不是必須的，所以繼續執行
            
            # 回程上車地點
            self.logger.info("填寫回程上車地點...")
            # 確保回程上車區域輸入框可用
            back_on_area = None
            try:
                # 多次嘗試不同的定位方式
                locators = [
                    "input[name='areain2']",
                    "#areain2",
                    "input#areain2",
                    "//input[@name='areain2']"
                ]
                
                for loc in locators:
                    self.logger.info(f"嘗試定位回程上車區域: {loc}")
                    if loc.startswith("//"):
                        # 如果是 XPath
                        try:
                            back_on_area = self.wait.until(
                                EC.presence_of_element_located((By.XPATH, loc))
                            )
                            if back_on_area and back_on_area.is_displayed():
                                self.logger.info(f"成功找到回程上車區域: {loc}")
                                break
                        except:
                            continue
                    else:
                        # 如果是 CSS 選擇器
                        back_on_area = self.wait_for_element(loc)
                        if back_on_area:
                            self.logger.info(f"成功找到回程上車區域: {loc}")
                            break
                
                if not back_on_area:
                    self.logger.error("無法找到回程上車區域輸入框，嘗試使用JavaScript定位")
                    back_on_area = self.driver.execute_script("return document.querySelector('input[name=\"areain2\"]')")
                    if not back_on_area:
                        self.logger.error("找不到回程上車地區按鈕")
                        return False
            except Exception as e:
                self.logger.error(f"定位回程上車區域時發生錯誤: {str(e)}")
                return False
            
            # 確保元素可見並可點擊
            if not back_on_area.is_displayed():
                self.logger.warning("回程上車區域輸入框不可見，嘗試使其可見")
                self.driver.execute_script("arguments[0].style.display = 'block';", back_on_area)
            
            # 點擊前確保元素可點擊
            try:
                back_on_area.click()
                self.logger.info("已成功點擊回程上車區域")
            except Exception as e:
                self.logger.error(f"點擊回程上車區域失敗: {str(e)}")
                try:
                    self.logger.info("嘗試使用JavaScript點擊")
                    self.driver.execute_script("arguments[0].click();", back_on_area)
                except Exception as js_error:
                    self.logger.error(f"使用JavaScript點擊失敗: {str(js_error)}")
                    return False
            
            # 從字段名稱中提取城市和地區，增加更健壯的處理
            try:
                if '_' not in self.booking_data.return_pickup_area:
                    self.logger.warning(f"回程上車區域格式不正確: {self.booking_data.return_pickup_area}，應該是'城市_地區'格式")
                    # 嘗試使用預設值或其他方式處理
                    return_pickup_city = "a"  # 預設使用新北市
                    return_pickup_area = self.booking_data.return_pickup_area
                else:
                    return_pickup_parts = self.booking_data.return_pickup_area.split('_')
                    return_pickup_city = return_pickup_parts[0]
                    return_pickup_area = return_pickup_parts[1]
            except Exception as e:
                self.logger.error(f"處理回程上車區域數據時出錯: {str(e)}")
                return_pickup_city = "a"  # 預設使用新北市
                return_pickup_area = "淡水"  # 預設區域
            
            # 選擇回程上車城市和地區
            try:
                # 等待城市選擇框出現
                self.logger.info("等待回程上車城市選擇框出現")
                back_on_city = None
                try:
                    back_on_city = self.wait_for_element("select[name='cityb']")
                    if not back_on_city:
                        back_on_city = self.driver.find_element(By.NAME, "cityb")
                except:
                    try:
                        back_on_city = self.driver.find_element(By.NAME, "cityb")
                    except:
                        self.logger.error("找不到回程上車城市選擇框")
                        return False
                
                # 確保選擇框可見
                if not back_on_city.is_displayed():
                    self.logger.warning("回程上車城市選擇框不可見，嘗試使其可見")
                    self.driver.execute_script("arguments[0].style.display = 'block';", back_on_city)
                
                self.logger.info(f"選擇回程上車城市: {return_pickup_city}")
                back_on_city_select = Select(back_on_city)
                back_on_city_select.select_by_value(return_pickup_city)
                
                # 等待加載地區選擇框
                time.sleep(1)
                
                # 選擇地區
                self.logger.info(f"選擇回程上車地區: {return_pickup_area}")
                back_on_area_select = None
                try:
                    back_on_area_select = self.wait_for_element("select[name='areain2_u']")
                    if not back_on_area_select:
                        back_on_area_select = self.driver.find_element(By.NAME, "areain2_u")
                except:
                    try:
                        back_on_area_select = self.driver.find_element(By.NAME, "areain2_u")
                    except:
                        self.logger.error("找不到回程上車地區選擇框")
                        return False
                
                # 確保選擇框可見
                if not back_on_area_select.is_displayed():
                    self.logger.warning("回程上車地區選擇框不可見，嘗試使其可見")
                    self.driver.execute_script("arguments[0].style.display = 'block';", back_on_area_select)
                
                back_on_area_select_obj = Select(back_on_area_select)
                
                # 檢查選擇框是否有該地區選項
                options_text = [o.text for o in back_on_area_select_obj.options]
                if return_pickup_area in options_text:
                    back_on_area_select_obj.select_by_visible_text(return_pickup_area)
                elif options_text:
                    self.logger.warning(f"找不到指定地區 {return_pickup_area}，使用第一個可用選項: {options_text[0]}")
                    back_on_area_select_obj.select_by_visible_text(options_text[0])
                else:
                    self.logger.error("回程上車地區選擇框沒有選項")
                    return False
            except Exception as e:
                self.logger.error(f"選擇回程上車城市和地區時出錯: {str(e)}")
                # 試圖直接設置 areain2 的值
                try:
                    self.logger.info("嘗試直接設置回程上車區域")
                    back_on_area = self.wait_for_element("input[name='areain2']")
                    if back_on_area:
                        back_on_area.clear()
                        back_on_area.send_keys(return_pickup_area)
                    else:
                        self.logger.error("無法找到回程上車區域輸入框")
                        return False
                except Exception as direct_error:
                    self.logger.error(f"直接設置回程上車區域失敗: {str(direct_error)}")
                    return False
            
            # 填寫回程上車地址
            try:
                back_on_address = self.wait_for_element("input[name='pointin2']")
                if not back_on_address:
                    self.logger.error("找不到回程上車地址輸入框")
                    return False
                
                back_on_address.clear()
                back_on_address.send_keys(self.booking_data.return_pickup_address)
                self.logger.info(f"已填寫回程上車地址: {self.booking_data.return_pickup_address}")
            except Exception as e:
                self.logger.error(f"填寫回程上車地址失敗: {str(e)}")
                return False
            
            # 回程下車地點
            self.logger.info("填寫回程下車地點...")
            # 確保回程下車區域輸入框可用
            back_off_area = None
            try:
                # 多次嘗試不同的定位方式
                locators = [
                    "input[name='areaoff2']",
                    "#areaoff2",
                    "input#areaoff2",
                    "//input[@name='areaoff2']"
                ]
                
                for loc in locators:
                    self.logger.info(f"嘗試定位回程下車區域: {loc}")
                    if loc.startswith("//"):
                        # 如果是 XPath
                        try:
                            back_off_area = self.wait.until(
                                EC.presence_of_element_located((By.XPATH, loc))
                            )
                            if back_off_area and back_off_area.is_displayed():
                                self.logger.info(f"成功找到回程下車區域: {loc}")
                                break
                        except:
                            continue
                    else:
                        # 如果是 CSS 選擇器
                        back_off_area = self.wait_for_element(loc)
                        if back_off_area:
                            self.logger.info(f"成功找到回程下車區域: {loc}")
                            break
                
                if not back_off_area:
                    self.logger.error("無法找到回程下車區域輸入框，嘗試使用JavaScript定位")
                    back_off_area = self.driver.execute_script("return document.querySelector('input[name=\"areaoff2\"]')")
                    if not back_off_area:
                        self.logger.error("找不到回程下車地區按鈕")
                        return False
            except Exception as e:
                self.logger.error(f"定位回程下車區域時發生錯誤: {str(e)}")
                return False
            
            # 確保元素可見並可點擊
            if not back_off_area.is_displayed():
                self.logger.warning("回程下車區域輸入框不可見，嘗試使其可見")
                self.driver.execute_script("arguments[0].style.display = 'block';", back_off_area)
            
            # 點擊前確保元素可點擊
            try:
                back_off_area.click()
                self.logger.info("已成功點擊回程下車區域")
            except Exception as e:
                self.logger.error(f"點擊回程下車區域失敗: {str(e)}")
                try:
                    self.logger.info("嘗試使用JavaScript點擊")
                    self.driver.execute_script("arguments[0].click();", back_off_area)
                except Exception as js_error:
                    self.logger.error(f"使用JavaScript點擊失敗: {str(js_error)}")
                    return False
            
            # 從字段名稱中提取城市和地區，增加更健壯的處理
            try:
                if '_' not in self.booking_data.return_dropoff_area:
                    self.logger.warning(f"回程下車區域格式不正確: {self.booking_data.return_dropoff_area}，應該是'城市_地區'格式")
                    # 嘗試使用預設值或其他方式處理
                    return_dropoff_city = "a"  # 預設使用新北市
                    return_dropoff_area = self.booking_data.return_dropoff_area
                else:
                    return_dropoff_parts = self.booking_data.return_dropoff_area.split('_')
                    return_dropoff_city = return_dropoff_parts[0]
                    return_dropoff_area = return_dropoff_parts[1]
            except Exception as e:
                self.logger.error(f"處理回程下車區域數據時出錯: {str(e)}")
                return_dropoff_city = "a"  # 預設使用新北市
                return_dropoff_area = "三芝"  # 預設區域
            
            # 選擇回程下車城市和地區
            try:
                # 等待城市選擇框出現
                self.logger.info("等待回程下車城市選擇框出現")
                back_off_city = None
                try:
                    back_off_city = self.wait_for_element("select[name='citym']")
                    if not back_off_city:
                        back_off_city = self.driver.find_element(By.NAME, "citym")
                except:
                    try:
                        back_off_city = self.driver.find_element(By.NAME, "citym")
                    except:
                        self.logger.error("找不到回程下車城市選擇框")
                        return False
                
                # 確保選擇框可見
                if not back_off_city.is_displayed():
                    self.logger.warning("回程下車城市選擇框不可見，嘗試使其可見")
                    self.driver.execute_script("arguments[0].style.display = 'block';", back_off_city)
                
                self.logger.info(f"選擇回程下車城市: {return_dropoff_city}")
                back_off_city_select = Select(back_off_city)
                back_off_city_select.select_by_value(return_dropoff_city)
                
                # 等待加載地區選擇框
                time.sleep(1)
                
                # 選擇地區
                self.logger.info(f"選擇回程下車地區: {return_dropoff_area}")
                back_off_area_select = None
                try:
                    back_off_area_select = self.wait_for_element("select[name='areaoffb_u']")
                    if not back_off_area_select:
                        back_off_area_select = self.driver.find_element(By.NAME, "areaoffb_u")
                except:
                    try:
                        back_off_area_select = self.driver.find_element(By.NAME, "areaoffb_u")
                    except:
                        self.logger.error("找不到回程下車地區選擇框")
                        return False
                
                # 確保選擇框可見
                if not back_off_area_select.is_displayed():
                    self.logger.warning("回程下車地區選擇框不可見，嘗試使其可見")
                    self.driver.execute_script("arguments[0].style.display = 'block';", back_off_area_select)
                
                back_off_area_select_obj = Select(back_off_area_select)
                
                # 檢查選擇框是否有該地區選項
                options_text = [o.text for o in back_off_area_select_obj.options]
                if return_dropoff_area in options_text:
                    back_off_area_select_obj.select_by_visible_text(return_dropoff_area)
                elif options_text:
                    self.logger.warning(f"找不到指定地區 {return_dropoff_area}，使用第一個可用選項: {options_text[0]}")
                    back_off_area_select_obj.select_by_visible_text(options_text[0])
                else:
                    self.logger.error("回程下車地區選擇框沒有選項")
                    return False
            except Exception as e:
                self.logger.error(f"選擇回程下車城市和地區時出錯: {str(e)}")
                # 試圖直接設置 areaoff2 的值
                try:
                    self.logger.info("嘗試直接設置回程下車區域")
                    back_off_area = self.wait_for_element("input[name='areaoff2']")
                    if back_off_area:
                        back_off_area.clear()
                        back_off_area.send_keys(return_dropoff_area)
                    else:
                        self.logger.error("無法找到回程下車區域輸入框")
                        return False
                except Exception as direct_error:
                    self.logger.error(f"直接設置回程下車區域失敗: {str(direct_error)}")
                    return False
            
            # 填寫回程下車地址
            try:
                back_off_address = self.wait_for_element("input[name='pointoff2']")
                if not back_off_address:
                    self.logger.error("找不到回程下車地址輸入框")
                    return False
                
                back_off_address.clear()
                back_off_address.send_keys(self.booking_data.return_dropoff_address)
                self.logger.info(f"已填寫回程下車地址: {self.booking_data.return_dropoff_address}")
            except Exception as e:
                self.logger.error(f"填寫回程下車地址失敗: {str(e)}")
                return False
            
            self.logger.info("地址詳情填寫完成")
            return True, screenshot_path
        except Exception as e:
            self.logger.error(f"填寫地址詳情失敗: {str(e)}")
            return False, None

    def save_booking(self) -> bool:
        """儲存預約"""
        try:
            self.logger.info("儲存預約...")
            save_button = self.wait_for_element("#btnSave")
            if not save_button:
                self.logger.error("找不到儲存按鈕")
                return False
            
            save_button.click()
            self.logger.info("預約已儲存")
            return True
        except Exception as e:
            self.logger.error(f"儲存預約失敗: {str(e)}")
            return False

    def capture_confirmation(self) -> str:
        """擷取確認畫面"""
        try:
            self.logger.info("導航到確認頁面...")
            # 先導航到主頁面
            main_page_button = self.wait_for_element("input[name='btn1']")
            if main_page_button:
                main_page_button.click()
            
            # 點擊查看預約按鈕
            view_button = self.wait_for_element(".btn_grey[name='btn19']")
            if not view_button:
                self.logger.error("找不到查看預約按鈕")
                return ""
            view_button.click()
            
            self.logger.info("截取確認畫面...")
            # 獲取視窗大小
            original_size = self.driver.get_window_size()
            
            # 設置視窗大小為頁面大小
            width = self.driver.execute_script("return document.body.parentNode.scrollWidth")
            height = self.driver.execute_script("return document.body.parentNode.scrollHeight")
            self.driver.set_window_size(width, height)
            
            # 等待頁面加載
            time.sleep(3)
            
            # 生成時間戳記檔名
            now_time = datetime.datetime.now()
            date_time = now_time.strftime("%Y_%m%d_%H%M_%S")
            screenshot_path = f"{date_time}_screen_shot.png"
            
            # 截取表單區域
            form_element = self.driver.find_element(By.CSS_SELECTOR, "#form1")
            form_element.screenshot(screenshot_path)
            
            # 恢復原始視窗大小
            self.driver.set_window_size(original_size["width"], original_size["height"])
            
            self.logger.info(f"確認畫面已保存至: {screenshot_path}")
            return screenshot_path
        except Exception as e:
            self.logger.error(f"擷取確認畫面失敗: {str(e)}")
            return ""

    def navigate_to_login_page(self):
        """導航到登入頁面"""
        try:
            self.logger.info(f"正在導航到登入頁面: {BASE_URL}")
            self.driver.get(BASE_URL)
            # 等待頁面加載完成
            self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            self.logger.info("成功導航到登入頁面")
            return True
        except Exception as e:
            self.logger.error(f"導航到登入頁面失敗: {str(e)}")
            return False
