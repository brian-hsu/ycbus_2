import os
import re

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
from datetime import datetime
import datetime
import requests
from read_google_sheet import ReadGSheet


def read_txt_to_dict(file_name):
    # 開啟文件並讀取內容
    with open(file_name, "r", encoding="utf-8") as file:
        lines = file.readlines()

    result = {}
    for line in lines:
        # 分割每行的資訊並加入字典 用':'當作介質
        key, value = line.strip().split(":")
        result[key] = value

    return result


file_name = "data.txt"
data = read_txt_to_dict(file_name)

LINE_NOTIFY_TOKEN = data["line_token"]


def line_notify(msg, png_path=None):
    url = "https://notify-api.line.me/api/notify"
    headers = {"Authorization": "Bearer " + LINE_NOTIFY_TOKEN}  # 設定權杖
    data = {"message": msg}  # 設定要發送的訊息
    if png_path:
        image = open(png_path, "rb")  # 以二進位方式開啟圖片
        image_file = {"imageFile": image}  # 設定圖片資訊
    else:
        image_file = None

    # 發送 LINE Notify
    return requests.post(url, headers=headers, data=data, files=image_file, timeout=5)

    # return requests.post(url, headers=headers, data=data)


def gc_load():
    def check_area(area):
        # 新北地區
        area_a = [
            "板橋",
            "新莊",
            "蘆洲",
            "三重",
            "泰山",
            "五股",
            "淡水",
            "樹林",
            "中和",
            "永和",
            "土城",
            "新店",
            "石碇",
            "深坑",
            "烏來",
            "三峽",
            "鶯歌",
            "瑞芳",
            "貢寮",
            "雙溪",
            "平溪",
            "三芝",
            "汐止",
            "坪林",
            "瑞芳",
            "萬里",
            "金山",
            "林口",
            "石門",
            "八里",
        ]
        # 台北市地區
        area_b = [
            "北投",
            "大安",
            "萬華",
            "大同",
            "中山",
            "松山",
            "信義",
            "南港",
            "中正",
            "文山",
            "士林",
            "內湖",
        ]
        if area in area_a:
            return "a"
        elif area in area_b:
            return "b"

    read_gc = ReadGSheet()
    gc_dict = read_gc.gsheet_cover(data["gsheet_cover"])

    user_data = dict()

    user_data["name"] = data["name"]
    user_data["num"] = data["ycbus_password"]

    user_data["Message"] = gc_dict["note_message"]
    user_data["date"] = gc_dict["ride_date"]
    user_data["go_time"] = gc_dict["goto_time"]
    user_data["back_time"] = gc_dict["return_time"]
    # goOn Address
    user_data["go_on_city"] = check_area(gc_dict["goto_pickup_area"])
    user_data["go_on_area"] = gc_dict["goto_pickup_area"]
    user_data["go_on_address"] = gc_dict["goto_pickup_address"]
    # goOff Address
    user_data["go_off_city"] = check_area(gc_dict["goto_dropoff_area"])
    user_data["go_off_area"] = gc_dict["goto_dropoff_area"]
    user_data["go_off_address"] = gc_dict["goto_dropoff_address"]
    # goOn Address
    if gc_dict["return_pickup_area"] == "same_goto_dropoff":
        user_data["back_on_city"] = user_data["go_off_city"]
        user_data["back_on_area"] = user_data["go_off_area"]
    else:
        user_data["back_on_city"] = check_area(gc_dict["return_pickup_area"])
        user_data["back_on_area"] = gc_dict["return_pickup_area"]

    if gc_dict["return_pickup_address"] == "same_goto_dropoff":
        user_data["back_on_address"] = user_data["go_off_address"]
    else:
        user_data["back_on_address"] = gc_dict["return_pickup_address"]

    # goOff Address
    if gc_dict["return_dropoff_area"] == "same_pickup":
        user_data["back_off_city"] = user_data["go_on_city"]
        user_data["back_off_area"] = user_data["go_on_area"]
    else:
        user_data["back_off_city"] = check_area(gc_dict["return_dropoff_area"])
        user_data["back_off_area"] = gc_dict["return_dropoff_area"]

    if gc_dict["return_dropoff_address"] == "same_pickup":
        user_data["back_off_address"] = user_data["go_on_address"]
    else:
        user_data["back_off_address"] = gc_dict["return_dropoff_address"]

    return user_data


def start_count(set_time):
    status = 1
    while status == 1:
        time.sleep(0.33)
        import datetime

        now_time = datetime.datetime.now()

        lock_time = str(set_time)
        try:
            if now_time.strftime("%H:%M") == lock_time:
                print("時間到: %s" % now_time)
                status = 0

            else:
                print("尚未到 %s, 現在時間: [%s]" % (lock_time, now_time))

        except:
            print("function: start_count something wrong")

    print(
        """
    ======================

        啟動 Web driver ...

    ======================
    """
    )


class AutoReserve:
    def __init__(self, my_data, browser_type, headless):
        # User Data
        self.data = my_data
        self.xpath = None

        # CSS Element
        self.css = dict()
        self.css["customerName"] = "input#cusname.w3-large"
        self.css["idCode"] = "input#idcode.w3-large"
        self.css["loginButton"] = "input#btn101.btn"
        self.css["checkNumber"] = "input#chknumber"
        self.css["contentNumber"] = ".w3-center > #chknumber"
        self.css["checkMyID"] = "form#form1 h1:nth-child(4)"
        self.css["backMain"] = 'form#form1 div:nth-child(2) > input[name="btn1"]'

        self.css["sendButton"] = "input#next5"
        print(f"self.data['date'] :: {self.data['date']}")
        self.css["dateButton"] = f"input[type=button][value*='{self.data['date']}']"
        print(f"self.css['dateButton'] :: {self.css['dateButton']}")
        self.css["setGoButton"] = "input#setgom2"
        #                           input[type=radio][onclick*="go_time"]
        self.css["goTimeButton"] = (
            'input[type=radio][onclick*="%s"]' % self.data["go_time"]
        )
        self.css["setBackButton"] = "input#setgon"
        self.css["backTimeButton"] = (
            'input[type=radio][onclick*="%s"]' % self.data["back_time"]
        )
        # goOn Address
        self.css["goOnAreaIn"] = "input[name='areain']"
        self.css["goOnCitySelect"] = "select[name='city']"
        self.css["goOnAreaSelect"] = "select[name='areain_u']"
        self.css["goOnAddress"] = "input[type=text][name='pointin']"
        # goOff Address
        self.css["goOffAreaIn"] = "input[name='areaoff']"
        self.css["goOffCitySelect"] = "select[name='citya']"
        self.css["goOffAreaSelect"] = "select[name='areaoff_u']"
        self.css["goOffAddress"] = "input[type=text][name='pointoff']"
        # Message
        self.css["message"] = "textarea[name='pmark']"
        # backOn Address
        self.css["backOnAreaIn"] = "input[name='areain2']"
        self.css["backOnCitySelect"] = "select[name='cityb']"
        self.css["backOnAreaSelect"] = "select[name='areain2_u']"
        self.css["backOnAddress"] = "input[type=text][name='pointin2']"
        # backOff Address
        self.css["backOffAreaIn"] = "input[name='areaoff2']"
        self.css["backOffCitySelect"] = "select[name='citym']"
        self.css["backOffAreaSelect"] = "select[name='areaoffb_u']"
        self.css["backOffAddress"] = "input[type=text][name='pointoff2']"
        # time clock
        self.css["timeClock"] = "span#time"
        # booking
        self.css["bookingButton"] = "tr#chktmf input#uya"
        self.css["saveButton"] = "#btnSave"
        self.css["viewButton"] = '.btn_grey[name="btn19"]'
        self.css["mainPage"] = "#form1"

        # # webdriver
        # chrome_options = webdriver.ChromeOptions()
        # chrome_options.add_argument('disable-gpu')
        # chrome_options.add_argument('--start-maximized')
        # chrome_options.add_argument('--enable-application-cache')
        # # self.driver = webdriver.Chrome(executable_path=ChromeDriverManager().install())
        # self.driver = webdriver.Chrome(options=chrome_options,
        #                                executable_path=ChromeDriverManager().install())  # 启动时添加定制的选项

        def driver_chrome():
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_argument("--disable-gpu")
            # chrome_options.add_argument('--start-maximized')
            chrome_options.add_argument("--enable-application-cache")
            if headless == 0:
                chrome_options.add_argument("--headless")
            # self.driver = webdriver.Chrome(executable_path=ChromeDriverManager().install())
            # return webdriver.Chrome(options=chrome_options,
            #                                executable_path=ChromeDriverManager().install())  # 启动时添加定制的选项

            return webdriver.Chrome(
                service=ChromeService(
                    ChromeDriverManager(cache_valid_range=28).install()
                ),
                options=chrome_options,
            )

        def driver_firefox():
            from selenium.webdriver.firefox.service import Service as FirefoxService
            from webdriver_manager.firefox import GeckoDriverManager
            from selenium.webdriver.firefox.options import Options

            options = Options()
            if headless == 0:
                options.headless = True
            options.set_preference("javascript.enabled", True)
            # options.set_preference("intl.accept_languages", 'zh-TW, zh')
            options.add_argument("--disable-gpu")
            options.add_argument("lang=zh-TW")
            return webdriver.Firefox(
                service=FirefoxService(
                    GeckoDriverManager(cache_valid_range=28).install()
                ),
                options=options,
            )

        def set_driver():
            if browser_type == "chrome":
                return driver_chrome()
            else:
                return driver_firefox()

        self.driver = set_driver()

        try:
            self.driver.maximize_window()
            self.driver.get("http://rayman.ycbus.org.tw/rayman/book_inq.php")
        except WebDriverException:
            self.driver.quit()
            exit("Cannot navigate to invalid URL !")

    def wait_element(self, element_name, selector="css", seconds=5, freq=0.5):
        i = 0
        while i < 3:
            if selector == "css":
                try:
                    wait_element = WebDriverWait(
                        self.driver, timeout=seconds, poll_frequency=freq
                    ).until(
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR, self.css[element_name])
                        )
                    )
                    return wait_element
                except TimeoutException as timeout:
                    i += 1
                    print(timeout)
                    print("%s is Timeout" % element_name)
                    print("Fail try %s count" % i)
            else:
                try:
                    wait_element = WebDriverWait(self.driver, seconds).until(
                        EC.presence_of_element_located(
                            (By.XPATH, self.xpath[element_name])
                        )
                    )
                    return wait_element
                except TimeoutException as timeout:
                    i += 1
                    print(timeout)
                    print("%s is Timeout" % element_name)
                    print("Fail try %s count" % i)

    def loop_now_time(self, set_lock, debug_flag=0):
        status = 1
        if debug_flag == 1:
            status = 0

        while status == 1:
            time.sleep(1)
            # time_get = self.driver.find_element_by_css_selector(self.css['timeClock']).text
            time_get = self.driver.find_element(
                By.CSS_SELECTOR, self.css["timeClock"]
            ).text
            regex = re.compile(r"\d{4}-\d{2}-\d{2} (\d{2}:\d{2}):(\d{2})")
            search_time = regex.search(time_get)

            lock_time = set_lock

            try:
                if search_time.group(1) == lock_time and int(search_time.group(2)) > 0:
                    print("時間到: %s" % lock_time)
                    status = 0
                # elif debug_flag == 1:
                #     status = 0
                else:
                    print("尚未到 %s, 現在時間: [%s]" % (lock_time, search_time.group(0)))
            except AttributeError:
                print("search_time 可能為 None，請檢查是否正確抓取到時間")
            except IndexError:
                print("search_time 可能未找到所有時間組，請檢查時間格式是否正確")

        self.reserve()

    def check_enter(self):
        try:
            WebDriverWait(self.driver, 0.5).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "form#form1 div > h1:nth-child(1)")
                )
            )
            self.reserve()
        except TimeoutException:
            print("不用 再輸入乘客編號\n")

    def login(self):
        try:
            self.wait_element("customerName").send_keys(self.data["name"])
            self.wait_element("idCode").send_keys(self.data["num"])
            (self.wait_element("loginButton")).click()
        except AttributeError as attrErr:
            print(attrErr)
            exit("The login function error.")

    def wait_load(self):
        wait = 1
        while wait == 1:
            time.sleep(0.1)
            try:
                # self.driver.find_element_by_css_selector(self.css["backMain"])
                self.driver.find_element(By.CSS_SELECTOR, self.css["backMain"])
            except NoSuchElementException:
                time.sleep(0.1)
                wait = 0

    def reserve(self):
        try:
            (self.wait_element("bookingButton", "css")).click()
            # self.wait_load()
            # time.sleep(2)
            self.wait_element("contentNumber", "css", seconds=3, freq=0.1)
            (self.wait_element("checkNumber")).send_keys(self.data["num"])
            (self.wait_element("sendButton")).click()
        except (AttributeError, ElementNotInteractableException) as except_reserve:
            print(except_reserve)
            exit("The reserve function error.")

    def reduce_overflow(self, hh, mm, min):
        hh -= 1
        mm += 60
        new_time = datetime.time(hh, mm - min)
        return new_time.strftime("%H:%M")

    def add_overflow(self, hh, mm, min):
        hh += 1
        mm -= 60
        new_time = datetime.time(hh, mm + min)
        return new_time.strftime("%H:%M")

    def operation_time(self, hh, mm, min, mode):
        hh = int(hh)
        mm = int(mm)
        if mode == "reduce":
            try:
                # new_time = datetime.time(hh, mm - min)

                new_time = datetime.time(hh, mm - min)
                return new_time.strftime("%H:%M")
            except ValueError:
                return str(self.reduce_overflow(hh, mm, min))
        elif mode == "add":
            try:
                add_mins = mm + min
                new_time = datetime.time(hh, add_mins)

                return new_time.strftime("%H:%M")
            except ValueError:
                return str(self.add_overflow(hh, mm, min))

    def reset_time(self, mytime, mode):
        h, m = mytime.split(":")
        return self.operation_time(h, m, 15, mode)

    def check_has_car(self, get_car_time, time_data_name):
        print("== 檢查班次 是否有車班 ==")
        # regex = re.compile(r'%s \[(有車班|車班已滿\.排候補)\]' % self.data[time_data_name])
        regex = re.compile(r"%s \[(有車班|車班已滿\.排候補)]" % self.data[time_data_name])
        car_status = regex.search(get_car_time)
        # my_debug_log(car_status.groups())

        if car_status.group(1) != "有車班":
            print(self.data[time_data_name] + "=> 車班已滿")
            return 0
        else:
            print(self.data[time_data_name] + "=> 有車班")
            return 1

    def go_back_check(self, css_path_name, time_data_name, path_method):
        change_title = None
        if time_data_name == "go_time":
            change_title = "去程"
        elif time_data_name == "back_time":
            change_title = "回程"

        # get_time = self.driver.find_element_by_css_selector("table#innerTable > tbody").text
        get_time = self.driver.find_element(
            By.CSS_SELECTOR, "table#innerTable > tbody"
        ).text
        old_time = self.data[time_data_name]
        count = 0
        while count < 4:
            car_result = self.check_has_car(get_time, time_data_name)
            if car_result == 0:
                self.data[time_data_name] = self.reset_time(
                    self.data[time_data_name], path_method
                )
            else:
                self.css[css_path_name] = (
                    'input[type=radio][onclick*="%s"]' % self.data[time_data_name]
                )
                print(
                    "更改[%s] %s => [%s]"
                    % (change_title, old_time, self.data[time_data_name])
                )
                print("=======================\n")
                count = 999
            count += 1

    def save(self):
        try:
            (self.wait_element("saveButton")).click()
        except:
            print("function: save, something wrong")

    def choose(self):
        try:
            (self.wait_element("dateButton")).click()
            (self.wait_element("setGoButton")).click()
            self.go_back_check("goTimeButton", "go_time", "reduce")
            # self.check_enter()
            (self.wait_element("goTimeButton")).click()
            (self.wait_element("setBackButton")).click()
            # self.check_enter()
            self.go_back_check("backTimeButton", "back_time", "add")
            (self.wait_element("backTimeButton")).click()
            (self.wait_element("sendButton")).click()
        finally:
            pass
        # except AttributeError as attrErr:
        #     print(attrErr)
        #     exit("The choose function error.")

    def address(self):
        try:
            self.wait_element("goOnAreaIn").click()
            Select(self.wait_element("goOnCitySelect")).select_by_value(
                self.data["go_on_city"]
            )
            Select(self.wait_element("goOnAreaSelect")).select_by_value(
                self.data["go_on_area"]
            )
            go_on_addr = self.wait_element("goOnAddress")
            go_on_addr.clear()
            go_on_addr.send_keys(self.data["go_on_address"])

            (self.wait_element("goOffAreaIn")).click()
            Select(self.wait_element("goOffCitySelect")).select_by_value(
                self.data["go_off_city"]
            )
            Select(self.wait_element("goOffAreaSelect")).select_by_value(
                self.data["go_off_area"]
            )
            go_off_addr = self.wait_element("goOffAddress")
            go_off_addr.clear()
            go_off_addr.send_keys(self.data["go_off_address"])

            message = self.wait_element("message")
            message.clear()
            message.send_keys(self.data["Message"])

            (self.wait_element("backOnAreaIn")).click()
            Select(self.wait_element("backOnCitySelect")).select_by_value(
                self.data["back_on_city"]
            )
            Select(self.wait_element("backOnAreaSelect")).select_by_value(
                self.data["back_on_area"]
            )
            back_on_addr = self.wait_element("backOnAddress")
            back_on_addr.clear()
            back_on_addr.send_keys(self.data["back_on_address"])

            (self.wait_element("backOffAreaIn")).click()
            Select(self.wait_element("backOffCitySelect")).select_by_value(
                self.data["back_off_city"]
            )
            Select(self.wait_element("backOffAreaSelect")).select_by_value(
                self.data["back_off_area"]
            )
            back_off_addr = self.wait_element("backOffAddress")
            back_off_addr.clear()
            back_off_addr.send_keys(self.data["back_off_address"])
        except AttributeError as attrErr:
            print(attrErr)
            exit("The address function error.")

    def driver_quit(self):
        self.driver.quit()
        exit(0)

    def screen_shot_max_size(self):
        s = self.driver.get_window_size()
        # obtain browser height and width
        w = self.driver.execute_script("return document.body.parentNode.scrollWidth")
        h = self.driver.execute_script("return document.body.parentNode.scrollHeight")
        # set to new window size
        self.driver.set_window_size(w, h)
        # obtain screenshot of page within body tag
        print("sleep 3 sec")
        time.sleep(3)
        now_time = datetime.datetime.now()
        date_time = now_time.strftime("%Y_%m%d_%H%M_%S")
        # self.driver.find_element_by_tag_name('body').screenshot(f"{date_time}_screen_shot.png")
        self.driver.find_element(By.TAG_NAME, "body").screenshot(
            f"{date_time}_screen_shot.png"
        )
        self.driver.set_window_size(s["width"], s["height"])
        return f"{date_time}_screen_shot.png"

    def screen_shot_custom(self):
        s = self.driver.get_window_size()
        # obtain browser height and width
        w = self.driver.execute_script("return document.body.parentNode.scrollWidth")
        h = self.driver.execute_script("return document.body.parentNode.scrollHeight")
        # set to new window size
        self.driver.set_window_size(w, h)
        # obtain screenshot of page within body tag
        print("sleep 3 sec")
        time.sleep(3)
        now_time = datetime.datetime.now()
        date_time = now_time.strftime("%Y_%m%d_%H%M_%S")
        # self.driver.find_element_by_tag_name('body').screenshot(f"{date_time}_screen_shot.png")
        self.driver.find_element(By.CSS_SELECTOR, "#form1").screenshot(
            f"{date_time}_screen_shot.png"
        )
        self.driver.set_window_size(s["width"], s["height"])
        return f"{date_time}_screen_shot.png"

    def main_to_check_page(self):
        try:
            self.wait_element("mainPage")
            (self.wait_element("viewButton")).click()
        finally:
            pass


def run(mod):
    # f_data = data_load()
    f_data = gc_load()
    print(f'## 日期: {f_data["date"]}')
    print(
        "去程:\n"
        f'## 時間: {f_data["go_time"]}\n'
        f'@@ [上車]地址: {f_data["go_on_address"]}\n'
        f'@@ [下車]地址: {f_data["go_off_address"]}\n'
        f"回程:\n"
        f'## 時間: {f_data["back_time"]}\n'
        f'@@ [上車]地址: {f_data["back_on_address"]}\n'
        f'@@ [下車]地址: {f_data["back_off_address"]}\n'
        f'## 留言: {f_data["Message"]}\n'
        f"#### 確認無誤, 5秒後開始!\n"
    )
    time.sleep(5)

    if mod == "desktop":
        start_count("06:58")  # 設定時間 啟動 webdriver
        # yc_bus = AutoReserve(data, browser_type="firefox")
        yc_bus = AutoReserve(f_data, browser_type="chrome", headless=1)

        print("如果瀏覽器上未使用完成請勿關閉此視窗\n")
        yc_bus.login()
        yc_bus.loop_now_time(
            "07:00", debug_flag=0
        )  # debug_flag=1 忽略 7:00 倒數, debug_flag=0 啟用
        # yc_bus.reserve()
        yc_bus.choose()
        yc_bus.address()

        time.sleep(0.5)
        yc_bus.save()
        time.sleep(2)
        yc_bus.main_to_check_page()
        yc_bus.screen_shot_custom()

        print("the %s is finish" % os.path.basename(__file__))

        q = 1
        while q == 1:
            confirm = input('確認使用完瀏覽器,請輸入 "y" Enter離開此程式與瀏覽器: ')
            if confirm == "y":
                q = 0
                yc_bus.driver_quit()
    elif mod == "server":
        yc_bus = AutoReserve(f_data, browser_type="ff", headless=0)
        yc_bus.login()

        yc_bus.loop_now_time(
            "07:00", debug_flag=0
        )  # debug_flag=1 忽略 7:00 倒數, debug_flag=0 啟用
        # yc_bus.reserve() #feature is deprecated
        yc_bus.choose()
        yc_bus.address()
        time.sleep(0.5)
        yc_bus.save()

        time.sleep(2)
        yc_bus.main_to_check_page()
        screen_path = yc_bus.screen_shot_custom()
        line_notify(screen_path, screen_path)

        yc_bus.driver_quit()

    else:
        yc_bus = AutoReserve(f_data, browser_type="ff", headless=1)
        yc_bus.login()

        yc_bus.loop_now_time(
            "07:00", debug_flag=1
        )  # debug_flag=1 忽略 7:00 倒數, debug_flag=0 啟用
        # yc_bus.reserve()
        yc_bus.choose()
        yc_bus.address()
        time.sleep(0.5)
        # yc_bus.save()

        time.sleep(2)
        # yc_bus.main_to_check_page()
        screen_path = yc_bus.screen_shot_custom()
        # line_notify(screen_path, screen_path)

        yc_bus.driver_quit()


if __name__ == "__main__":
    # run("desktop")
    # run("Debug")
    run("server")
