# 基本設定
DEFAULT_TIMEOUT = 5
DEFAULT_POLL_FREQUENCY = 0.5
MAX_RETRIES = 3
SCREENSHOT_WAIT_TIME = 3

# 瀏覽器設定
BROWSER_OPTIONS = {
    "firefox": {
        "preferences": {
            "javascript.enabled": True
        },
        "arguments": [
            "--disable-gpu",
            "--disable-extensions", 
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--lang=zh-TW"
        ]
    },
    "chrome": {
        "arguments": [
            "--disable-gpu",
            "--enable-application-cache"
        ]
    }
}

# 網站相關設定
BASE_URL = "http://rayman.ycbus.org.tw/rayman/book_inq.php"

# CSS選擇器映射
CSS_SELECTORS = {
    "customerName": "input#cusname.w3-large",
    "idCode": "input#idcode.w3-large",
    "loginButton": "input#btn101.btn",
    # ... 其他選擇器
}

# 地區對應
AREA_MAPPINGS = {
    "新北": [
        "板橋", "新莊", "蘆洲", "三重", "泰山",
        # ... 其他新北地區
    ],
    "台北": [
        "北投", "大安", "萬華", "大同", "中山",
        # ... 其他台北地區
    ]
} 