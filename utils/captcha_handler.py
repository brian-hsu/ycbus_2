import pytesseract
from PIL import Image
import os
import time
import requests
from selenium.webdriver.common.by import By
from io import BytesIO
import numpy as np
import cv2
from utils.image_ocr import ImageOCR

class CaptchaHandler:
    def __init__(self, driver):
        self.driver = driver
        # 設定 Tesseract 執行檔路徑
        pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        # 初始化 ImageOCR
        self.image_ocr = ImageOCR()

    def preprocess_image(self, image):
        """
        預處理圖片以提高識別準確率
        """
        # 將 PIL Image 轉換為 numpy array
        img_array = np.array(image)
        
        # 檢查圖片是否已經是灰度圖
        if len(img_array.shape) == 2:
            gray = img_array
        else:
            gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
        
        # 增強對比度
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(4,4))
        gray = clahe.apply(gray)
        
        # 使用自適應閾值進行二值化
        binary = cv2.adaptiveThreshold(
            gray,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY_INV,
            11,
            2
        )
        
        # 去除小噪點
        kernel = np.ones((2,2), np.uint8)
        opening = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)
        
        # 稍微膨脹，使數字更清晰
        dilated = cv2.dilate(opening, kernel, iterations=1)
        
        # 轉回 PIL Image
        return Image.fromarray(dilated)

    def recognize_captcha(self, image_path):
        """
        識別驗證碼
        :param image_path: 驗證碼圖片的本地路徑
        :return: 識別出的驗證碼文字
        """
        try:
            # 檢查圖片是否存在
            if not os.path.exists(image_path):
                print(f"錯誤：驗證碼圖片不存在: {image_path}")
                return None
                
            # 檢查圖片大小
            file_size = os.path.getsize(image_path)
            if file_size < 100:  # 如果圖片太小，可能是下載失敗
                print(f"警告：驗證碼圖片太小 ({file_size} bytes)，可能下載不完整")
                return None
            
            # 使用新的 ImageOCR 類別進行辨識
            print("使用主要辨識方法...")
            result = self.image_ocr.recognize_captcha(image_path)
            
            # 如果新方法失敗，嘗試使用舊方法作為備用
            if not result or len(result) != 3 or not result.isdigit():
                print("主要辨識方法失敗，使用備用辨識方法...")
                # 讀取圖片
                image = Image.open(image_path)
                
                # 嘗試多種預處理方法
                processed_images = []
                
                # 方法1：基本預處理
                processed_image1 = self.preprocess_image(image)
                processed_images.append(processed_image1)
                
                # 方法2：增強對比度
                try:
                    from PIL import ImageEnhance
                    enhancer = ImageEnhance.Contrast(image)
                    enhanced_image = enhancer.enhance(2.0)  # 增強對比度
                    processed_image2 = self.preprocess_image(enhanced_image)
                    processed_images.append(processed_image2)
                except Exception as e:
                    print(f"對比度增強處理失敗: {str(e)}")
                
                # 方法3：銳化
                try:
                    from PIL import ImageFilter
                    sharpened_image = image.filter(ImageFilter.SHARPEN)
                    processed_image3 = self.preprocess_image(sharpened_image)
                    processed_images.append(processed_image3)
                except Exception as e:
                    print(f"銳化處理失敗: {str(e)}")
                
                # 方法4：降噪
                try:
                    import cv2
                    import numpy as np
                    img_array = np.array(image)
                    if len(img_array.shape) == 3:
                        img_array = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
                    denoised = cv2.fastNlMeansDenoising(img_array, None, 10, 7, 21)
                    denoised_image = Image.fromarray(denoised)
                    processed_image4 = self.preprocess_image(denoised_image)
                    processed_images.append(processed_image4)
                except Exception as e:
                    print(f"降噪處理失敗: {str(e)}")
                
                # 保存所有處理後的圖片以便檢查效果
                for i, processed_image in enumerate(processed_images):
                    debug_path = image_path.replace('.png', f'_processed_{i}.png')
                    processed_image.save(debug_path)
                    print(f"已保存處理後的圖片 {i}: {debug_path}")
                
                # 對每個處理後的圖片嘗試識別
                results = []
                for i, processed_image in enumerate(processed_images):
                    # 使用 Tesseract 進行 OCR，調整 PSM 模式
                    for psm in [7, 6, 8, 13]:
                        custom_config = f'--oem 3 --psm {psm} -c tessedit_char_whitelist=0123456789'
                        ocr_result = pytesseract.image_to_string(
                            processed_image, 
                            config=custom_config
                        ).strip()
                        
                        # 只保留數字
                        ocr_result = ''.join(filter(str.isdigit, ocr_result))
                        
                        # 驗證結果是否為3位數
                        if len(ocr_result) == 3:
                            print(f"處理方法 {i}, PSM {psm} 成功識別: {ocr_result}")
                            results.append(ocr_result)
                
                # 如果有多個結果，選擇出現頻率最高的
                if results:
                    from collections import Counter
                    most_common = Counter(results).most_common(1)[0][0]
                    print(f"多種方法中最常見的結果: {most_common}")
                    return most_common
                
                # 如果備用方法也失敗，嘗試使用 ddddocr
                try:
                    import ddddocr
                    ocr = ddddocr.DdddOcr(show_ad=False)
                    with open(image_path, 'rb') as f:
                        img_bytes = f.read()
                    dddd_result = ocr.classification(img_bytes)
                    print(f"ddddocr 識別結果: {dddd_result}")
                    
                    # 只保留數字
                    dddd_result = ''.join(filter(str.isdigit, dddd_result))
                    
                    # 驗證結果是否為3位數
                    if len(dddd_result) == 3:
                        return dddd_result
                except Exception as e:
                    print(f"ddddocr 識別失敗: {str(e)}")
            
            return result
                
        except Exception as e:
            print(f"驗證碼識別出錯: {str(e)}")
            return None 