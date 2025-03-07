import cv2
import numpy as np
from PIL import Image
import os
import ddddocr
import re
import pytesseract
from PIL import ImageEnhance, ImageFilter

def try_multiple_preprocessing(image_path):
    """嘗試多種預處理方法以提高辨識率"""
    image = cv2.imread(image_path)
    if image is None:
        print(f"無法讀取圖片: {image_path}")
        return []
    
    processed_images = []
    base_path = image_path.replace('.png', '')
    
    # 方法1: 基本二值化
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    path1 = f"{base_path}_binary.png"
    cv2.imwrite(path1, binary)
    processed_images.append(path1)
    
    # 方法2: 調整對比度並二值化
    alpha = 2.0
    beta = 10
    adjusted = cv2.convertScaleAbs(gray, alpha=alpha, beta=beta)
    _, binary2 = cv2.threshold(adjusted, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    path2 = f"{base_path}_contrast.png"
    cv2.imwrite(path2, binary2)
    processed_images.append(path2)
    
    # 方法3: 形態學處理
    kernel = np.ones((2, 2), np.uint8)
    dilated = cv2.dilate(binary, kernel, iterations=1)
    path3 = f"{base_path}_dilated.png"
    cv2.imwrite(path3, dilated)
    processed_images.append(path3)
    
    # 方法4: 自適應二值化
    adaptive_thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                           cv2.THRESH_BINARY, 11, 2)
    path4 = f"{base_path}_adaptive.png"
    cv2.imwrite(path4, adaptive_thresh)
    processed_images.append(path4)
    
    # 方法5: 使用PIL增強 - 修正模式問題
    try:
        pil_image = Image.open(image_path)
        # 確保圖片為RGB模式，若不是則轉換
        if pil_image.mode != 'RGB':
            pil_image = pil_image.convert('RGB')
        enhancer = ImageEnhance.Contrast(pil_image)
        enhanced_img = enhancer.enhance(2.5)
        enhanced_img = enhanced_img.filter(ImageFilter.SHARPEN)
        path5 = f"{base_path}_enhanced.png"
        enhanced_img.save(path5)
        processed_images.append(path5)
    except Exception as e:
        print(f"PIL圖像增強處理錯誤: {str(e)}")
    
    # 方法6: 降噪後二值化
    denoised = cv2.fastNlMeansDenoising(gray, None, 10, 7, 21)
    _, binary_denoised = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    path6 = f"{base_path}_denoised.png"
    cv2.imwrite(path6, binary_denoised)
    processed_images.append(path6)
    
    # 方法7: 圖像分割處理 - 嘗試分割成3個數字
    width = image.shape[1]
    part_width = width // 3
    for i in range(3):
        start_x = i * part_width
        end_x = (i + 1) * part_width if i < 2 else width
        digit_img = gray[:, start_x:end_x]
        _, digit_binary = cv2.threshold(digit_img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        path_digit = f"{base_path}_digit_{i+1}.png"
        cv2.imwrite(path_digit, digit_binary)
        processed_images.append(path_digit)
    
    return processed_images

def recognize_with_multiple_engines(image_path):
    """結合多種OCR引擎嘗試辨識"""
    results = []
    
    # 先使用原始圖片嘗試辨識
    try:
        # 使用ddddocr
        ocr = ddddocr.DdddOcr(show_ad=False)
        with open(image_path, 'rb') as f:
            image_bytes = f.read()
        result_dddd = ocr.classification(image_bytes)
        numbers_dddd = re.findall(r'\d+', result_dddd)
        if numbers_dddd:
            result_dddd = ''.join(numbers_dddd)
            print(f"ddddocr原始辨識結果: {result_dddd}")
            results.append(result_dddd)
        
        # 嘗試使用Tesseract
        try:
            img = Image.open(image_path)
            # 確保圖片格式兼容Tesseract
            if img.mode not in ['RGB', 'L']:
                img = img.convert('L')  # 轉換為灰階
            # 設定Tesseract配置，只辨識數字
            custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789'
            result_tesseract = pytesseract.image_to_string(img, config=custom_config)
            numbers_tesseract = re.findall(r'\d+', result_tesseract)
            if numbers_tesseract:
                result_tesseract = ''.join(numbers_tesseract)
                print(f"Tesseract原始辨識結果: {result_tesseract}")
                results.append(result_tesseract)
        except Exception as e:
            print(f"Tesseract辨識錯誤: {str(e)}")
    
        # 預處理圖片並嘗試辨識
        processed_images = try_multiple_preprocessing(image_path)
        for img_path in processed_images:
            # 確認圖片存在
            if not os.path.exists(img_path):
                print(f"預處理圖片不存在: {img_path}")
                continue
                
            # 使用ddddocr
            try:
                with open(img_path, 'rb') as f:
                    img_bytes = f.read()
                result = ocr.classification(img_bytes)
                numbers = re.findall(r'\d+', result)
                if numbers:
                    result = ''.join(numbers)
                    print(f"ddddocr處理後圖片 {os.path.basename(img_path)} 辨識結果: {result}")
                    results.append(result)
            except Exception as e:
                print(f"處理圖片 {img_path} 時發生錯誤: {str(e)}")
            
            # 使用Tesseract
            try:
                img = Image.open(img_path)
                # 確保圖片格式兼容Tesseract
                if img.mode not in ['RGB', 'L']:
                    img = img.convert('L')  # 轉換為灰階
                result_tess = pytesseract.image_to_string(img, config=custom_config)
                numbers_tess = re.findall(r'\d+', result_tess)
                if numbers_tess:
                    result_tess = ''.join(numbers_tess)
                    print(f"Tesseract處理後圖片 {os.path.basename(img_path)} 辨識結果: {result_tess}")
                    results.append(result_tess)
            except Exception as e:
                print(f"Tesseract處理圖片 {img_path} 錯誤: {str(e)}")
        
        # 處理分割的數字圖片 (最後三個)
        segment_results = []
        for i in range(3):
            digit_path = f"{image_path.replace('.png', '')}_digit_{i+1}.png"
            if os.path.exists(digit_path):
                try:
                    with open(digit_path, 'rb') as f:
                        img_bytes = f.read()
                    digit_result = ocr.classification(img_bytes)
                    numbers = re.findall(r'\d+', digit_result)
                    if numbers:
                        digit = numbers[0]
                        # 確保只取一個數字
                        digit = digit[0] if len(digit) > 0 else ""
                        segment_results.append(digit)
                except Exception as e:
                    print(f"處理分割圖片 {digit_path} 時發生錯誤: {str(e)}")
        
        if len(segment_results) > 0:
            combined = ''.join(segment_results)
            print(f"分割圖片組合辨識結果: {combined}")
            results.append(combined)
    
    except Exception as e:
        print(f"辨識過程出現錯誤: {str(e)}")
        import traceback
        print(traceback.format_exc())  # 輸出完整錯誤追蹤
    
    # 處理所有結果，尋找最佳答案
    final_result = find_best_result(results)
    return final_result

def find_best_result(results):
    """從多個辨識結果中選擇最佳的一個"""
    if not results:
        return None
    
    # 過濾掉空字串
    results = [r for r in results if r]
    
    # 優先選擇3位數字的結果
    three_digits = [r for r in results if len(r) == 3]
    if three_digits:
        # 統計所有3位數字結果，選擇出現頻率最高的
        from collections import Counter
        counts = Counter(three_digits)
        most_common = counts.most_common(1)[0][0]
        return most_common
    
    # 如果沒有3位數字結果，則選擇最接近3位數的結果
    results.sort(key=lambda x: abs(len(x) - 3))
    return results[0]

def cleanup_temp_files(base_image_path):
    """清理所有預處理產生的臨時圖片檔案"""
    base_path = base_image_path.replace('.png', '')
    patterns = [
        f"{base_path}_binary.png",
        f"{base_path}_contrast.png",
        f"{base_path}_dilated.png",
        f"{base_path}_adaptive.png",
        f"{base_path}_enhanced.png",
        f"{base_path}_denoised.png",
        f"{base_path}_digit_1.png",
        f"{base_path}_digit_2.png",
        f"{base_path}_digit_3.png"
    ]
    
    for pattern in patterns:
        try:
            if os.path.exists(pattern):
                os.remove(pattern)
                print(f"已刪除臨時檔案: {os.path.basename(pattern)}")
        except Exception as e:
            print(f"刪除檔案 {pattern} 時發生錯誤: {str(e)}")

def main():
    # 設定工作目錄
    base_dir = os.path.dirname(os.path.abspath(__file__))
    temp_dir = os.path.join(base_dir, "temp_captcha")
    
    # 確保temp_captcha目錄存在
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
    
    # 圖片路徑
    image_path = os.path.join(temp_dir, "captcha_2.png")
    
    # 檢查圖片是否存在
    if not os.path.exists(image_path):
        print(f"請將驗證碼圖片放在此路徑: {image_path}")
        return
    
    try:
        # 執行辨識
        result = recognize_with_multiple_engines(image_path)
        if result:
            print(f"最終辨識結果: {result}")
            if len(result) == 3 and result.isdigit():
                print("成功辨識出3位數字！")
                # 清理臨時檔案
                cleanup_temp_files(image_path)
            else:
                print("警告：辨識結果不是3位數字")
        else:
            print("無法辨識數字")
    except Exception as e:
        print(f"發生錯誤: {str(e)}")

if __name__ == "__main__":
    main()