import requests

class LineNotifier:
    def __init__(self, token: str):
        self.token = token
        self.api_url = "https://notify-api.line.me/api/notify"
        
    def send_notification(self, message: str, image_path: str = None):
        """
        發送 Line 通知
        
        Args:
            message: 要發送的訊息
            image_path: 圖片路徑（選擇性）
        """
        headers = {"Authorization": f"Bearer {self.token}"}
        payload = {"message": message}
        files = {"imageFile": open(image_path, "rb")} if image_path else None
        
        try:
            response = requests.post(
                self.api_url,
                headers=headers,
                data=payload,
                files=files
            )
            response.raise_for_status()
        except Exception as e:
            print(f"發送 Line 通知時發生錯誤：{str(e)}")
        finally:
            if files and "imageFile" in files:
                files["imageFile"].close() 