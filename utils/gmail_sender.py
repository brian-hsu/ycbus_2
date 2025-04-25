#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Gmail 郵件發送模組
用於發送預約系統的通知郵件
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from typing import List, Optional
import os


class GmailSender:
    """Gmail 郵件發送類別"""

    def __init__(self, sender_email: str, app_password: str):
        """
        初始化 Gmail 發送器
        
        Args:
            sender_email: 發件人 Gmail 信箱
            app_password: Gmail 應用程式密碼
        """
        self.sender_email = sender_email
        self.app_password = app_password

    def send_email(
        self,
        recipient_emails: List[str],
        subject: str,
        text_content: str = "",
        html_content: str = "",
        image_paths: Optional[List[str]] = None,
        sender_name: str = "預約系統通知"
    ) -> bool:
        """
        發送電子郵件
        
        Args:
            recipient_emails: 收件人郵箱列表
            subject: 郵件主題
            text_content: 純文字內容
            html_content: HTML 內容
            image_paths: 圖片附件路徑列表
            sender_name: 發件人顯示名稱
            
        Returns:
            bool: 發送是否成功
        """
        try:
            # 參數型別檢查
            if not isinstance(subject, str):
                raise TypeError("郵件主題必須是字串類型")
            if not isinstance(text_content, str):
                raise TypeError("純文字內容必須是字串類型")
            if not isinstance(html_content, str):
                raise TypeError("HTML內容必須是字串類型")
            if image_paths is not None and not isinstance(image_paths, list):
                raise TypeError("圖片路徑必須是列表類型")

            # 創建郵件
            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From'] = f"{sender_name} <{self.sender_email}>"
            msg['To'] = ", ".join(recipient_emails)

            # 添加純文字內容
            if text_content:
                text_part = MIMEText(text_content, 'plain', 'utf-8')
                msg.attach(text_part)

            # 添加 HTML 內容
            if html_content:
                html_part = MIMEText(html_content, 'html', 'utf-8')
                msg.attach(html_part)

            # 添加圖片附件
            if image_paths:
                for image_path in image_paths:
                    if not isinstance(image_path, str):
                        raise TypeError(f"圖片路徑必須是字串類型: {image_path}")
                    if not os.path.exists(image_path):
                        raise FileNotFoundError(f"找不到圖片檔案: {image_path}")
                    
                    try:
                        with open(image_path, 'rb') as f:
                            img_data = f.read()
                            img = MIMEImage(img_data)
                            img.add_header('Content-Disposition', 'attachment', 
                                         filename=os.path.basename(image_path))
                            msg.attach(img)
                    except Exception as e:
                        print(f"處理圖片 {image_path} 時發生錯誤: {str(e)}")
                        continue

            # 發送郵件
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(self.sender_email, self.app_password)
                smtp.send_message(msg)
                print("電子郵件通知發送成功")
                return True
        except Exception as e:
            error_msg = f"發送電子郵件通知失敗: {str(e)}"
            print(error_msg)
            return False 