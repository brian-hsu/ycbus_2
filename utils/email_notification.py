#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
邮件通知模块
用于预约系统的电子邮件通知功能
"""

from gmail_sender import GmailSender
from typing import Optional, List
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from datetime import datetime


class EmailNotifier:
    """邮件通知类，用于发送预约系统的通知邮件"""

    def __init__(self, sender_email: str, app_password: str, recipient_emails: List[str], 
                 sender_name: str = "预约系统通知"):
        """
        初始化邮件通知器
        
        参数:
            sender_email: 发件人Gmail信箱
            app_password: Gmail应用程序密码
            recipient_emails: 收件人邮箱列表
            sender_name: 发件人显示名称
        """
        self.sender_email = sender_email
        self.app_password = app_password
        self.recipient_emails = recipient_emails
        self.sender_name = sender_name

    def send_notification(self, subject="预约系统通知", text_content="", html_content="", image_paths=None):
        """發送電子郵件通知
        
        Args:
            subject (str): 郵件主題
            text_content (str): 純文字內容
            html_content (str): HTML 內容
            image_paths (list): 圖片附件路徑列表
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
            msg['From'] = f"{self.sender_name} <{self.sender_email}>"
            msg['To'] = ", ".join(self.recipient_emails)

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
        except Exception as e:
            error_msg = f"發送電子郵件通知失敗: {str(e)}"
            print(error_msg)
            # 可以考慮將錯誤記錄到日誌檔案
            self._log_error(error_msg)
            raise

    def _log_error(self, error_message):
        """記錄錯誤到日誌檔案"""
        try:
            log_dir = "logs"
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)
            
            log_file = os.path.join(log_dir, "email_errors.log")
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(f"[{timestamp}] {error_message}\n")
        except Exception as e:
            print(f"記錄錯誤時發生問題: {str(e)}") 