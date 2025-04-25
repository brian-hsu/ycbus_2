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
            # 創建郵件
            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From'] = f"{self.sender_name} <{self.sender_email}>"
            msg['To'] = ", ".join(self.recipient_emails)

            # 添加純文字內容
            if text_content:
                msg.attach(MIMEText(text_content, 'plain', 'utf-8'))

            # 添加 HTML 內容
            if html_content:
                msg.attach(MIMEText(html_content, 'html', 'utf-8'))

            # 添加圖片附件
            if image_paths:
                for image_path in image_paths:
                    if os.path.exists(image_path):
                        with open(image_path, 'rb') as f:
                            img = MIMEImage(f.read())
                            img.add_header('Content-Disposition', 'attachment', filename=os.path.basename(image_path))
                            msg.attach(img)

            # 發送郵件
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(self.sender_email, self.app_password)
                smtp.send_message(msg)
                print("電子郵件通知發送成功")
        except Exception as e:
            print(f"發送電子郵件通知失敗: {str(e)}")
            raise 