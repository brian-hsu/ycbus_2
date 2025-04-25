#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Gmail寄件模組
用於透過Gmail帳號發送包含文字和圖片的電子郵件
"""

import os
import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from typing import List, Optional, Union


class GmailSender:
    """Gmail寄件器類別，用於發送電子郵件"""

    def __init__(self, sender_email: str, app_password: str):
        """
        初始化Gmail寄件器
        
        參數:
            sender_email: 寄件者Gmail信箱
            app_password: Gmail應用程式密碼 (不是Gmail的登入密碼)
                          可在 Google帳號 > 安全性 > 應用程式密碼 中取得
        """
        self.sender_email = sender_email
        self.app_password = app_password
        self.smtp_server = "smtp.gmail.com"
        self.port = 587  # Gmail的TLS連接埠

    def send_email(
        self,
        recipient_emails: Union[str, List[str]],
        subject: str,
        text_content: str,
        html_content: Optional[str] = None,
        image_paths: Optional[List[str]] = None,
        sender_name: Optional[str] = None,
        cc_emails: Optional[Union[str, List[str]]] = None,
        bcc_emails: Optional[Union[str, List[str]]] = None,
    ) -> bool:
        """
        發送電子郵件
        
        參數:
            recipient_emails: 收件者信箱或多個收件者信箱的列表
            subject: 郵件主旨
            text_content: 純文字內容
            html_content: HTML格式內容 (可選)
            image_paths: 圖片檔案路徑列表 (可選)
            sender_name: 寄件者名稱 (可選)
            cc_emails: 副本收件者信箱或多個副本收件者信箱的列表 (可選)
            bcc_emails: 密件副本收件者信箱或多個密件副本收件者信箱的列表 (可選)
            
        返回:
            成功發送返回True，失敗返回False
        """
        # 轉換單一信箱字串為列表格式
        if isinstance(recipient_emails, str):
            recipient_emails = [recipient_emails]
        if isinstance(cc_emails, str):
            cc_emails = [cc_emails]
        if isinstance(bcc_emails, str):
            bcc_emails = [bcc_emails]

        # 創建郵件
        message = MIMEMultipart("alternative")
        message["Subject"] = subject
        
        # 設定寄件者格式 (可包含名稱)
        if sender_name:
            message["From"] = formataddr((sender_name, self.sender_email))
        else:
            message["From"] = self.sender_email
            
        message["To"] = ", ".join(recipient_emails)
        
        # 添加副本與密件副本
        if cc_emails:
            message["Cc"] = ", ".join(cc_emails)
        if bcc_emails:
            message["Bcc"] = ", ".join(bcc_emails)

        # 添加純文字內容
        message.attach(MIMEText(text_content, "plain", "utf-8"))
        
        # 添加HTML內容 (如有提供)
        if html_content:
            message.attach(MIMEText(html_content, "html", "utf-8"))

        # 添加圖片附件
        if image_paths:
            for image_path in image_paths:
                self._attach_image(message, image_path)

        try:
            # 建立安全連線
            context = ssl.create_default_context()
            
            # 連接到Gmail SMTP伺服器
            with smtplib.SMTP(self.smtp_server, self.port) as server:
                server.starttls(context=context)  # 啟用TLS加密
                server.login(self.sender_email, self.app_password)
                
                # 所有收件者 (主要收件者、副本、密件副本)
                all_recipients = []
                all_recipients.extend(recipient_emails)
                if cc_emails:
                    all_recipients.extend(cc_emails)
                if bcc_emails:
                    all_recipients.extend(bcc_emails)
                    
                # 發送郵件
                server.sendmail(
                    self.sender_email,
                    all_recipients,
                    message.as_string()
                )
            return True
        except Exception as e:
            print(f"寄送郵件時發生錯誤: {e}")
            return False

    def _attach_image(self, message: MIMEMultipart, image_path: str) -> None:
        """
        添加圖片附件到郵件中
        
        參數:
            message: 郵件物件
            image_path: 圖片檔案路徑
        """
        if not os.path.exists(image_path):
            print(f"警告: 找不到圖片檔案 {image_path}")
            return
            
        # 取得檔案名稱
        filename = os.path.basename(image_path)
        
        # 開啟並讀取圖片檔案
        with open(image_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            
        # 將二進位檔案編碼為可傳輸格式
        encoders.encode_base64(part)
        
        # 新增檔案標頭資訊
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )
        
        # 添加圖片到郵件中
        message.attach(part)