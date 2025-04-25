#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Gmail寄件模組使用範例
展示如何使用GmailSender類別發送包含文字和圖片的電子郵件
"""

from gmail_sender import GmailSender

def main():
    # 設定寄件者資訊
    # 注意: 請使用您的Gmail帳號及應用程式密碼
    sender_email = "goodog7772@gmail.com"
    app_password = "bnrt nbra diae xnlv"  # 不是Gmail的登入密碼
    
    # 初始化Gmail寄件器
    gmail = GmailSender(sender_email, app_password)
    
    # 設定郵件內容
    recipient_email = ["brian.hsu0702@gmail.com", "goodog7772@gmail.com"]
    subject = "測試郵件 - 包含圖片"
    
    # 純文字內容
    text_content = """
    您好！
    
    這是一封測試郵件，包含文字和圖片附件。
    如果您看不到HTML內容，請使用支援HTML的郵件客戶端查看。
    
    祝您有美好的一天！
    """
    
    # HTML內容 (可呈現更豐富的格式)
    html_content = """
    <html>
      <body>
        <h2 style="color: #2e6c80;">您好！</h2>
        <p>這是一封測試郵件，包含<strong>文字</strong>和<em>圖片</em>附件。</p>
        <p>以下是內嵌的圖片：</p>
        <img src="cid:embedded_image" width="300" />
        <p>此外還有圖片附件。</p>
        <p style="color: #808080; font-style: italic;">祝您有美好的一天！</p>
      </body>
    </html>
    """
    
    # 圖片路徑列表
    image_paths = [
        "temp_captcha/captcha_0.png",  # 請確保此路徑存在
        "confirmation.png"   # 請確保此路徑存在
    ]
    
    # 發送郵件
    success = gmail.send_email(
        recipient_emails=recipient_email,
        subject=subject,
        text_content=text_content,
        html_content=html_content,
        image_paths=image_paths,
        sender_name="Python 測試郵件",  # 顯示的寄件者名稱

    )
    
    if success:
        print("郵件已成功發送！")
    else:
        print("郵件發送失敗。")


if __name__ == "__main__":
    main()