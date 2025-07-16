import smtplib
from email.message import EmailMessage
import os
import logging

def send_email_with_attachment(smtp_server, smtp_port, sender_email, sender_password, send_emails: dict, subject, body, attachment_path):
    """
    发送带附件邮件
    :param smtp_server: SMTP服务器地址
    :param smtp_port: SMTP端口，一般465或587
    :param sender_email: 发件人邮箱
    :param sender_password: 发件人邮箱授权码或密码
    :param send_emails: dict，收件人格式 {"hwy": "17613139826@163.com", ...}
    :param subject: 邮件主题
    :param body: 邮件正文
    :param attachment_path: 附件路径
    """

    # 创建邮件对象
    msg = EmailMessage()
    msg["From"] = sender_email
    msg["To"] = ", ".join(send_emails.values())
    msg["Subject"] = subject
    msg.set_content(body)

    # 读取附件
    with open(attachment_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(attachment_path)

    # 添加附件
    msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)

    try:
        # SMTP连接
        if smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
        server.login(sender_email, sender_password)

        server.send_message(msg)
        server.quit()
        print(f"邮件发送成功，收件人：{list(send_emails.keys())}")
        logging.info(f"邮件发送成功，收件人：{list(send_emails.keys())}\n\n")
    except Exception as e:
        print(f"邮件发送失败: {e}")
        logging.info(f"邮件发送失败: {e}\n\n")

