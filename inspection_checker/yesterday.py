from checker import InspectionChecker
from datetime import datetime, timedelta
from dotenv import load_dotenv
from email_sender import send_email_with_attachment
import os

# 加载 .env 文件中的环境变量
load_dotenv()

# 发送邮件配置
email = os.getenv("EMAIL_ACCOUNT")
password = os.getenv("EMAIL_PASSWORD")
imap_server = os.getenv("IMAP_SERVER")
smtp_server = os.getenv("SMTP_SERVER", "smtp.163.com")
smtp_port = int(os.getenv("SMTP_PORT", 465))
mailbox_folder = os.getenv("MAILBOX_FOLDER", "INBOX")

# 巡检计划
platform_schedules = {
    '网管中心-基础资源': [("07:30", 20), ("11:30", 20), ("15:30", 20), ("19:30", 20)],
    '网管中心-网络安全': [("08:30", 30), ("11:30", 30), ("14:30", 30), ("17:30", 30), ("20:30", 30), ("23:30", 30)],
    '网管中心-北京通业务系统': [("07:30", 10), ("19:30", 10)],
    '权益平台（政务外网）-基础资源': [("08:00", 60), ("12:00", 60), ("16:00", 60), ("20:00", 60)],
    '权益平台（政务外网）-网络安全': [("09:15", 10), ("11:15", 10), ("15:15", 10), ("18:15", 10), ("21:15", 10), ("23:15", 10)],
    '权益平台（政务外网）-系统负载': [("09:30", 30), ("13:30", 30), ("17:30", 30), ("21:30", 30)],
    '权益平台（互联网）-网络安全': [("08:00", 10), ("11:00", 10), ("14:00", 10), ("17:00", 10), ("20:00", 10), ("23:00", 10)],
    '权益平台（互联网）-基础资源': [("08:10", 50), ("12:10", 50), ("16:10", 50), ("20:10", 50)],
    '综合系统（腾讯云）-网络安全': [("08:00", 20), ("12:00", 20), ("16:00", 20), ("20:00", 20)],
    '民生卡业务网络专线': [("08:10", 20), ("12:10", 20), ("16:10", 20), ("20:10", 20)],
    '综合系统（腾讯云）-基础资源': [("08:30", 150), ("12:30", 150), ("16:30", 150), ("20:30", 150)],
}

# 收件人配置
send_emails = {
    "hwy": "17613139826@163.com", # 胡文扬
    # "csz": "csz1009@yeah.net", # 陈思志
}

# 巡检日期配置
# check_date = datetime(2025, 7, 13)  # 指定日期结果
check_date = datetime.now() - timedelta(days=1)   # 昨天巡检结果
# check_date = datetime.now()  # 今天巡检结果


if __name__ == "__main__":
    checker = InspectionChecker(
        email=email,
        password=password,
        imap_server=imap_server,
        platform_schedules=platform_schedules,
        mailbox_folder=mailbox_folder
    )
    # checker.debug_list_emails(datetime.now())  # 用于调试邮箱邮件内容

    body_text = checker.run(check_date)  # 返回指定日期邮件结果

    # 邮件主题和附件路径
    date_str = check_date.strftime('%Y-%m-%d')
    subject = f"巡检邮件统计报告 - {date_str}"
    file_path = f"巡检邮件统计_{date_str}.xlsx"

    send_email_with_attachment(
        smtp_server=smtp_server,
        smtp_port=smtp_port,
        sender_email=email,
        sender_password=password,
        send_emails=send_emails,
        subject=subject,
        body=body_text,  # 邮件正文内容
        attachment_path=file_path
    )
