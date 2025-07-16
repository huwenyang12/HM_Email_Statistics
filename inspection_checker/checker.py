# checker.py
import pandas as pd
from imapclient import IMAPClient
from email.header import decode_header
from datetime import datetime, timedelta
from excel_exporter import export_to_excel
import logging

# 初始化日志配置
logging.basicConfig(
    filename='inspection.log',  # 日志文件
    level=logging.INFO,         # 日志等级
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)


class InspectionChecker:
    def __init__(self, email, password, imap_server, platform_schedules, mailbox_folder):
        self.email = email
        self.password = password
        self.imap_server = imap_server
        self.platform_schedules = platform_schedules
        self.mailbox_folder = mailbox_folder  # 新增指定文件夹

    def decode_subject(self, raw_subject):
        if not raw_subject:
            return '[无主题]'
        decoded_parts = decode_header(raw_subject)
        subject = ''
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                subject += part.decode(encoding or 'utf-8', errors='ignore')
            else:
                subject += part
        return subject

    def in_time_range(self, sched_time_str, email_time, max_minutes):
        sched_time = datetime.strptime(sched_time_str, "%H:%M").time()
        start_dt = email_time.replace(hour=sched_time.hour, minute=sched_time.minute, second=0, microsecond=0)
        end_dt = start_dt + timedelta(minutes=max_minutes)
        return start_dt <= email_time <= end_dt

    def debug_list_emails(self, date_obj):
        """调试打印所有邮件"""
        emails = self.fetch_email_subjects_for_day(date_obj)
        print(f"共获取邮件 {len(emails)} 封：\n")
        for i, (subject, time_obj) in enumerate(emails, 1):
            print(f"[{i}] {time_obj.strftime('%Y-%m-%d %H:%M:%S')} | {subject}")

    def fetch_email_subjects_for_day(self, date_obj):
        with IMAPClient(self.imap_server) as client:
            client.login(self.email, self.password)
            # client.select_folder('INBOX')
            client.select_folder(self.mailbox_folder, readonly=True)  # 使用配置的文件夹名
            since = date_obj.strftime("%d-%b-%Y")
            messages = client.search(['SINCE', since])
            response = client.fetch(messages, ['ENVELOPE'])

            return [
                (self.decode_subject(data[b'ENVELOPE'].subject.decode() if data[b'ENVELOPE'].subject else ''),
                 data[b'ENVELOPE'].date.replace(tzinfo=None))  # 去掉时区信息
                for msgid, data in response.items()
            ]

    def check_schedule_detailed(self, platform, schedule_items, emails, date_obj):
        results, now = [], datetime.now()
        date_str = date_obj.strftime("%Y-%m-%d")
        hit_count, real_check_count = 0, 0
        pending_count = 0

        for sched_time_str, max_minutes in schedule_items:
            sched_time = datetime.strptime(sched_time_str, "%H:%M").time()
            sched_dt = datetime.combine(date_obj.date(), sched_time)
            end_dt = sched_dt + timedelta(minutes=max_minutes)

            # 巡检还未开始
            if now < sched_dt:
                results.append("暂未巡检")
                continue

            matched = False
            for subject, time_obj in emails:
                if platform in subject and '巡检报告' in subject and date_str in subject:
                    if sched_dt <= time_obj <= end_dt:
                        results.append(time_obj.strftime("%H:%M"))
                        hit_count += 1
                        real_check_count += 1
                        matched = True
                        break

            if not matched:
                if now <= end_dt:
                    results.append("巡检中")  # 尚在允许范围内，视为“巡检中”
                    pending_count += 1
                else:
                    results.append("巡检失败")

        effective_total = hit_count + (len([r for r in results if r == "巡检失败"]))
        success_rate = round(hit_count / effective_total * 100) if effective_total else 0

        return results, success_rate, hit_count, effective_total, pending_count


    def generate_headers(self):
        max_times = max(len(times) for times in self.platform_schedules.values())
        headers = ["序号", "巡检平台"]
        for i in range(max_times):
            headers += [f"巡检时间{i+1}", f"巡检结果{i+1}"]
        headers += ["成功率", "备注"]  
        return headers

    def run(self, date_obj):
        date_str = date_obj.strftime("%Y-%m-%d")
        emails = self.fetch_email_subjects_for_day(date_obj)

        headers = self.generate_headers()
        result_data = []
        abnormal_platforms = []

        total_hits = total_checks = total_pending = total_running = 0
        platform_count = len(self.platform_schedules)
        max_times = max(len(v) for v in self.platform_schedules.values())

        output_lines = []  # 收集用于邮件正文的输出内容

        def log_and_collect(msg):
            print(msg)
            output_lines.append(msg)
        logging.info(f"开始巡检任务：{date_str}")
        for idx, (platform, sched_list) in enumerate(self.platform_schedules.items(), start=1):
            results, success_rate, hit, check_total, pending_now = self.check_schedule_detailed(
                platform, sched_list, emails, date_obj
            )
            logging.info(f"{platform} 巡检成功 {hit}/{check_total}，成功率 {success_rate}%")
            row = [idx, platform]
            failed_times = []

            for i, (time_str, _) in enumerate(sched_list):
                result = results[i]
                row += [time_str, result]
                if result == "巡检失败":
                    failed_times.append(time_str)
                elif result == "巡检中":
                    total_running += 1
                elif result == "暂未巡检":
                    total_pending += 1

            while len(row) < 2 + max_times * 2:
                row += ["", ""]

            row.append(f"{success_rate}%")
            row.append("")  # 备注
            result_data.append(row)

            if success_rate < 100:
                abnormal_platforms.append((platform, success_rate, failed_times))

            total_hits += hit
            total_checks += check_total

        # 控制台输出 + 邮件正文构建
        log_and_collect(f"巡检日期：{date_str}")
        log_and_collect(f"巡检概览（共 {platform_count} 份报告，{total_checks} 次应巡检）：")
        log_and_collect(f"成功：{total_hits} 次")
        log_and_collect(f"失败：{total_checks - total_hits} 次")
        log_and_collect(f"巡检中：{total_running} 次")
        log_and_collect(f"暂未巡检：{total_pending} 次")
        total_rate = round(total_hits / total_checks * 100, 1) if total_checks else 0.0
        log_and_collect(f"成功率：{total_rate}%\n")

        log_and_collect("巡检异常平台：")
        for i, (platform, rate, failed_list) in enumerate(abnormal_platforms, start=1):
            fail_str = ", ".join(failed_list)
            logging.warning(f"[异常平台] {platform} 成功率 {rate}%，失败时段：{fail_str}")
            line1 = f"[{i}] {platform}（成功率: {rate}%）"
            line2 = f"    失败时段：{fail_str}"
            line3 = f"    失败原因："
            log_and_collect(line1)
            log_and_collect(line2)
            log_and_collect(line3)

        logging.info(f"总共成功 {total_hits}/{total_checks}，成功率 {total_rate}%")

        file_path = f"巡检邮件统计_{date_str}.xlsx"
        df = pd.DataFrame(result_data, columns=headers)
        export_to_excel(df, headers, file_path)

        log_and_collect(f"\n详细结果已保存至：{file_path}")
        logging.info(f"巡检结果已保存至 Excel 文件：{file_path}")

        return "\n".join(output_lines)  # 邮件正文内容字符串

