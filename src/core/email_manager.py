import os
import json
import re
import glob
from datetime import datetime
from pathlib import Path
from email import policy
from email.parser import BytesParser
from email.header import decode_header
import chardet

class EmailManager:
    """邮件管理器 - 处理本地邮件文件夹"""
    
    def __init__(self, base_path="D:\\学习\\2025\\app\\邮件导出\\收件箱"):
        self.base_path = base_path
        self.old_patent_folder = os.path.join(base_path, "老专利代理审查提醒")
        self.new_patent_folder = os.path.join(base_path, "新专利代理事务提醒")
        
    def decode_header_value(self, header_value):
        """解码邮件头信息"""
        if not header_value:
            return ""
        
        try:
            decoded_parts = decode_header(header_value)
            result = ""
            for part, encoding in decoded_parts:
                if isinstance(part, bytes):
                    if encoding:
                        result += part.decode(encoding)
                    else:
                        # 尝试检测编码
                        detected = chardet.detect(part)
                        encoding = detected.get('encoding', 'utf-8')
                        result += part.decode(encoding, errors='ignore')
                else:
                    result += str(part)
            return result
        except Exception as e:
            return str(header_value)
    
    def parse_email_file(self, file_path):
        """解析邮件文件"""
        try:
            with open(file_path, 'rb') as f:
                msg = BytesParser(policy=policy.default).parse(f)
            
            # 提取基本信息
            subject = self.decode_header_value(msg.get('Subject', ''))
            from_addr = self.decode_header_value(msg.get('From', ''))
            date_str = msg.get('Date', '')
            
            # 提取邮件内容
            content = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        payload = part.get_payload(decode=True)
                        if payload:
                            try:
                                content += payload.decode('utf-8', errors='ignore')
                            except:
                                content += str(payload)
            else:
                payload = msg.get_payload(decode=True)
                if payload:
                    try:
                        content += payload.decode('utf-8', errors='ignore')
                    except:
                        content += str(payload)
            
            return {
                'file_path': file_path,
                'subject': subject,
                'from': from_addr,
                'date': date_str,
                'content': content
            }
        except Exception as e:
            print(f"解析邮件文件失败 {file_path}: {e}")
            return None
    
    def get_emails_from_folder(self, folder_path):
        """从文件夹获取所有邮件"""
        emails = []
        if not os.path.exists(folder_path):
            return emails
        
        # 支持的邮件文件格式
        email_extensions = ['*.eml', '*.msg', '*.mbox']
        
        for ext in email_extensions:
            pattern = os.path.join(folder_path, '**', ext)
            for file_path in glob.glob(pattern, recursive=True):
                email_data = self.parse_email_file(file_path)
                if email_data:
                    emails.append(email_data)
        
        return emails
    
    def extract_patent_reminder_info(self, email):
        """提取专利审查提醒信息"""
        subject = email.get('subject', '')
        content = email.get('content', '')
        
        print(f"处理邮件: {subject[:100]}...")
        
        # 检查是否为专利审查提醒邮件 - 扩展匹配模式以适应更多格式
        reminder_patterns = [
            # 原有格式
            r'【提醒函】CN\s*申请号：([^;]+);贵方编号：([^;]+);我方编号：([^;]+)',
            # 新增格式1 - 没有贵方编号的情况
            r'【提醒函】CN\s*申请号：([^;]+);我方编号：([^;]+)',
            # 新增格式2 - 简单包含申请号的情况
            r'【提醒函】CN\s*申请号：([^;,，]+)',
            # 新增格式3 - 其他可能的格式
            r'申请号[为：:]\s*([\d\.]+)',
            # 新增格式4 - 邮件标题中包含申请号的情况
            r'申请号[：:](\d+\.\d+)',
            # 新增格式5 - 邮件标题中包含CN和申请号的情况
            r'CN\s*(\d+\.\d+)',
            # 新增格式6 - 邮件标题中包含专利申请的情况
            r'专利申请.*?([\d\.]+)'
        ]
        
        application_no = None
        client_no = ""
        our_no = ""
        
        # 尝试从标题中提取信息
        for pattern in reminder_patterns:
            match = re.search(pattern, subject)
            if match:
                groups = match.groups()
                if len(groups) >= 1:
                    application_no = groups[0].strip()
                if len(groups) >= 2:
                    if '我方编号' in pattern:
                        our_no = groups[1].strip()
                    else:
                        client_no = groups[1].strip()
                if len(groups) >= 3:
                    our_no = groups[2].strip()
                break
        
        # 如果标题中没有找到申请号，尝试从内容中提取
        if not application_no:
            app_no_patterns = [
                r'申请号[为：:]\s*([\d\.]+)',
                r'国家申请号[：:]\s*([\d\.]+)'
            ]
            for pattern in app_no_patterns:
                match = re.search(pattern, content)
                if match:
                    application_no = match.group(1).strip()
                    break
        
        # 如果仍然没有找到申请号，则不是我们要找的邮件
        if not application_no:
            return None
        
        # 尝试从内容中提取我方编号（如果标题中没有）
        if not our_no:
            our_no_patterns = [
                r'我方编号[：:]\s*([\w\-\.]+)',
                r'我方编号：([^\s\n]+)'
            ]
            for pattern in our_no_patterns:
                match = re.search(pattern, content)
                if match:
                    our_no = match.group(1).strip()
                    break
        
        # 尝试从内容中提取客户编号（如果标题中没有）
        if not client_no:
            client_no_patterns = [
                r'贵方编号[：:]\s*([\w\-\.]*)',
                r'客户编号[：:]\s*([\w\-\.]*)',
            ]
            for pattern in client_no_patterns:
                match = re.search(pattern, content)
                if match:
                    client_no = match.group(1).strip()
                    break
        
        # 提取答复期限
        deadline_patterns = [
            r'答复该通知书的期限是(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'答复期限[：:]?\s*(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'期限[：:]?\s*(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'(\d{4})年(\d{1,2})月(\d{1,2})日.*?期限',
            r'(\d{4})年(\d{1,2})月(\d{1,2})日.*?答复',
            r'截止日期[：:]?\s*(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'截止日期[：:]?\s*(\d{4})-(\d{1,2})-(\d{1,2})',
            r'截止[：:]?\s*(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'截止[：:]?\s*(\d{4})-(\d{1,2})-(\d{1,2})'
        ]
        
        deadline = None
        for pattern in deadline_patterns:
            match = re.search(pattern, content)
            if match:
                year, month, day = match.groups()
                try:
                    deadline = datetime(int(year), int(month), int(day))
                    print(f"找到答复期限: {deadline.strftime('%Y年%m月%d日')}")
                    break
                except ValueError as e:
                    print(f"日期格式错误: {year}-{month}-{day}, 错误: {e}")
                    continue
        
        if not deadline:
            print(f"未找到答复期限，邮件将被忽略: {subject[:50]}...")
            return None
        
        return {
            'application_no': application_no,
            'client_no': client_no,
            'our_no': our_no,
            'deadline': deadline,
            'subject': subject,
            'content': content[:500] + '...' if len(content) > 500 else content,
            'file_path': email.get('file_path', ''),
            'from': email.get('from', ''),
            'date': email.get('date', '')
        }
    
    def get_patent_examination_reminders(self):
        """获取专利审查临期提醒"""
        print(f"\n开始获取专利审查临期提醒邮件...")
        print(f"邮件文件夹路径: {self.old_patent_folder}")
        
        if not os.path.exists(self.old_patent_folder):
            print(f"警告: 邮件文件夹路径不存在: {self.old_patent_folder}")
            return []
            
        emails = self.get_emails_from_folder(self.old_patent_folder)
        print(f"找到 {len(emails)} 封邮件")
        
        reminders = []
        
        for i, email in enumerate(emails):
            print(f"\n处理第 {i+1}/{len(emails)} 封邮件...")
            reminder_info = self.extract_patent_reminder_info(email)
            if reminder_info:
                reminders.append(reminder_info)
                print(f"成功提取专利审查提醒信息: 申请号={reminder_info['application_no']}, 期限={reminder_info['deadline'].strftime('%Y年%m月%d日')}")
        
        print(f"\n总共找到 {len(reminders)} 条专利审查临期提醒")
        
        # 按期限排序（由近及远）
        reminders.sort(key=lambda x: x['deadline'])
        
        return reminders
    
    def get_all_categories(self):
        """获取所有分类的邮件数据"""
        return {
            'patent_examination_reminders': self.get_patent_examination_reminders(),
            'patent_certificates': [],  # 待实现
            'patent_invoices': [],      # 待实现
            'software_notices': []      # 待实现
        }