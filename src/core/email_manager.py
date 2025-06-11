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
import sys

# 添加utils路径
sys.path.append(os.path.join(os.path.dirname(__file__), '../utils'))
from database import DatabaseManager

class EmailManager:
    """邮件管理器 - 处理本地邮件文件夹"""
    
    def __init__(self, base_path="D:\\学习\\2025\\app\\邮件导出\\收件箱"):
        self.base_path = base_path
        self.old_patent_folder = os.path.join(base_path, "老专利代理审查提醒")
        self.new_patent_folder = os.path.join(base_path, "新专利代理事务提醒")
        # 初始化数据库管理器
        self.db_manager = DatabaseManager()
        
        # 验证基础路径是否存在
        if not os.path.exists(self.base_path):
            print(f"警告: 基础邮件文件夹不存在: {self.base_path}")
        else:
            print(f"邮件基础路径已确认: {self.base_path}")
            # 列出所有子目录
            subdirs = [d for d in os.listdir(self.base_path) if os.path.isdir(os.path.join(self.base_path, d))]
            print(f"发现子目录: {subdirs}")
        
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
        from_addr = email.get('from', '')
        
        print(f"处理邮件: {subject[:100]}...")
        
        # 1. 首先检查发件人是否为info@sptl.com.cn，过滤掉xx@cffex.com.cn的邮件
        if 'info@sptl.com.cn' not in from_addr.lower():
            print(f"跳过邮件：发件人不是info@sptl.com.cn，实际发件人：{from_addr}")
            return None
            
        # 2. 检查邮件内容是否包含指定的关键信息
        content_patterns = [
            r'关于中国国家知识产权局发出的第一次审查意见通知书，请留意该通知书的答复期限为：(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'请参见我方于(\d{4})年(\d{1,2})月(\d{1,2})日报告的《第1次审查意见通知书》.*?答复该通知书的期限是(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'请参见我方于(\d{4})年(\d{1,2})月(\d{1,2})日报告的.*?第1次审查意见通知书.*?答复.*?期限.*?(\d{4})年(\d{1,2})月(\d{1,2})日'
        ]
        
        content_match = False
        for pattern in content_patterns:
            if re.search(pattern, content, re.DOTALL | re.IGNORECASE):
                content_match = True
                print(f"邮件内容匹配模式：{pattern[:50]}...")
                break
                
        if not content_match:
            print(f"跳过邮件：内容不匹配指定模式")
            return None
        
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
            r'专利申请.*?([\d\.]+)',
            # 新增格式7 - 从内容中提取申请号
            r'申请号[：:]?\s*([\d\.]+)',
            r'国家申请号[：:]?\s*([\d\.]+)'
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
        
        # 提取答复期限 - 增强匹配能力
        deadline_patterns = [
            # 从用户指定的内容模式中提取期限
            r'答复期限为：(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'答复该通知书的期限是(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'答复期限[：:]?\s*(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'期限[：:]?\s*(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'(\d{4})年(\d{1,2})月(\d{1,2})日.*?期限',
            r'(\d{4})年(\d{1,2})月(\d{1,2})日.*?答复',
            r'截止日期[：:]?\s*(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'截止日期[：:]?\s*(\d{4})-(\d{1,2})-(\d{1,2})',
            r'截止[：:]?\s*(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'截止[：:]?\s*(\d{4})-(\d{1,2})-(\d{1,2})',
            # 新增：从用户指定的具体内容中提取
            r'(\d{4})年(\d{1,2})月(\d{1,2})日。详细内容请见附件',
            r'期限是(\d{4})年(\d{1,2})月(\d{1,2})日。我方希望能尽快收到'
        ]
        
        deadline = None
        for pattern in deadline_patterns:
            match = re.search(pattern, content)
            if match:
                year, month, day = match.groups()
                try:
                    deadline = datetime(int(year), int(month), int(day))
                    try:
                        print(f"找到答复期限: {deadline.strftime('%Y-%m-%d')}")
                    except UnicodeEncodeError:
                        print(f"找到答复期限: {deadline.isoformat()}")
                    break
                except ValueError as e:
                    print(f"日期格式错误: {year}-{month}-{day}, 错误: {e}")
                    continue
        
        if not deadline:
            print(f"未找到答复期限，邮件将被忽略: {subject[:50]}...")
            return None
        
        # 计算紧急程度
        today = datetime.now()
        days_left = (deadline - today).days
        
        if days_left < 0:
            urgency_level = 'overdue'  # 逾期 - 红色
            urgency_text = '已逾期'
        elif days_left <= 7:
            urgency_level = 'urgent'   # 紧急 - 黄色
            urgency_text = '紧急'
        else:
            urgency_level = 'normal'   # 正常 - 米粉色
            urgency_text = '正常'
        
        return {
            'application_no': application_no,
            'client_no': client_no,
            'our_no': our_no,
            'deadline': deadline,
            'days_left': days_left,
            'urgency_level': urgency_level,
            'urgency_text': urgency_text,
            'subject': subject,
            'content': content[:500] + '...' if len(content) > 500 else content,
            'file_path': email.get('file_path', ''),
            'from': email.get('from', ''),
            'date': email.get('date', ''),
            'completed': False  # 默认未完成
        }
    
    def get_patent_examination_reminders(self, include_completed=False):
        """获取专利审查临期提醒"""
        print(f"\n开始获取专利审查临期提醒邮件...")
        print(f"扫描基础路径: {self.base_path}")
        
        if not os.path.exists(self.base_path):
            print(f"警告: 基础邮件文件夹路径不存在: {self.base_path}")
            return []
        
        # 扫描整个收件箱目录及其所有子目录
        all_emails = []
        
        # 首先尝试老专利代理审查提醒文件夹
        if os.path.exists(self.old_patent_folder):
            print(f"扫描专门文件夹: {self.old_patent_folder}")
            emails = self.get_emails_from_folder(self.old_patent_folder)
            all_emails.extend(emails)
            print(f"从专门文件夹找到 {len(emails)} 封邮件")
        
        # 然后扫描整个收件箱的所有子目录
        for root, dirs, files in os.walk(self.base_path):
            # 跳过已经处理过的老专利代理审查提醒文件夹
            if root == self.old_patent_folder:
                continue
                
            eml_files = [f for f in files if f.lower().endswith('.eml')]
            if eml_files:
                print(f"扫描目录: {root} (找到 {len(eml_files)} 个.eml文件)")
                folder_emails = self.get_emails_from_folder(root)
                all_emails.extend(folder_emails)
        
        print(f"总共找到 {len(all_emails)} 封邮件")
        
        reminders = []
        
        for i, email in enumerate(all_emails):
            print(f"\n处理第 {i+1}/{len(all_emails)} 封邮件...")
            reminder_info = self.extract_patent_reminder_info(email)
            if reminder_info:
                # 检查数据库中的完成状态
                is_completed = self.db_manager.is_email_completed(
                    reminder_info['application_no'], 
                    reminder_info['file_path']
                )
                reminder_info['completed'] = is_completed
                
                # 根据include_completed参数决定是否包含已完成的邮件
                if include_completed or not is_completed:
                    reminders.append(reminder_info)
                    try:
                        print(f"成功提取专利审查提醒信息: 申请号={reminder_info['application_no']}, 期限={reminder_info['deadline'].strftime('%Y-%m-%d')}, 完成状态={is_completed}")
                    except UnicodeEncodeError:
                        print(f"成功提取专利审查提醒信息: 申请号={reminder_info['application_no']}, 期限={reminder_info['deadline'].isoformat()}, 完成状态={is_completed}")
                else:
                    print(f"跳过已完成的邮件: 申请号={reminder_info['application_no']}")
        
        print(f"\n总共找到 {len(reminders)} 条专利审查临期提醒")
        
        # 按期限排序（由近及远）
        reminders.sort(key=lambda x: x['deadline'])
        
        return reminders
    
    def mark_reminder_completed(self, application_no, file_path, subject, deadline):
        """标记提醒为已完成"""
        return self.db_manager.mark_email_completed(application_no, file_path, subject, deadline)
    
    def mark_reminder_uncompleted(self, application_no, file_path):
        """标记提醒为未完成"""
        return self.db_manager.mark_email_uncompleted(application_no, file_path)
    
    def get_completion_stats(self):
        """获取完成状态统计"""
        return self.db_manager.get_completion_stats()
    
    def get_all_categories(self):
        """获取所有分类的邮件数据"""
        return {
            'patent_examination_reminders': self.get_patent_examination_reminders(),
            'patent_certificates': [],  # 待实现
            'patent_invoices': [],      # 待实现
            'software_notices': []      # 待实现
        }