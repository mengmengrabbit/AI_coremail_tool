import os
import json
import re
import glob
import uuid
from datetime import datetime
from pathlib import Path
from email import policy
from email.parser import BytesParser
from email.header import decode_header
from sre_parse import BRANCH
import chardet
import sys
import shutil

# 添加utils路径
sys.path.append(os.path.join(os.path.dirname(__file__), '../utils'))
from database import DatabaseManager

class EmailManager:
    """邮件管理器 - 处理本地邮件文件夹"""
    
    def __init__(self, base_path="D:\\学习\\2025\\app\\tools\\coremail-connect\\data\\收件箱"):
        self.base_path = base_path
        self.old_patent_folder = os.path.join(base_path, "老专利代理审查提醒")
        self.new_patent_folder = os.path.join(base_path, "新专利代理事务提醒")
        # 初始化数据库管理器
        self.db_manager = DatabaseManager()
        
        # 创建专利授权证书保存目录
        self.certificates_folder = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'static', 'certificates')
        os.makedirs(self.certificates_folder, exist_ok=True)
        
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
            html_content = ""
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    if content_type == "text/plain":
                        payload = part.get_payload(decode=True)
                        if payload:
                            try:
                                content += payload.decode('utf-8', errors='ignore')
                            except:
                                content += str(payload)
                    elif content_type == "text/html":
                        payload = part.get_payload(decode=True)
                        if payload:
                            try:
                                html_content += payload.decode('utf-8', errors='ignore')
                            except:
                                html_content += str(payload)
            else:
                content_type = msg.get_content_type()
                payload = msg.get_payload(decode=True)
                if payload:
                    try:
                        if content_type == "text/html":
                            html_content += payload.decode('utf-8', errors='ignore')
                        else:
                            content += payload.decode('utf-8', errors='ignore')
                    except:
                        content += str(payload)
            
            # 如果纯文本内容为空但HTML内容不为空，使用HTML内容
            if not content.strip() and html_content.strip():
                # 简单去除HTML标签，保留文本内容
                content = re.sub(r'<[^>]+>', ' ', html_content)
                content = re.sub(r'\s+', ' ', content).strip()
            
            # 提取附件信息
            attachments = []
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_disposition() == 'attachment':
                        filename = self.decode_header_value(part.get_filename())
                        if filename:
                            attachments.append({
                                'filename': filename,
                                'content_type': part.get_content_type(),
                                'part': part
                            })
            
            return {
                'file_path': file_path,
                'filename': os.path.basename(file_path),
                'subject': subject,
                'from': from_addr,
                'date': date_str,
                'content': content,
                'html_content': html_content,
                'attachments': attachments
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
        
        # 扫描整个收件箱的所有子目录
        for root, dirs, files in os.walk(self.base_path):
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
            'patent_certificates': self.get_patent_certificates(),  # 实现专利授权证书汇总
            'patent_invoices': [],      # 待实现
            'software_notices': []      # 待实现
        }
        
    def get_patent_certificates(self):
        """获取专利授权证书汇总"""
        print(f"\n开始获取专利授权证书邮件...")
        print(f"扫描基础路径: {self.base_path}")
        
        if not os.path.exists(self.base_path):
            print(f"警告: 基础邮件文件夹路径不存在: {self.base_path}")
            return []
        
        # 扫描整个收件箱目录及其所有子目录
        all_emails = []
        
        # 扫描整个收件箱的所有子目录
        for root, dirs, files in os.walk(self.base_path):
            eml_files = [f for f in files if f.lower().endswith('.eml')]
            if eml_files:
                print(f"扫描目录: {root} (找到 {len(eml_files)} 个.eml文件)")
                folder_emails = self.get_emails_from_folder(root)
                all_emails.extend(folder_emails)
        
        print(f"总共找到 {len(all_emails)} 封邮件")
        
        certificates = []
        
        for i, email in enumerate(all_emails):
            print(f"\n处理第 {i+1}/{len(all_emails)} 封邮件...")
            certificate_info = self.extract_patent_certificate_info(email, self.certificates_folder)
            
            if certificate_info:
                certificates.append(certificate_info)
                print(f"成功提取专利授权证书信息: {certificate_info['filename']}")
        
        print(f"\n总共找到 {len(certificates)} 个专利授权证书")
        
        # 按日期排序（由新到旧）
        certificates.sort(key=lambda x: x.get('date', ''), reverse=True)
        
        return certificates
    
    def extract_patent_certificate_info(self, email, certificates_folder):
        """从邮件中提取专利证书信息"""
        from_addr = email.get('from', '')
        content = email.get('content', '')
        subject = email.get('subject', '')
        date = email.get('date','')
        attachments = email.get('attachments', [])
        
        print(f"处理邮件: {subject[:100]}...")
        print(f"发件人: {from_addr}")
        print(f"邮件内容前100字符: {content[:100]}...")
        print(f"附件数量: {len(attachments)}")
        
        
        
        # 1. 检查邮件内容是否包含指定的关键信息
        # 更灵活的匹配方式，使用多个关键词组合
        keywords_sets = [
            # 组合1：标准表述
            ["专利证书公告", "电子件转给贵方"],
            # 组合2：变体表述1
            ["专利证书", "电子件", "转给"],
            # 组合3：变体表述2
            ["专利", "证书", "公告", "电子件"],
            # 组合4：最小关键词集
            ["专利", "证书"]
        ]
        
        # 清理内容，去除空格和换行符
        clean_content = re.sub(r'\s+', '', content)
        
        # 检查是否匹配任一关键词组合
        content_match = False
        for keywords in keywords_sets:
            if all(keyword in content for keyword in keywords):
                content_match = True
                print(f"邮件内容匹配成功：{keywords}")
                break
        
        if not content_match:
            print(f"跳过邮件：内容不包含指定的关键信息")
            return None
            
        # 2. 检查附件中是否有PDF文件，并且文件名包含"证书"
        certificate_files = []
        for attachment in attachments:
            filename = attachment.get('filename', '')
            content_type = attachment.get('content_type', '')
            
            # 检查是否为PDF文件，并且文件名包含"证书"或者是专利相关文件
            is_pdf = content_type.lower() == 'application/pdf' or filename.lower().endswith('.pdf')
            is_certificate = '证书' in filename or '专利' in filename or '授权' in filename
            
            if is_pdf and is_certificate:
                print(f"找到证书附件: {filename}")
                
                # 保存附件
                part = attachment.get('part')
                if part:
                    try:
                        # 确保目录存在
                        os.makedirs(certificates_folder, exist_ok=True)
                        
                        # 生成唯一文件名
                        unique_filename = f"{uuid.uuid4()}_{filename}"
                        file_path = os.path.join(certificates_folder, unique_filename)
                        
                        # 保存文件
                        with open(file_path, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                        
                        certificate_files.append({
                            'original_name': filename,
                            'saved_path': file_path
                        })
                        print(f"保存证书文件到: {file_path}")
                    except Exception as e:
                        print(f"保存附件失败: {e}")
        
        if not certificate_files:
            print("未找到证书附件，跳过处理")
            return None
            
        # 3. 从邮件内容、主题或文件名中提取专利号和专利名称
        patent_number = None
        patent_name = None
        
        # 从邮件内容中提取专利号
        patent_number_patterns = [
            r'申请号[：:]*\s*([\d\.]+)',
            r'专利号[：:]*\s*([\d\.]+)',
            r'CN\s*([\d\.]+)',
            r'([\d]{12,13}\.[\d])',
            r'([\d]{8,10}\.[\d])',
        ]
        
        # 优先从邮件内容中提取
        for pattern in patent_number_patterns:
            matches = re.search(pattern, content)
            if matches:
                patent_number = matches.group(1).strip()
                print(f"从邮件内容中提取到专利号: {patent_number}")
                break
        
        # 如果邮件内容中没有找到，尝试从主题中提取
        if not patent_number:
            for pattern in patent_number_patterns:
                matches = re.search(pattern, subject)
                if matches:
                    patent_number = matches.group(1).strip()
                    print(f"从邮件主题中提取到专利号: {patent_number}")
                    break
        
        # 从邮件内容中提取专利名称
        patent_name_patterns = [
            r'发明名称[：:]*\s*([^\n\r]+)',
            r'专利名称[：:]*\s*([^\n\r]+)',
            r'名称[：:]*\s*([^\n\r]+)'
        ]
        
        for pattern in patent_name_patterns:
            matches = re.search(pattern, content)
            if matches:
                patent_name = matches.group(1).strip()
                print(f"从邮件内容中提取到专利名称: {patent_name}")
                break
        
        # 解析日期
        date_str = date
        try:
            # 尝试解析日期字符串为datetime对象
            from email.utils import parsedate_to_datetime
            parsed_date = parsedate_to_datetime(date)
            date_str = parsed_date.strftime('%Y-%m-%d')
        except Exception as e:
            print(f"日期解析失败: {e}，使用原始日期字符串")
        
        # 构建结果
        result = {
            'patent_number': patent_number,
            'patent_name': patent_name,
            'email_subject': subject,
            'email_from': from_addr,
            'email_date': date_str,
            'certificate_files': certificate_files,
            'download_urls': [f'/download/certificate/{os.path.basename(cf["saved_path"])}' for cf in certificate_files],
            'filename': f"专利证书-{patent_number if patent_number else '未知专利号'}-{patent_name if patent_name else subject[:20]}"
        }
        
        print(f"提取结果: {result}")
        return result
            
        