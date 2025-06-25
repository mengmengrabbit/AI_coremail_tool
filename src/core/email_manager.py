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
import quopri
import os

# 添加utils路径
sys.path.append(os.path.join(os.path.dirname(__file__), '../utils'))
from database import DatabaseManager
from utils.settings import EMAIL_FOLDER_PATH

class EmailManager:
    """邮件管理器 - 处理本地邮件文件夹"""
    
    def __init__(self, base_path=None):
        self.base_path = base_path or EMAIL_FOLDER_PATH
        self.old_patent_folder = os.path.join(self.base_path, "老专利代理审查提醒")
        self.new_patent_folder = os.path.join(self.base_path, "新专利代理事务提醒")
        # 初始化数据库管理器
        self.db_manager = DatabaseManager()
        
        # 初始化邮件分类缓存
        self.classification_cache = {}
        
        # 创建专利授权证书保存目录
        self.certificates_folder = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'static', 'certificates')
        os.makedirs(self.certificates_folder, exist_ok=True)
        
        # 创建专利费用发票保存目录
        self.invoices_folder = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'static', 'invoices')
        os.makedirs(self.invoices_folder, exist_ok=True)
        
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
    
    def _clean_html_content(self, html_content):
        """清理HTML内容，转换为纯文本，增强编码处理"""
        if not html_content.strip():
            return ""
        
        html_to_process = html_content
        
        # 首先尝试解码quoted-printable编码
        try:
            if '=' in html_to_process and re.search(r'=[0-9A-F]{2}', html_to_process):
                # 检测到quoted-printable编码，进行解码
                decoded_bytes = quopri.decodestring(html_to_process.encode('ascii', errors='ignore'))
                # 尝试多种编码解码
                for encoding in ['utf-8', 'gbk', 'gb2312', 'gb18030', 'big5']:
                    try:
                        html_to_process = decoded_bytes.decode(encoding, errors='strict')
                        print(f"quoted-printable解码成功，使用编码: {encoding}")
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    # 如果所有编码都失败，使用utf-8并忽略错误
                    html_to_process = decoded_bytes.decode('utf-8', errors='ignore')
                    print(f"quoted-printable解码使用utf-8备用方案")
        except Exception as e:
            print(f"quoted-printable解码失败: {e}")
            pass
        
        # 增强的HTML清理逻辑
        # 1. 移除CSS样式块（包括内联样式）
        html_to_process = re.sub(r'<style[^>]*>.*?</style>', '', html_to_process, flags=re.DOTALL | re.IGNORECASE)
        html_to_process = re.sub(r'<link[^>]*stylesheet[^>]*>', '', html_to_process, flags=re.IGNORECASE)
        
        # 2. 移除script标签
        html_to_process = re.sub(r'<script[^>]*>.*?</script>', '', html_to_process, flags=re.DOTALL | re.IGNORECASE)
        
        # 3. 移除CSS类定义（如.csD270A203{...}）
        html_to_process = re.sub(r'\.[a-zA-Z0-9_-]+\{[^}]*\}', '', html_to_process, flags=re.DOTALL)
        
        # 4. 移除HTML注释
        html_to_process = re.sub(r'<!--.*?-->', '', html_to_process, flags=re.DOTALL)
        
        # 5. 移除head标签及其内容
        html_to_process = re.sub(r'<head[^>]*>.*?</head>', '', html_to_process, flags=re.DOTALL | re.IGNORECASE)
        
        # 6. 替换常见的HTML实体
        html_to_process = html_to_process.replace('&nbsp;', ' ')
        html_to_process = html_to_process.replace('&lt;', '<')
        html_to_process = html_to_process.replace('&gt;', '>')
        html_to_process = html_to_process.replace('&amp;', '&')
        html_to_process = html_to_process.replace('&quot;', '"')
        html_to_process = html_to_process.replace('&apos;', "'")
        html_to_process = html_to_process.replace('&copy;', '©')
        html_to_process = html_to_process.replace('&reg;', '®')
        html_to_process = html_to_process.replace('&trade;', '™')
        
        # 7. 将块级元素标签替换为换行符
        block_tags = ['br', 'p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'tr', 'td', 'th']
        for tag in block_tags:
            html_to_process = re.sub(f'<{tag}[^>]*>', '\n', html_to_process, flags=re.IGNORECASE)
            html_to_process = re.sub(f'</{tag}>', '\n', html_to_process, flags=re.IGNORECASE)
        
        # 8. 移除所有剩余的HTML标签
        display_content = re.sub(r'<[^>]+>', '', html_to_process)
        
        # 9. 清理多余的空白字符和特殊字符
        display_content = re.sub(r'\n\s*\n', '\n\n', display_content)  # 多个连续换行变为两个
        display_content = re.sub(r'[ \t]+', ' ', display_content)  # 多个空格变为一个
        display_content = re.sub(r'^\s+|\s+$', '', display_content, flags=re.MULTILINE)  # 移除行首行尾空格
        display_content = display_content.strip()
        
        return display_content
    
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
            
            # 对纯文本内容也进行quoted-printable解码
            if content.strip():
                try:
                    if '=' in content and re.search(r'=[0-9A-F]{2}', content):
                        decoded_bytes = quopri.decodestring(content.encode('ascii', errors='ignore'))
                        content = decoded_bytes.decode('utf-8', errors='ignore')
                        print(f"纯文本内容检测到quoted-printable编码，解码后: {content[:200]}...")
                except Exception as e:
                    print(f"纯文本quoted-printable解码失败: {e}")
                    pass
            
            # 如果纯文本内容为空但HTML内容不为空，使用HTML内容
            if not content.strip() and html_content.strip():
                # 使用增强的HTML清理逻辑
                content = self._clean_html_content(html_content)
            
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
        
        # 1. 检查发件人，只处理来自sptl.com.cn的邮件
        # 严格限制：只处理专利代理机构的邮件
        if '@sptl.com.cn' not in from_addr.lower():
            print(f"跳过邮件：发件人不是sptl.com.cn，发件人：{from_addr}，标题：{subject[:50]}")
            return None
            
        # 2. 检查邮件内容是否包含专利相关信息（放宽条件）
        content_patterns = [
            r'关于中国国家知识产权局发出的第一次审查意见通知书，请留意该通知书的答复期限为：(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'请参见我方于(\d{4})年(\d{1,2})月(\d{1,2})日报告的《第1次审查意见通知书》.*?答复该通知书的期限是(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'请参见我方于(\d{4})年(\d{1,2})月(\d{1,2})日报告的.*?第1次审查意见通知书.*?答复.*?期限.*?(\d{4})年(\d{1,2})月(\d{1,2})日',
            # 新增：更宽泛的专利相关内容匹配
            r'申请号.*?(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'专利.*?(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'贵方编号.*?我方编号.*?(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'到期日.*?(\d{4})年(\d{1,2})月(\d{1,2})日',
            r'期限.*?(\d{4})年(\d{1,2})月(\d{1,2})日'
        ]
        
        content_match = False
        for pattern in content_patterns:
            if re.search(pattern, content, re.DOTALL | re.IGNORECASE):
                content_match = True
                print(f"邮件内容匹配模式：{pattern[:50]}...")
                break
        
        # 如果内容不匹配但包含专利关键词，也允许通过
        if not content_match:
            if ('申请号' in content or '专利' in content or '贵方编号' in content or '我方编号' in content):
                content_match = True
                print(f"邮件包含专利关键词，允许处理")
                
        if not content_match:
            print(f"跳过邮件：内容不匹配专利相关模式")
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
        
        # 处理HTML内容，转换为纯文本
        html_content = email.get('html_content', '')
        display_content = content
        
        # 如果content为空但html_content不为空，或者content看起来像HTML，则处理HTML内容
        if (not content.strip() and html_content.strip()) or ('<' in content and '>' in content):
            # 使用html_content或content中的HTML内容
            html_to_process = html_content if html_content.strip() else content
            
            # 使用统一的HTML清理方法
            display_content = self._clean_html_content(html_to_process)
            
            # 如果处理后内容为空或只有少量无意义字符，尝试从原始内容中提取有意义的文本
            if len(display_content.strip()) < 10:
                # 尝试提取可能的文本内容
                text_patterns = [
                    r'申请号[：:]?\s*([\d\.]+)',
                    r'我方编号[：:]?\s*([\w\-\.]+)',
                    r'贵方编号[：:]?\s*([\w\-\.]*)',
                    r'答复期限[：:]?\s*(\d{4}年\d{1,2}月\d{1,2}日)',
                    r'期限[：:]?\s*(\d{4}年\d{1,2}月\d{1,2}日)'
                ]
                extracted_info = []
                for pattern in text_patterns:
                    matches = re.findall(pattern, content + html_content)
                    if matches:
                        extracted_info.extend(matches)
                
                if extracted_info:
                    display_content = '\n'.join(extracted_info)
        
        # 从内容中提取专利名称作为标题
        patent_name = None
        patent_name_patterns = [
            # 匹配专利名称（在贵方编号之前的内容）
            r'([^\n\r]+?)贵方编号',
            # 匹配专利名称（在我方编号之前的内容）
            r'([^\n\r]+?)我方编号',
            # 匹配专利名称（在申请号之后到编号之前的内容）
            r'申请号[：:]?[\d\.]+[^\n\r]*?([^\n\r]+?)(?:贵方编号|我方编号)',
            # 匹配以"一种"开头的专利名称
            r'(一种[^\n\r]+?)(?:贵方编号|我方编号|申请号)',
            # 匹配以"基于"开头的专利名称
            r'(基于[^\n\r]+?)(?:贵方编号|我方编号|申请号)',
            # 匹配其他可能的专利名称模式
            r'([^\n\r]{10,50}?)(?:贵方编号|我方编号)'
        ]
        
        try:
            for pattern in patent_name_patterns:
                match = re.search(pattern, display_content)
                if match:
                    potential_name = match.group(1).strip()
                    # 过滤掉明显不是专利名称的内容
                    if (len(potential_name) > 5 and 
                        not re.match(r'^[\d\s\.\-：:]+$', potential_name) and
                        '李悦萌' not in potential_name and
                        '女士' not in potential_name and
                        '您好' not in potential_name):
                        patent_name = potential_name
                        break
        except Exception as e:
            print(f"提取专利名称时出错: {e}")
            patent_name = None
        
        # 如果找到了专利名称，从内容中移除它并用作标题
        final_subject = subject
        final_content = display_content
        
        if patent_name:
            final_subject = patent_name
            # 从内容中移除专利名称部分，保留其余内容
            final_content = re.sub(re.escape(patent_name), '', display_content, count=1).strip()
            # 清理开头可能残留的标点符号
            final_content = re.sub(r'^[\s：:]+', '', final_content)
        
        return {
            'application_no': application_no,
            'client_no': client_no,
            'our_no': our_no,
            'deadline': deadline,
            'days_left': days_left,
            'urgency_level': urgency_level,
            'urgency_text': urgency_text,
            'subject': final_subject,
            'original_subject': subject,  # 添加原始邮件标题
            'content': final_content[:500] + '...' if len(final_content) > 500 else final_content,
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
        # 用于去重的集合，基于申请号和期限的组合
        seen_reminders = set()
        
        for i, email in enumerate(all_emails):
            print(f"\n处理第 {i+1}/{len(all_emails)} 封邮件...")
            reminder_info = self.extract_patent_reminder_info(email)
            if reminder_info:
                # 创建去重键：申请号 + 期限
                dedup_key = f"{reminder_info['application_no']}_{reminder_info['deadline'].strftime('%Y-%m-%d')}"
                
                # 检查是否已存在相同的提醒
                if dedup_key in seen_reminders:
                    print(f"跳过重复邮件: 申请号={reminder_info['application_no']}, 期限={reminder_info['deadline'].strftime('%Y-%m-%d')}")
                    continue
                
                # 添加到已见集合
                seen_reminders.add(dedup_key)
                
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
        
        print(f"\n去重前找到 {len(all_emails)} 封邮件，去重后找到 {len(reminders)} 条专利审查临期提醒")
        
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
        
        # 从邮件内容和标题中提取专利名称
        patent_name_patterns = [
            r'发明名称[：:]*\s*([^\n\r]+)',
            r'专利名称[：:]*\s*([^\n\r]+)',
            r'名称[：:]*\s*([^\n\r]+)'
        ]
        
        # 首先从邮件内容中提取
        for pattern in patent_name_patterns:
            matches = re.search(pattern, content)
            if matches:
                patent_name = matches.group(1).strip()
                print(f"从邮件内容中提取到专利名称: {patent_name}")
                break
        
        # 如果内容中没有找到，尝试从邮件标题中提取发明名称
        if not patent_name:
            # 从邮件标题中提取专利名称的模式
            subject_patterns = [
                # 匹配"专利证书公告(发明)"格式中的发明名称
                r'([^-]+)\s*-\s*专利证书',
                r'([^-]+)\s*-\s*发明专利证书',
                r'([^-]+)\s*专利证书公告',
                # 匹配"一种..."开头的发明名称
                r'(一种[^-，,；;]+)',
                # 匹配"基于..."开头的发明名称
                r'(基于[^-，,；;]+)',
                # 匹配其他可能的发明名称模式
                r'([^-，,；;]*(?:方法|系统|装置|设备|平台|工具|技术)[^-，,；;]*)',
                # 匹配包含"及"的发明名称
                r'([^-，,；;]*及[^-，,；;]*)',
                # 提取邮件标题中第一个有意义的部分（排除Re:, Fw:等前缀）
                r'(?:Re:|Fw:|转发:|回复:)?\s*([^-，,；;]{10,50})'
            ]
            
            for pattern in subject_patterns:
                matches = re.search(pattern, subject)
                if matches:
                    potential_name = matches.group(1).strip()
                    # 过滤掉明显不是专利名称的内容
                    if (len(potential_name) >= 8 and 
                        not any(x in potential_name.lower() for x in ['专利证书', '公告', '通知', '提醒', '转发', '回复']) and
                        not re.match(r'^\d+', potential_name)):
                        patent_name = potential_name
                        print(f"从邮件标题中提取到专利名称: {patent_name}")
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
    
    def get_patent_invoices(self):
        """获取专利费用发票汇总"""
        print(f"\n开始获取专利费用发票邮件...")
        
        # 指定搜索目录：新专利代理事务提醒
        invoice_folder = os.path.join(self.base_path, "新专利代理事务提醒")
        print(f"扫描发票邮件路径: {invoice_folder}")
        
        if not os.path.exists(invoice_folder):
            print(f"警告: 发票邮件文件夹路径不存在: {invoice_folder}")
            return []
        
        # 获取该文件夹下的所有邮件
        all_emails = self.get_emails_from_folder(invoice_folder)
        print(f"在发票文件夹中找到 {len(all_emails)} 封邮件")
        
        invoices = []
        
        for i, email in enumerate(all_emails):
            print(f"\n处理第 {i+1}/{len(all_emails)} 封邮件...")
            invoice_info = self.extract_patent_invoice_info(email, self.invoices_folder)
            
            if invoice_info:
                invoices.append(invoice_info)
                print(f"成功提取专利发票信息: {invoice_info['filename']}")
        
        print(f"\n总共找到 {len(invoices)} 个专利费用发票")
        
        # 按日期排序（由新到旧）
        invoices.sort(key=lambda x: x.get('date', ''), reverse=True)
        
        return invoices
    
    def extract_patent_invoice_info(self, email, invoices_folder):
        """从邮件中提取专利发票信息"""
        from_addr = email.get('from', '')
        content = email.get('content', '')
        subject = email.get('subject', '')
        date = email.get('date', '')
        attachments = email.get('attachments', [])
        
        print(f"处理邮件: {subject[:100]}...")
        print(f"发件人: {from_addr}")
        print(f"附件数量: {len(attachments)}")
        
        # 1. 检查邮件标题是否包含"【科盛】转电子发票invoice-"
        if not ("【科盛】转电子发票invoice-" in subject):
            print(f"跳过邮件：标题不包含指定的发票关键信息")
            return None
        
        # 2. 检查附件并分类保存
        invoice_files = {
            'official_receipt': None,  # 官方票据 - 专利电子票据-xxx.pdf
            'notice': None,           # 通知 - invoice-xxx.pdf
            'agent_receipt': None,    # 代理票据 - dzfp_xxx.pdf
            'agent_xml': None        # 代理xml文件 - xxx.xml
        }
        
        for attachment in attachments:
            filename = attachment.get('filename', '')
            content_type = attachment.get('content_type', '')
            
            print(f"检查附件: {filename}")
            
            # 分类附件
            if filename.startswith('专利电子票据-') and filename.endswith('.pdf'):
                category = 'official_receipt'
                display_name = '官方票据'
            elif filename.startswith('invoice-') and filename.endswith('.pdf'):
                category = 'notice'
                display_name = '通知'
            elif filename.startswith('dzfp_') and filename.endswith('.pdf'):
                category = 'agent_receipt'
                display_name = '代理票据'
            elif filename.endswith('.xml'):
                category = 'agent_xml'
                display_name = '代理xml文件'
            else:
                print(f"跳过未分类的附件: {filename}")
                continue
            
            # 保存附件
            part = attachment.get('part')
            if part:
                try:
                    # 确保目录存在
                    os.makedirs(invoices_folder, exist_ok=True)
                    
                    # 生成唯一文件名
                    unique_filename = f"{uuid.uuid4()}_{filename}"
                    file_path = os.path.join(invoices_folder, unique_filename)
                    
                    # 保存文件
                    with open(file_path, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    
                    invoice_files[category] = {
                        'original_name': filename,
                        'saved_path': file_path,
                        'display_name': display_name,
                        'download_url': f'/download/invoice/{os.path.basename(file_path)}'
                    }
                    print(f"保存{display_name}文件到: {file_path}")
                except Exception as e:
                    print(f"保存附件失败: {e}")
        
        # 检查是否至少有一个附件被保存
        if not any(invoice_files.values()):
            print("未找到符合条件的发票附件，跳过处理")
            return None
        
        # 3. 从邮件标题中提取发票编号
        invoice_number = None
        invoice_match = re.search(r'invoice-(\d+)', subject)
        if invoice_match:
            invoice_number = invoice_match.group(1)
            print(f"提取到发票编号: {invoice_number}")
        
        # 解析日期
        date_str = date
        try:
            from email.utils import parsedate_to_datetime
            parsed_date = parsedate_to_datetime(date)
            date_str = parsed_date.strftime('%Y-%m-%d')
        except Exception as e:
            print(f"日期解析失败: {e}，使用原始日期字符串")
        
        # 构建结果
        result = {
            'invoice_number': invoice_number,
            'email_subject': subject,
            'email_from': from_addr,
            'email_date': date_str,
            'invoice_files': invoice_files,
            'filename': f"专利发票-{invoice_number if invoice_number else '未知编号'}"
        }
        
        print(f"提取结果: {result}")
        return result
    
    def get_software_notices(self):
        """获取软件协会通知汇总"""
        try:
            print("开始获取软件协会通知...")
            
            # 查找所有softline.org.cn邮件
            softline_emails = []
            for filename in os.listdir(self.base_path):
                if filename.endswith('.eml'):
                    file_path = os.path.join(self.base_path, filename)
                    try:
                        with open(file_path, 'rb') as f:
                            content = f.read()
                            if b'softline.org.cn' in content:
                                softline_emails.append(file_path)
                    except Exception as e:
                        print(f"读取文件 {filename} 失败: {e}")
                        continue
            
            print(f"找到 {len(softline_emails)} 个软件协会邮件")
            
            # 解析邮件并分类
            notices = []
            for email_path in softline_emails:
                try:
                    notice = self._parse_software_notice(email_path)
                    if notice:
                        notices.append(notice)
                except Exception as e:
                    print(f"解析邮件 {email_path} 失败: {e}")
                    continue
            
            print(f"成功解析 {len(notices)} 条软件协会通知")
            return notices
            
        except Exception as e:
            print(f"获取软件协会通知失败: {e}")
            return []
    
    def _extract_email_content(self, msg):
        """提取邮件内容，增强编码处理"""
        content = ""
        html_content = ""
        
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    payload = part.get_payload(decode=True)
                    if payload:
                        decoded_text = self._decode_payload_with_detection(payload, part)
                        content += decoded_text
                elif content_type == "text/html":
                    payload = part.get_payload(decode=True)
                    if payload:
                        decoded_text = self._decode_payload_with_detection(payload, part)
                        html_content += decoded_text
        else:
            content_type = msg.get_content_type()
            payload = msg.get_payload(decode=True)
            if payload:
                decoded_text = self._decode_payload_with_detection(payload, msg)
                if content_type == "text/html":
                    html_content += decoded_text
                else:
                    content += decoded_text
        
        # 如果纯文本内容为空但HTML内容不为空，使用HTML内容
        if not content.strip() and html_content.strip():
            content = self._clean_html_content(html_content)
        
        return content.strip()
    
    def _decode_payload_with_detection(self, payload, part):
        """使用编码检测解码邮件内容"""
        if not payload:
            return ""
        
        # 首先尝试从邮件头获取字符集
        charset = part.get_content_charset()
        
        if charset:
            try:
                return payload.decode(charset, errors='ignore')
            except (UnicodeDecodeError, LookupError) as e:
                print(f"使用指定字符集 {charset} 解码失败: {e}")
        
        # 尝试常见的编码格式
        common_encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030', 'big5', 'iso-8859-1', 'windows-1252']
        
        for encoding in common_encodings:
            try:
                decoded = payload.decode(encoding, errors='strict')
                print(f"成功使用 {encoding} 编码解码")
                return decoded
            except (UnicodeDecodeError, LookupError):
                continue
        
        # 使用chardet进行编码检测
        try:
            detected = chardet.detect(payload)
            detected_encoding = detected.get('encoding')
            confidence = detected.get('confidence', 0)
            
            print(f"chardet检测到编码: {detected_encoding}, 置信度: {confidence}")
            
            if detected_encoding and confidence > 0.7:
                try:
                    return payload.decode(detected_encoding, errors='ignore')
                except (UnicodeDecodeError, LookupError) as e:
                    print(f"使用检测到的编码 {detected_encoding} 解码失败: {e}")
        except Exception as e:
            print(f"编码检测失败: {e}")
        
        # 最后的备用方案：使用utf-8并忽略错误
        try:
            return payload.decode('utf-8', errors='ignore')
        except:
            # 如果所有方法都失败，返回字符串表示
            return str(payload)
    
    def _parse_software_notice(self, email_path):
        """解析单个软件协会通知邮件"""
        try:
            with open(email_path, 'rb') as f:
                msg = BytesParser(policy=policy.default).parse(f)
            
            # 提取基本信息
            subject = self.decode_header_value(msg.get('Subject', ''))
            from_addr = self.decode_header_value(msg.get('From', ''))
            date = msg.get('Date', '')
            
            # 提取邮件内容
            content = self._extract_email_content(msg)
            
            # 使用InternLM API进行分类（带缓存）
            category = self._classify_notice_with_internlm(email_path, subject, content)
            
            # 解析日期
            date_str = date
            try:
                from email.utils import parsedate_to_datetime
                parsed_date = parsedate_to_datetime(date)
                date_str = parsed_date.strftime('%Y-%m-%d')
            except Exception as e:
                print(f"日期解析失败: {e}")
            
            return {
                'subject': subject,
                'from': from_addr,
                'date': date_str,
                'content': content[:500] + '...' if len(content) > 500 else content,  # 截取前500字符
                'category': category,
                'filename': os.path.basename(email_path)
            }
            
        except Exception as e:
            print(f"解析邮件失败: {e}")
            return None
    
    def _classify_notice_with_internlm(self, email_path, subject, content):
        """使用InternLM API对通知进行分类（带缓存机制）"""
        try:
            # 生成缓存键：使用文件路径和修改时间
            file_stat = os.stat(email_path)
            cache_key = f"{email_path}_{file_stat.st_mtime}"
            
            # 检查缓存
            if cache_key in self.classification_cache:
                cached_result = self.classification_cache[cache_key]
                print(f"使用缓存的分类结果: {cached_result}")
                return cached_result
            
            print(f"邮件 {os.path.basename(email_path)} 未找到缓存，调用大模型进行分类...")
            
            from openai import OpenAI
            from dotenv import load_dotenv
            import os
            
            # 加载环境变量
            load_dotenv()
            InternLM_api_key = os.getenv("InternLM")
            
            if not InternLM_api_key:
                print("未找到InternLM API密钥")
                fallback_result = "未分类"
                # 即使是备用结果也要缓存
                self.classification_cache[cache_key] = fallback_result
                return fallback_result
            
            # 创建OpenAI客户端
            client = OpenAI(
                api_key=InternLM_api_key,
                base_url="https://chat.intern-ai.org.cn/api/v1/",
            )
            
            # 构建分类提示
            prompt = f"""请将以下软件协会通知邮件分类到四个类别之一：
1. 评奖评优公告
2. 活动提醒
3. 服务采购
4. 企业资质证书

邮件主题：{subject}
邮件内容：{content[:1000]}

请只回答类别名称，不要其他内容。"""
            
            # 调用API
            response = client.chat.completions.create(
                model="internlm2.5-latest",
                messages=[
                    {"role": "user", "content": prompt}
                ],
                max_tokens=50,
                temperature=0.1
            )
            
            category = response.choices[0].message.content.strip()
            print(f"AI分类结果: {category}")
            
            # 验证分类结果
            valid_categories = ["评奖评优公告", "活动提醒", "服务采购", "企业资质证书"]
            if category not in valid_categories:
                # 如果AI返回的不是预期的类别，尝试匹配
                for valid_cat in valid_categories:
                    if valid_cat in category:
                        category = valid_cat
                        break
                else:
                    category = "其他"
            
            # 将结果保存到缓存
            self.classification_cache[cache_key] = category
            print(f"分类结果已缓存: {cache_key} -> {category}")
            
            return category
            
        except Exception as e:
            print(f"AI分类失败: {e}")
            # 基于关键词的简单分类作为备选
            fallback_result = self._simple_classify_notice(subject, content)
            
            # 即使是备用结果也要缓存
            try:
                file_stat = os.stat(email_path)
                cache_key = f"{email_path}_{file_stat.st_mtime}"
                self.classification_cache[cache_key] = fallback_result
                print(f"备用分类结果已缓存: {cache_key} -> {fallback_result}")
            except:
                pass
            
            return fallback_result
    
    def _simple_classify_notice(self, subject, content):
        """基于关键词的简单分类"""
        text = (subject + " " + content).lower()
        
        if any(keyword in text for keyword in ['评奖', '评优', '公示', '获奖', '表彰','创新']):
            return "评奖评优公告"
        elif any(keyword in text for keyword in ['活动', '会议', '培训', '研讨', '论坛','理事']):
            return "活动提醒"
        elif any(keyword in text for keyword in ['采购', '招标', '服务', '报价']):
            return "服务采购"
        elif any(keyword in text for keyword in ['证书', '资质', '认证', '高新','企业']):
            return "企业资质证书"
        else:
            return "其他"
    
    def clear_classification_cache(self):
        """清空分类缓存"""
        cache_size = len(self.classification_cache)
        self.classification_cache.clear()
        print(f"已清空分类缓存，共清理 {cache_size} 条记录")
    
    def clean_expired_cache(self):
        """清理过期的缓存条目（文件已被修改或删除）"""
        expired_keys = []
        
        for cache_key in list(self.classification_cache.keys()):
            try:
                # 解析缓存键
                parts = cache_key.rsplit('_', 1)
                if len(parts) != 2:
                    expired_keys.append(cache_key)
                    continue
                
                file_path, cached_mtime = parts
                cached_mtime = float(cached_mtime)
                
                # 检查文件是否存在
                if not os.path.exists(file_path):
                    expired_keys.append(cache_key)
                    continue
                
                # 检查文件修改时间是否变化
                current_mtime = os.stat(file_path).st_mtime
                if current_mtime != cached_mtime:
                    expired_keys.append(cache_key)
                    
            except Exception as e:
                print(f"检查缓存键 {cache_key} 时出错: {e}")
                expired_keys.append(cache_key)
        
        # 删除过期的缓存
        for key in expired_keys:
            del self.classification_cache[key]
        
        if expired_keys:
            print(f"已清理 {len(expired_keys)} 条过期缓存")
        else:
            print("没有发现过期缓存")
    
    def get_cache_stats(self):
        """获取缓存统计信息"""
        total_cache = len(self.classification_cache)
        
        # 统计各分类的缓存数量
        category_stats = {}
        for cached_result in self.classification_cache.values():
            category_stats[cached_result] = category_stats.get(cached_result, 0) + 1
        
        return {
            'total_cached': total_cache,
            'category_distribution': category_stats
        }
            
        