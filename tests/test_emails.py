import os
import sys
import json
from datetime import datetime, timedelta

# 添加项目根目录到系统路径
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.core.email_manager import EmailManager

def create_test_email_data():
    """创建测试邮件数据"""
    test_emails = [
        # 符合用户要求的邮件1 - 包含第一种内容模式
        {
            'file_path': 'test_email_1.eml',
            'subject': '【提醒函】CN 申请号：202310123456.7;贵方编号：ABC-2023-001;我方编号：XYZ-2023-001',
            'from': 'info@sptl.com.cn',
            'date': '2024-01-15',
            'content': '''尊敬的客户：

您好！

关于中国国家知识产权局发出的第一次审查意见通知书，请留意该通知书的答复期限为：2025年09月29日。详细内容请见附件。

申请号：202310123456.7
申请名称：一种新型智能设备

请您及时准备相关材料并在期限内提交答复。

此致
敬礼！

专利代理机构'''
        },
        # 符合用户要求的邮件2 - 包含第二种内容模式
        {
            'file_path': 'test_email_2.eml',
            'subject': '【提醒函】CN 申请号：202310234567.8;贵方编号：DEF-2023-002;我方编号：XYZ-2023-002',
            'from': 'info@sptl.com.cn',
            'date': '2024-01-20',
            'content': '''尊敬的客户：

您好！

请参见我方于2025年3月11日报告的《第1次审查意见通知书》，我方至今尚未收到贵方的相关指令。根据该通知书规定，答复该通知书的期限是2025年07月04日。我方希望能尽快收到贵方的相关指令，以便更好地准备和提交答辩文件。

申请号：202310234567.8
申请名称：智能控制系统及方法

请您及时准备相关材料并在期限内提交答复。

此致
敬礼！

专利代理机构'''
        },
        # 逾期邮件
        {
            'file_path': 'test_email_3.eml',
            'subject': '【提醒函】CN 申请号：202310345678.9;贵方编号：GHI-2023-003;我方编号：XYZ-2023-003',
            'from': 'info@sptl.com.cn',
            'date': '2024-01-25',
            'content': '''尊敬的客户：

您好！

关于中国国家知识产权局发出的第一次审查意见通知书，请留意该通知书的答复期限为：2024年01月15日。详细内容请见附件。

申请号：202310345678.9
申请名称：数据处理装置

请您及时准备相关材料并在期限内提交答复。

此致
敬礼！

专利代理机构'''
        },
        # 一周内到期邮件
        {
            'file_path': 'test_email_4.eml',
            'subject': '【提醒函】CN 申请号：202310456789.0;贵方编号：JKL-2023-004;我方编号：XYZ-2023-004',
            'from': 'info@sptl.com.cn',
            'date': '2024-12-20',
            'content': '''尊敬的客户：

您好！

请参见我方于2024年10月15日报告的《第1次审查意见通知书》，我方至今尚未收到贵方的相关指令。根据该通知书规定，答复该通知书的期限是2025年01月05日。我方希望能尽快收到贵方的相关指令，以便更好地准备和提交答辩文件。

申请号：202310456789.0
申请名称：智能传感器系统

请您及时准备相关材料并在期限内提交答复。

此致
敬礼！

专利代理机构'''
        },
        # 正常期限邮件
        {
            'file_path': 'test_email_5.eml',
            'subject': '【提醒函】CN 申请号：202310567890.1;贵方编号：MNO-2023-005;我方编号：XYZ-2023-005',
            'from': 'info@sptl.com.cn',
            'date': '2024-12-25',
            'content': '''尊敬的客户：

您好！

关于中国国家知识产权局发出的第一次审查意见通知书，请留意该通知书的答复期限为：2025年02月15日。详细内容请见附件。

申请号：202310567890.1
申请名称：自动化控制装置

请您及时准备相关材料并在期限内提交答复。

此致
敬礼！

专利代理机构'''
        },
        # 不符合要求的邮件 - 发件人不对
        {
            'file_path': 'test_email_6.eml',
            'subject': '【提醒函】CN 申请号：202310678901.2;贵方编号：PQR-2023-006;我方编号：XYZ-2023-006',
            'from': 'someone@cffex.com.cn',
            'date': '2024-12-25',
            'content': '''尊敬的客户：

您好！

关于中国国家知识产权局发出的第一次审查意见通知书，请留意该通知书的答复期限为：2025年03月15日。详细内容请见附件。

申请号：202310678901.2
申请名称：机器学习算法

请您及时准备相关材料并在期限内提交答复。

此致
敬礼！

专利代理机构'''
        },
        # 不符合要求的邮件 - 内容不匹配
        {
            'file_path': 'test_email_7.eml',
            'subject': '【提醒函】CN 申请号：202310789012.3;贵方编号：STU-2023-007;我方编号：XYZ-2023-007',
            'from': 'info@sptl.com.cn',
            'date': '2024-12-25',
            'content': '''尊敬的客户：

您好！

根据国家知识产权局的通知，您的专利申请需要在规定期限内答复审查意见通知书。

申请号：202310789012.3
申请名称：普通专利申请
答复期限：2025年04月15日

请您及时准备相关材料并在期限内提交答复。

此致
敬礼！

专利代理机构'''
        }
    ]
    
    return test_emails

class TestEmailManager(EmailManager):
    """测试用的邮件管理器"""
    
    def __init__(self):
        # 不调用父类的__init__，避免路径问题
        self.test_emails = create_test_email_data()
        self.old_patent_folder = "D:\\学习\\2025\\app\\tools\\coremail-connect\\data\\收件箱\\老专利代理审查提醒"  # 添加缺失的属性
        self.base_path = "D:\\学习\\2025\\app\\tools\\coremail-connect\\data\\收件箱"  # 添加缺失的base_path属性
        self.invoices_folder = "D:\\学习\\2025\\app\\tools\\coremail-connect\\data\\收件箱\\新专利代理事务提醒"  # 添加缺失的invoices_folder属性
        # 初始化数据库管理器（如果需要的话）
        try:
            from src.utils.database import DatabaseManager
            self.db_manager = DatabaseManager()
        except:
            self.db_manager = None
    
    def get_emails_from_folder(self, folder_path):
        """返回测试邮件数据"""
        return self.test_emails
    
    def parse_email_file(self, file_path):
        """模拟解析邮件文件"""
        # 在实际应用中，这里会解析真实的邮件文件
        for email in self.test_emails:
            if email['file_path'] == file_path:
                return email
        return None
    
    def get_completion_stats(self):
        """获取完成统计信息"""
        reminders = self.get_patent_examination_reminders(include_completed=True)
        
        # 统计紧急程度
        now = datetime.now()
        high_urgency = sum(1 for r in reminders if (r['deadline'] - now).days <= 7)
        medium_urgency = sum(1 for r in reminders if 7 < (r['deadline'] - now).days <= 30)
        low_urgency = len(reminders) - high_urgency - medium_urgency
        
        return {
            'total': len(reminders),
            'high_urgency': high_urgency,
            'medium_urgency': medium_urgency,
            'low_urgency': low_urgency
        }

def test_patent_examination_reminders():
    """测试专利审查提醒功能"""
    print("=== 测试专利审查临期提醒功能 ===")
    
    manager = TestEmailManager()
    reminders = manager.get_patent_examination_reminders()
    
    print(f"找到 {len(reminders)} 条专利审查提醒")
    
    for i, reminder in enumerate(reminders, 1):
        print(f"\n{i}. 申请号: {reminder['application_no']}")
        print(f"   贵方编号: {reminder['client_no']}")
        print(f"   我方编号: {reminder['our_no']}")
        try:
            print(f"   答复期限: {reminder['deadline'].strftime('%Y-%m-%d')}")
        except UnicodeEncodeError:
            print(f"   答复期限: {reminder['deadline'].isoformat()}")
        print(f"   主题: {reminder['subject']}")
        print(f"   内容摘要: {reminder['content'][:100]}...")

if __name__ == '__main__':
    test_patent_examination_reminders()