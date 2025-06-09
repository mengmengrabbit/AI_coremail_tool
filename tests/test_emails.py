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
        {
            'file_path': 'test_email_1.eml',
            'subject': '【提醒函】CN 申请号：202310123456.7;贵方编号：ABC-2023-001;我方编号：XYZ-2023-001',
            'from': 'info@sptl.com.cn',
            'date': '2024-01-15',
            'content': '''尊敬的客户：

您好！

根据国家知识产权局的通知，您的专利申请需要在规定期限内答复审查意见通知书。

申请号：202310123456.7
申请名称：一种新型智能设备
答复该通知书的期限是2024年02月15日

请您及时准备相关材料并在期限内提交答复。

此致
敬礼！

专利代理机构'''
        },
        {
            'file_path': 'test_email_2.eml',
            'subject': '【提醒函】CN 申请号：202310234567.8;贵方编号：DEF-2023-002;我方编号：XYZ-2023-002',
            'from': 'info@sptl.com.cn',
            'date': '2024-01-20',
            'content': '''尊敬的客户：

您好！

根据国家知识产权局的通知，您的专利申请需要在规定期限内答复审查意见通知书。

申请号：202310234567.8
申请名称：智能控制系统及方法
答复该通知书的期限是2024年03月01日

请您及时准备相关材料并在期限内提交答复。

此致
敬礼！

专利代理机构'''
        },
        {
            'file_path': 'test_email_3.eml',
            'subject': '【提醒函】CN 申请号：202310345678.9;贵方编号：GHI-2023-003;我方编号：XYZ-2023-003',
            'from': 'info@sptl.com.cn',
            'date': '2024-01-25',
            'content': '''尊敬的客户：

您好！

根据国家知识产权局的通知，您的专利申请需要在规定期限内答复审查意见通知书。

申请号：202310345678.9
申请名称：数据处理装置
答复期限：2024年01月30日

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
        print(f"   答复期限: {reminder['deadline'].strftime('%Y年%m月%d日')}")
        print(f"   主题: {reminder['subject']}")
        print(f"   内容摘要: {reminder['content'][:100]}...")

if __name__ == '__main__':
    test_patent_examination_reminders()