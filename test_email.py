import sys
sys.path.append('D:\\学习\\2025\\app\\tools\\coremail-connect\\src')
from core.email_manager import EmailManager

em = EmailManager("D:\\学习\\2025\\app\\tools\\coremail-connect\\data\\收件箱")

# 获取指定目录的邮件
folder_path = "D:\\学习\\2025\\app\\tools\\coremail-connect\\data\\收件箱\\老专利代理审查提醒"
emails = em.get_emails_from_folder(folder_path)

print(f"找到 {len(emails)} 封邮件")

# 查找目标邮件
target_filename = 'CN 申请号：202210038761.0;贵方编号：;我方编号：CNJRQH-0131.219852.eml'
target_email = None

for email in emails:
    filename = email.get('filename', '')
    print(f"检查邮件: {filename}")
    if filename == target_filename:
        target_email = email
        print(f"找到目标邮件: {filename}")
        break

if target_email:
    print("\n开始处理目标邮件...")
    result = em.extract_patent_certificate_info(target_email, 'd:/temp/certificates')
    print(f"处理结果: {result}")
else:
    print("未找到目标邮件")
    print("\n所有邮件文件名:")
    for email in emails[:10]:  # 只显示前10个
        print(f"  - {email.get('filename', '')}")