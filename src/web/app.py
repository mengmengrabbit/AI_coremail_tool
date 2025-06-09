from flask import Flask, render_template, jsonify, request
import sys
import os

# 添加项目根目录到系统路径
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))

# 导入模块
from src.core.email_manager import EmailManager
from datetime import datetime

# 导入测试邮件管理器
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../../tests')))
from test_emails import TestEmailManager

# 设置模板和静态文件目录
app = Flask(__name__, 
           template_folder=os.path.abspath(os.path.join(os.path.dirname(__file__), '../../templates')),
           static_folder=os.path.abspath(os.path.join(os.path.dirname(__file__), '../../static')))

# 尝试使用真实的邮件管理器，如果失败则使用测试数据
try:
    print("\n=== 初始化邮件管理器 ===")
    email_manager = EmailManager()
    
    # 检查邮件文件夹路径是否存在
    old_patent_folder = email_manager.old_patent_folder
    if not os.path.exists(old_patent_folder):
        print(f"警告: 专利审查提醒邮件文件夹不存在: {old_patent_folder}")
        raise FileNotFoundError(f"邮件文件夹不存在: {old_patent_folder}")
    
    # 测试是否能访问邮件文件夹
    test_reminders = email_manager.get_patent_examination_reminders()
    print(f"使用真实邮件数据，找到 {len(test_reminders)} 条提醒")
    
    if len(test_reminders) == 0:
        print("警告: 未找到任何专利审查提醒邮件，将使用测试数据")
        email_manager = TestEmailManager()
        print(f"使用测试数据模式，加载了 {len(email_manager.get_patent_examination_reminders())} 条测试提醒")
        
except Exception as e:
    print(f"无法访问真实邮件文件夹: {e}")
    print("使用测试数据模式")
    email_manager = TestEmailManager()
    print(f"使用测试数据模式，加载了 {len(email_manager.get_patent_examination_reminders())} 条测试提醒")

@app.route('/')
def index():
    """主页面"""
    return render_template('index.html')

@app.route('/api/patent-examination-reminders')
def get_patent_examination_reminders():
    """获取专利审查临期提醒数据"""
    try:
        print("\n=== API调用: /api/patent-examination-reminders ===")
        reminders = email_manager.get_patent_examination_reminders()
        print(f"API返回: 找到 {len(reminders)} 条专利审查临期提醒")
        
        # 转换datetime对象为字符串
        for reminder in reminders:
            if isinstance(reminder['deadline'], datetime):
                reminder['deadline_str'] = reminder['deadline'].strftime('%Y年%m月%d日')
                reminder['deadline_iso'] = reminder['deadline'].isoformat()
                # 计算距离今天的天数
                days_left = (reminder['deadline'] - datetime.now()).days
                reminder['days_left'] = days_left
                reminder['urgency_level'] = 'high' if days_left <= 7 else 'medium' if days_left <= 30 else 'low'
                print(f"处理提醒: 申请号={reminder['application_no']}, 期限={reminder['deadline_str']}, 紧急程度={reminder['urgency_level']}")
        
        return jsonify({
            'success': True,
            'data': reminders,
            'count': len(reminders)
        })
    except Exception as e:
        print(f"API错误: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/patent-certificates')
def get_patent_certificates():
    """获取专利授权证书汇总"""
    return jsonify({
        'success': True,
        'data': [],
        'message': '功能待实现'
    })

@app.route('/api/patent-invoices')
def get_patent_invoices():
    """获取专利费用发票汇总"""
    return jsonify({
        'success': True,
        'data': [],
        'message': '功能待实现'
    })

@app.route('/api/software-notices')
def get_software_notices():
    """获取软件协会通知汇总"""
    return jsonify({
        'success': True,
        'data': [],
        'message': '功能待实现'
    })

@app.route('/api/stats')
def get_stats():
    """获取统计信息"""
    try:
        print("\n=== API调用: /api/stats ===")
        reminders = email_manager.get_patent_examination_reminders()
        
        # 统计紧急程度
        high_urgency = sum(1 for r in reminders if (r['deadline'] - datetime.now()).days <= 7)
        medium_urgency = sum(1 for r in reminders if 7 < (r['deadline'] - datetime.now()).days <= 30)
        low_urgency = len(reminders) - high_urgency - medium_urgency
        
        print(f"统计结果: 总数={len(reminders)}, 高紧急度={high_urgency}, 中紧急度={medium_urgency}, 低紧急度={low_urgency}")
        
        return jsonify({
            'success': True,
            'data': {
                'total_reminders': len(reminders),
                'high_urgency': high_urgency,
                'medium_urgency': medium_urgency,
                'low_urgency': low_urgency
            }
        })
    except Exception as e:
        print(f"API错误: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)