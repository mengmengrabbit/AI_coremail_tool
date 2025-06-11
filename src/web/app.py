import os
import sys
from datetime import datetime
from flask import Flask, render_template, jsonify, request

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
# 添加tests目录到Python路径
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), '..', 'tests'))

from core.email_manager import EmailManager
from test_emails import TestEmailManager

# 设置模板和静态文件目录
app = Flask(__name__, 
           template_folder=os.path.abspath(os.path.join(os.path.dirname(__file__), '../../templates')),
           static_folder=os.path.abspath(os.path.join(os.path.dirname(__file__), '../../static')))

# 初始化邮件管理器并预加载数据
print("\n=== 初始化邮件管理器 ===")
email_manager = None
use_test_data = False
cached_reminders = []
cached_stats = {}

try:
    # 先尝试创建真实的邮件管理器
    temp_manager = EmailManager()
    old_patent_folder = temp_manager.old_patent_folder
    
    if os.path.exists(old_patent_folder):
        print(f"邮件文件夹路径存在: {old_patent_folder}")
        email_manager = temp_manager
        print("开始预加载邮件数据...")
        
        # 预加载数据
        cached_reminders = email_manager.get_patent_examination_reminders(include_completed=True)
        print(f"预加载完成: 找到 {len(cached_reminders)} 条专利审查临期提醒")
        
        # 预计算统计数据 - 基于紧急程度而不是完成状态
        now = datetime.now()
        high_urgency = sum(1 for r in cached_reminders if (r['deadline'] - now).days <= 7)
        medium_urgency = sum(1 for r in cached_reminders if 7 < (r['deadline'] - now).days <= 30)
        low_urgency = len(cached_reminders) - high_urgency - medium_urgency
        
        cached_stats = {
            'total': len(cached_reminders),
            'high_urgency': high_urgency,
            'medium_urgency': medium_urgency,
            'low_urgency': low_urgency
        }
        print(f"统计数据: 总数={cached_stats['total']}, 高紧急度={cached_stats['high_urgency']}, 中紧急度={cached_stats['medium_urgency']}, 低紧急度={cached_stats['low_urgency']}")
        
    else:
        print(f"警告: 专利审查提醒邮件文件夹不存在: {old_patent_folder}")
        raise FileNotFoundError(f"邮件文件夹不存在: {old_patent_folder}")
        
except Exception as e:
    print(f"无法访问真实邮件文件夹: {e}")
    print("使用测试数据模式")
    email_manager = TestEmailManager()
    use_test_data = True
    cached_reminders = email_manager.get_patent_examination_reminders(include_completed=True)
    
    # 计算测试模式的统计数据
    now = datetime.now()
    high_urgency = sum(1 for r in cached_reminders if (r['deadline'] - now).days <= 7)
    medium_urgency = sum(1 for r in cached_reminders if 7 < (r['deadline'] - now).days <= 30)
    low_urgency = len(cached_reminders) - high_urgency - medium_urgency
    
    cached_stats = {
        'total': len(cached_reminders),
        'high_urgency': high_urgency,
        'medium_urgency': medium_urgency,
        'low_urgency': low_urgency
    }
    print(f"测试数据模式: 加载了 {len(cached_reminders)} 条测试提醒")

@app.route('/')
def index():
    """主页面"""
    return render_template('index.html')

@app.route('/api/patent-examination-reminders')
def get_patent_examination_reminders():
    """获取专利审查临期提醒数据"""
    try:
        print("\n=== API调用: /api/patent-examination-reminders ===")
        include_completed = request.args.get('include_completed', 'false').lower() == 'true'
        print(f"参数: include_completed={include_completed}")
        
        # 使用缓存的提醒数据
        reminders = cached_reminders
        
        # 如果不包含已完成的，需要过滤
        if not include_completed:
            reminders = [r for r in reminders if not r.get('completed', False)]
        
        print(f"返回 {len(reminders)} 条提醒")
        
        # 处理数据格式
        processed_reminders = []
        for i, reminder in enumerate(reminders):
            try:
                processed_reminder = dict(reminder)  # 创建副本
                
                # 处理deadline字段
                if 'deadline' in processed_reminder and processed_reminder['deadline']:
                    deadline = processed_reminder['deadline']
                    if isinstance(deadline, datetime):
                        processed_reminder['deadline_str'] = deadline.strftime('%Y年%m月%d日')
                        processed_reminder['deadline_iso'] = deadline.isoformat()
                    else:
                        processed_reminder['deadline_str'] = str(deadline)
                        processed_reminder['deadline_iso'] = str(deadline)
                else:
                    processed_reminder['deadline_str'] = '未知'
                    processed_reminder['deadline_iso'] = ''
                
                processed_reminders.append(processed_reminder)
                print(f"处理提醒 {i+1}: 申请号={processed_reminder.get('application_no', 'N/A')}, 期限={processed_reminder['deadline_str']}")
                
            except Exception as item_error:
                print(f"处理第 {i+1} 条提醒时出错: {str(item_error)}")
                continue
        
        print(f"API成功返回: {len(processed_reminders)} 条处理后的提醒")
        return jsonify({
            'success': True,
            'data': processed_reminders,
            'count': len(processed_reminders)
        })
        
    except Exception as e:
        import traceback
        error_msg = f"API错误: {str(e)}"
        print(error_msg)
        print(f"错误详情: {traceback.format_exc()}")
        return jsonify({
            'success': False,
            'error': error_msg
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

@app.route('/api/mark-completed', methods=['POST'])
def mark_reminder_completed():
    """标记提醒为已完成"""
    try:
        data = request.get_json()
        application_no = data.get('application_no')
        file_path = data.get('file_path')
        subject = data.get('subject')
        deadline = data.get('deadline')
        
        if not all([application_no, file_path, subject, deadline]):
            return jsonify({
                'success': False,
                'error': '缺少必要参数'
            }), 400
        
        # 转换deadline字符串为datetime对象
        if isinstance(deadline, str):
            deadline = datetime.fromisoformat(deadline.replace('Z', '+00:00'))
        
        success = email_manager.mark_reminder_completed(application_no, file_path, subject, deadline)
        
        # 更新缓存中的邮件状态
        global cached_reminders
        for reminder in cached_reminders:
            if reminder.get('application_no') == application_no and reminder.get('file_path') == file_path:
                reminder['completed'] = True
                print(f"已更新缓存中的邮件状态: {application_no} 标记为已完成")
                break
        
        return jsonify({
            'success': success,
            'message': '标记成功' if success else '标记失败'
        })
    except Exception as e:
        print(f"标记完成API错误: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/mark-uncompleted', methods=['POST'])
def mark_reminder_uncompleted():
    """标记提醒为未完成"""
    try:
        data = request.get_json()
        application_no = data.get('application_no')
        file_path = data.get('file_path')
        
        if not all([application_no, file_path]):
            return jsonify({
                'success': False,
                'error': '缺少必要参数'
            }), 400
        
        success = email_manager.mark_reminder_uncompleted(application_no, file_path)
        
        # 更新缓存中的邮件状态
        global cached_reminders
        for reminder in cached_reminders:
            if reminder.get('application_no') == application_no and reminder.get('file_path') == file_path:
                reminder['completed'] = False
                print(f"已更新缓存中的邮件状态: {application_no} 标记为未完成")
                break
        
        return jsonify({
            'success': success,
            'message': '取消标记成功' if success else '取消标记失败'
        })
    except Exception as e:
        print(f"取消标记API错误: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/completion-stats')
def get_completion_stats():
    """获取完成状态统计"""
    try:
        stats = email_manager.get_completion_stats()
        return jsonify({
            'success': True,
            'data': stats
        })
    except Exception as e:
        print(f"获取统计API错误: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

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
    print("\n=== API调用: /api/stats ===")
    
    try:
        # 使用缓存的统计数据
        return jsonify(cached_stats)
    except Exception as e:
        print(f"获取统计信息时出错: {e}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)