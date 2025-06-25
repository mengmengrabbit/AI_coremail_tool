import os
import sys
import webbrowser
import time
from threading import Thread

def check_dependencies():
    """检查依赖是否已安装"""
    try:
        import flask
        import chardet
        import requests
        print("✓ 所有依赖已安装")
        return True
    except ImportError as e:
        print(f"✗ 缺少依赖: {e}")
        return False

def install_dependencies():
    """安装所需依赖"""
    print("正在安装依赖...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "src/requirements.txt"])
        print("✓ 依赖安装成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ 依赖安装失败: {e}")
        return False

def start_flask_app():
    """在单独线程中启动Flask应用"""
    # 添加项目根目录到Python路径
    project_root = os.path.dirname(os.path.abspath(__file__))
    if project_root not in sys.path:
        sys.path.insert(0, project_root)
    
    # 添加src目录到Python路径
    src_path = os.path.join(project_root, 'src')
    if src_path not in sys.path:
        sys.path.insert(0, src_path)
    
    # 添加tests目录到Python路径
    tests_path = os.path.join(project_root, 'tests')
    if tests_path not in sys.path:
        sys.path.insert(0, tests_path)
    
    # 导入Flask应用
    from src.web.app import app
    
    # 启动Flask应用
    app.run(debug=False, host='0.0.0.0', port=5000, use_reloader=False)

def start_server():
    """启动Flask服务器"""
    print("\n=== 邮件智能管理系统 ===\n")
    print("正在启动服务...")
    
    # 检查依赖
    if not check_dependencies():
        if not install_dependencies():
            print("无法安装依赖，请手动运行: pip install -r requirements.txt")
            input("按任意键退出...")
            return
    
    # 启动Flask应用
    try:
        print("正在启动Web服务器...")
        
        # 在单独线程中启动Flask应用
        flask_thread = Thread(target=start_flask_app, daemon=True)
        flask_thread.start()
        
        # 等待服务器启动
        time.sleep(3)
        
        # 打开浏览器
        url = "http://localhost:5000"
        print(f"正在打开浏览器: {url}")
        webbrowser.open(url)
        
        print("\n服务已启动!")
        print("访问地址: http://localhost:5000")
        print("按Ctrl+C停止服务")
        
        # 保持主线程运行
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n正在停止服务...")
            print("服务已停止")
    
    except Exception as e:
        print(f"启动服务失败: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    start_server()