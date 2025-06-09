import os
import sys
import subprocess
import webbrowser
import time

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
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✓ 依赖安装成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ 依赖安装失败: {e}")
        return False

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
        # 使用子进程启动Flask应用
        server_process = subprocess.Popen(
            [sys.executable, "web/app.py"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            universal_newlines=True
        )
        
        # 等待服务器启动
        print("正在启动Web服务器...")
        time.sleep(2)
        
        # 打开浏览器
        url = "http://localhost:5000"
        print(f"正在打开浏览器: {url}")
        webbrowser.open(url)
        
        print("\n服务已启动!")
        print("按Ctrl+C停止服务")
        
        # 保持程序运行，直到用户中断
        try:
            while True:
                output = server_process.stdout.readline()
                if output:
                    print(output.strip())
                
                error = server_process.stderr.readline()
                if error:
                    print(f"错误: {error.strip()}", file=sys.stderr)
                
                # 检查进程是否还在运行
                if server_process.poll() is not None:
                    break
                
                time.sleep(0.1)
        except KeyboardInterrupt:
            print("\n正在停止服务...")
            server_process.terminate()
            server_process.wait()
            print("服务已停止")
    
    except Exception as e:
        print(f"启动服务失败: {e}")

if __name__ == "__main__":
    start_server()