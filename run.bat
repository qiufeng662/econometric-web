# -*- coding: utf-8 -*-
"""
一键实证分析 - 桌面启动器
"""

import subprocess
import sys
import os

def main():
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    app_file = os.path.join(current_dir, 'app.py')
    
    # 启动 Streamlit
    print("=" * 50)
    print("一键实证分析 正在启动...")
    print("=" * 50)
    print("\n请在浏览器中访问: http://localhost:8501")
    print("\n按 Ctrl+C 停止服务\n")
    
    try:
        subprocess.run([
            sys.executable, '-m', 'streamlit', 'run', app_file,
            '--server.headless', 'true',
            '--browser.gatherUsageStats', 'false'
        ])
    except KeyboardInterrupt:
        print("\n服务已停止")

if __name__ == '__main__':
    main()