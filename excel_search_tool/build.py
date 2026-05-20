import os
import sys
import subprocess

def build():
    print("开始打包 Excel 智能搜索助手...")
    try:
        import PyInstaller
    except ImportError:
        print("未检测到 PyInstaller，正在为您自动安装...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        
    try:
        import openpyxl
    except ImportError:
        print("正在安装 openpyxl...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
        
    try:
        import PySide6
    except ImportError:
        print("正在安装 PySide6...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyside6"])

    spec_path = "Excel搜索工具.spec"
    if not os.path.exists(spec_path):
        print(f"未找到 spec 配置文件: {spec_path}")
        return

    print("正在执行 PyInstaller 进行打包...")
    # 使用 spec 文件打包，使用 python -m PyInstaller 避免环境变量缺失的问题
    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", spec_path]
    result = subprocess.run(cmd)
    if result.returncode == 0:
        print("\n打包成功！可执行文件已生成在 dist 目录下。")
    else:
        print("\n打包失败，请检查报错日志。")

if __name__ == "__main__":
    build()
