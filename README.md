<p align="center">
  <img src="https://img.shields.io/badge/Python-3.10+-blue?style=flat-square&logo=python" alt="Python">
  <img src="https://img.shields.io/badge/PySide6-GUI-green?style=flat-square&logo=qt" alt="GUI">
  <img src="https://img.shields.io/badge/license-MIT-orange?style=flat-square" alt="License">
</p>

# Excel 智能全文检索系统 / Excel Intelligent Search Tool

> 一个现代化的 Excel 批量搜索工具，支持在大量 `.xlsx` / `.xlsm` 文件中快速检索学生信息。由 **沈宇 (Shen Yu)** 开发。

## 功能特性 (Features)

- **智能列识别** — 自动检测姓名、学号、班级列，无需手动配置表头
- **多线程并发搜索** — 并发处理多个文件，大幅缩短搜索时间
- **多维搜索** — 支持同时搜索多个关键词（姓名、学号等），用空格或逗号分隔
- **结果导出** — 将搜索结果一键导出为 `.xlsx` 汇总表格
- **现代 UI** — 基于 PySide6 的无边框窗口、自定义标题栏、进度条反馈
- **可取消搜索** — 支持随时中止长时间搜索任务
- **启动屏动画** — 程序启动时有精美的启动过渡动画

## 界面预览 (Screenshots)

<p align="center">
  <img src="excel_search_tool/5386e28338ba5d1a1c69101fbaf2967a.png" width="45%" alt="搜索界面">
  <img src="excel_search_tool/fdc47aa4b700830f598d56c060a5f830.png" width="45%" alt="联系作者">
</p>

## 技术栈 (Tech Stack)

| 技术 | 用途 |
|------|------|
| **Python 3.10+** | 核心语言 |
| **PySide6** | GUI 框架 (Qt for Python) |
| **openpyxl** | Excel 文件读写 |
| **Pillow** | 图片支持 |

## 快速开始 (Quick Start)

### 1. 克隆仓库 (Clone)

```bash
git clone https://github.com/你的用户名/excel-search-tool.git
cd excel-search-tool
```

### 2. 创建虚拟环境 (Setup venv)

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate
```

### 3. 安装依赖 (Install Dependencies)

```bash
pip install -r excel_search_tool/requirements.txt
```

### 4. 运行 (Run)

```bash
python excel_search_tool/main.py
```

### 5. 打包为 EXE (Build)

```bash
python excel_search_tool/build.py
```

## 使用方法 (Usage)

1. 点击**浏览**，选择包含 `.xlsx` 文件的文件夹
2. 在搜索框中输入关键字（支持多关键字，用空格或逗号分隔）
3. 点击**开始分析搜索**，等待结果
4. 双击结果行可直接打开对应文件
5. 点击**导出汇总表格**将结果保存为 Excel 文件

## 项目结构 (Project Structure)

```
tool/
├── excel_search_tool/
│   ├── main.py            # 主程序
│   ├── requirements.txt   # Python 依赖
│   ├── build.py           # 打包脚本
│   └── *.png              # 界面资源图片
├── .gitignore
└── README.md
```

## 作者 (Author)

**沈宇 (Shen Yu)** — 软工24-3

如果有任何问题或建议，欢迎通过程序内的「联系作者」功能扫码联系我。

---

*本工具主要用于教育场景下的学生信息批量检索，请勿用于非法用途。*
*This tool is designed for educational use (batch student info lookup). Please do not use it for illegal purposes.*
