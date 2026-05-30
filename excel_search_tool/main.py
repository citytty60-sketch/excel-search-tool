import os
import sys
import threading
import re
import time
import gc
import subprocess
from concurrent.futures import ProcessPoolExecutor, as_completed

# PySide6
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QLineEdit, QFrame, QProgressBar,
    QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, 
    QHeaderView, QAbstractItemView, QSizeGrip, QComboBox, QMenu
)
from PySide6.QtCore import Qt, Signal, QSettings, QPoint, QTimer, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QColor, QPainter, QMouseEvent, QFont, QIcon, QPixmap, QAction
from PySide6.QtWidgets import QGraphicsOpacityEffect

# ============================================================
# BMW M-Power 视觉规范
# ============================================================
M_LIGHT_BLUE = "#00A3E0"
M_BLUE       = "#0066B1"
M_DARK_BLUE  = "#003A70"
M_RED        = "#E4002B"
M_DARK       = "#101820"
CANVAS       = "#FFFFFF"
CANVAS_SOFT  = "#F2F5F8"
TEXT_INK     = "#111827"
TEXT_MUTE    = "#6B7280"
BORDER       = "#DDE4ED"

FONT_MAIN = "'Inter', 'SF Pro Display', 'Microsoft YaHei UI', sans-serif"

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ============================================================
# 极速引擎 (Rust Calamine) + 修复版启发式识别
# ============================================================
def fast_worker_rust(file_path, keywords):
    results = []
    fname = os.path.basename(file_path)
    try:
        from python_calamine import CalamineWorkbook
        wb = CalamineWorkbook.from_path(file_path)
        for sheet_name in wb.sheet_names:
            sheet = wb.get_sheet_by_name(sheet_name)
            if not sheet: continue
            rows_iter = sheet.iter_rows()
            
            sample = []
            try:
                for _ in range(30):
                    r = next(rows_iter, None)
                    if r is None: break
                    sample.append(r)
            except: pass
            if not sample: continue
            
            num_cols = max(len(r) for r in sample)
            scores = {"姓名": [0]*num_cols, "学号": [0]*num_cols, "班级": [0]*num_cols}
            
            for row in sample:
                for idx, val in enumerate(row):
                    if val is None or idx >= num_cols: continue
                    v_str = str(val).strip()
                    if not v_str: continue
                    
                    v_low = v_str.lower()
                    # 姓名特征：2-4位中文
                    if 2 <= len(v_str) <= 4 and re.match(r'^[\u4e00-\u9fa5]{2,4}$', v_str):
                        scores["姓名"][idx] += 2
                    
                    # 学号特征优化：排除 100 以内的序号，优先识别包含年份或字母的学号
                    if re.match(r'^[a-z0-9-]{7,15}$', v_low):
                        scores["学号"][idx] += 5 # 长度特征是学号的核心
                        if any(c.isalpha() for c in v_low): scores["学号"][idx] += 3 # 含字母加分
                        if "202" in v_low: scores["学号"][idx] += 4 # 含年份强加分
                    elif v_str.isdigit() and int(v_str) < 100:
                        scores["学号"][idx] -= 10 # 极大概率是序号，强力扣分

                    # 班级特征
                    if "班" in v_str or re.search(r'[0-9]{2,4}-[0-9]{1,2}', v_str):
                        scores["班级"][idx] += 2
                    
                    # 标题行特征匹配（权重最高）
                    if any(x in v_low for x in ["姓名", "name"]): scores["姓名"][idx] += 20
                    if any(x in v_low for x in ["学号", "id", "学籍", "卡号"]): scores["学号"][idx] += 20
                    if any(x in v_low for x in ["班级", "class"]): scores["班级"][idx] += 20

            col_map = {"姓名": -1, "学号": -1, "班级": -1}
            for key in col_map:
                max_score = max(scores[key])
                if max_score > 0:
                    col_map[key] = scores[key].index(max_score)

            import itertools
            for row_tuple in itertools.chain(sample, rows_iter):
                if not row_tuple: continue
                row_str = " ".join(str(v).strip() for v in row_tuple if v is not None).lower()
                if all(kw in row_str for kw in keywords):
                    results.append({
                        "filename": fname,
                        "filepath": file_path,
                        "sheet": sheet_name,
                        "name": str(row_tuple[col_map["姓名"]]).strip() if 0 <= col_map["姓名"] < len(row_tuple) and row_tuple[col_map["姓名"]] is not None else "未识别",
                        "id": str(row_tuple[col_map["学号"]]).strip() if 0 <= col_map["学号"] < len(row_tuple) and row_tuple[col_map["学号"]] is not None else "未识别",
                        "class": str(row_tuple[col_map["班级"]]).strip() if 0 <= col_map["班级"] < len(row_tuple) and row_tuple[col_map["班级"]] is not None else "未识别"
                    })
    except: pass
    return results

# ============================================================
# UI 组件库
# ============================================================
class ContactOverlay(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setVisible(False)
        self.setFixedSize(parent.size() if parent else (1100, 850))
        
    def paintEvent(self, event):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor(0, 0, 0, 100)) # 半透明遮罩
        
    def show_card(self):
        self.setFixedSize(self.parent().size())
        self.show()
        self.raise_()
        
        # 居中显示一个卡片
        card = QFrame(self)
        card.setFixedSize(360, 420)
        card.setStyleSheet(f"background: white; border-radius: 12px; border: 1px solid {BORDER};")
        card.move((self.width()-360)//2, (self.height()-420)//2)
        card.show()
        
        layout = QVBoxLayout(card)
        layout.setContentsMargins(30, 20, 30, 30)
        
        top = QHBoxLayout()
        title = QLabel("联系作者 / 打赏"); title.setStyleSheet("font-weight: 800; font-size: 16px;")
        close_btn = QPushButton("✕"); close_btn.setFixedSize(30, 30); close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        close_btn.setStyleSheet("background: transparent; border: none; font-size: 18px;")
        close_btn.clicked.connect(self.hide)
        top.addWidget(title); top.addStretch(); top.addWidget(close_btn)
        layout.addLayout(top)
        
        info = QLabel("如有问题或打赏，请扫码添加作者微信：")
        info.setStyleSheet(f"color: {TEXT_MUTE}; font-size: 13px; margin: 5px 0;")
        info.setWordWrap(True)
        layout.addWidget(info)
        
        img = QLabel()
        img_path = resource_path("2.png")
        if os.path.exists(img_path):
            pix = QPixmap(img_path).scaled(240, 240, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            img.setPixmap(pix)
        else:
            img.setText("[ 2.png 未找到 ]")
        img.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(img)
        
        name = QLabel("沈宇 (软工24-3)")
        name.setAlignment(Qt.AlignmentFlag.AlignCenter)
        name.setStyleSheet("font-weight: 700; font-size: 14px; margin-top: 10px;")
        layout.addWidget(name)

class SplashScreen(QWidget):
    finished = Signal()
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setFixedSize(500, 280)
        self.setStyleSheet(f"background: white; border: 1px solid {BORDER};")
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0); layout.setSpacing(0)
        
        stripe = QWidget(); stripe.setFixedHeight(6)
        sl = QHBoxLayout(stripe); sl.setContentsMargins(0, 0, 0, 0); sl.setSpacing(0)
        for c in (M_LIGHT_BLUE, M_BLUE, M_RED):
            f = QFrame(); f.setStyleSheet(f"background: {c}; border: none;"); sl.addWidget(f)
        layout.addWidget(stripe)
        
        content = QVBoxLayout(); content.setContentsMargins(40, 40, 40, 40); content.setAlignment(Qt.AlignmentFlag.AlignCenter)
        t = QLabel("Excel 智能全文检索系统")
        t.setStyleSheet(f"font-size: 26px; font-weight: 900; color: {M_DARK}; border: none;")
        author = QLabel("本软件由软工24-3 沈宇 开发完成")
        author.setStyleSheet(f"font-size: 14px; color: {TEXT_MUTE}; margin-top: 10px; border: none;")
        self.pb = QProgressBar(); self.pb.setFixedHeight(4); self.pb.setTextVisible(False); self.pb.setRange(0, 100)
        self.pb.setStyleSheet(f"QProgressBar {{ background: {CANVAS_SOFT}; border: none; }} QProgressBar::chunk {{ background: {M_BLUE}; }}")
        
        content.addWidget(t, 0, Qt.AlignmentFlag.AlignCenter)
        content.addWidget(author, 0, Qt.AlignmentFlag.AlignCenter)
        content.addStretch()
        content.addWidget(self.pb)
        layout.addLayout(content)
        
        self.timer = QTimer(); self.timer.timeout.connect(self._tick)
        self.val = 0; self.timer.start(25)
        geo = QApplication.primaryScreen().availableGeometry()
        self.move((geo.width()-500)//2, (geo.height()-280)//2)

    def _tick(self):
        self.val += 2; self.pb.setValue(self.val)
        if self.val >= 100: self.timer.stop(); self.finished.emit(); self.close()

class TitleBar(QFrame):
    def __init__(self, window):
        super().__init__()
        self.window = window
        self.setFixedHeight(46)
        self.setStyleSheet(f"background: white; border-bottom: 1px solid {BORDER};")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 0, 0, 0)
        title = QLabel("Excel 智能全文检索系统")
        title.setStyleSheet(f"color: {M_DARK}; font-weight: 700; font-size: 14px;")
        layout.addWidget(title); layout.addStretch()
        for icon, cmd in [("—", "min"), ("✕", "close")]:
            btn = QPushButton(icon); btn.setFixedSize(46, 46); btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setStyleSheet(f"QPushButton {{ background: transparent; color: {M_DARK}; border: none; font-size: 14px; }} "
                            f"QPushButton:hover {{ background: {CANVAS_SOFT}; {'color: white; background: #e81123;' if cmd=='close' else ''} }}")
            if cmd == "min": btn.clicked.connect(window.showMinimized)
            else: btn.clicked.connect(window.close)
            layout.addWidget(btn)
        self._drag_pos = None

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton: self._drag_pos = e.position().toPoint()
    def mouseMoveEvent(self, e):
        if self._drag_pos: self.window.move(e.globalPosition().toPoint() - self._drag_pos)
    def mouseReleaseEvent(self, e): self._drag_pos = None

# ============================================================
# 主应用
# ============================================================
class ExcelSearchApp(QWidget):
    progress_sig = Signal(int, int, str)
    finished_sig = Signal(list)

    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.resize(1100, 850)
        self.setAcceptDrops(True)
        self.searching = False
        self.stop_requested = False
        self.start_time = 0
        self.total_files = 0
        self.settings = QSettings("M-Power", "ExcelSearchV4")
        self.results_data = []
        self.dark_mode = self.settings.value("dark_mode", "false") == "true"
        
        self._init_ui()
        self._apply_theme()
        self._load_config()
        
        self.overlay = ContactOverlay(self)
        self.progress_sig.connect(self._on_progress)
        self.finished_sig.connect(self._on_finished)

    def _init_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0); root.setSpacing(0)
        root.addWidget(TitleBar(self))

        self.body = QWidget()
        root.addWidget(self.body)
        self.main_layout = QVBoxLayout(self.body); self.main_layout.setContentsMargins(30, 25, 30, 25); self.main_layout.setSpacing(15)

        self.card = QFrame()
        self.main_layout.addWidget(self.card)
        cl = QVBoxLayout(self.card); cl.setContentsMargins(25, 20, 25, 20); cl.setSpacing(12)

        def create_row(label, is_combo=False):
            row = QHBoxLayout()
            lbl = QLabel(label); lbl.setFixedWidth(70); lbl.setStyleSheet("font-weight: 700;")
            if is_combo:
                edit = QComboBox(); edit.setEditable(True); edit.setFixedHeight(40)
            else:
                edit = QLineEdit(); edit.setFixedHeight(40)
            row.addWidget(lbl); row.addWidget(edit)
            if label == "搜索路径":
                btn = QPushButton("浏览"); btn.setFixedSize(70, 40); btn.setCursor(Qt.CursorShape.PointingHandCursor)
                btn.clicked.connect(self._browse)
                row.addWidget(btn)
                self.path_browse_btn = btn
            cl.addLayout(row)
            return edit, lbl

        self.path_edit, self.path_label = create_row("搜索路径", is_combo=True)
        self.kw_edit, self.kw_label = create_row("搜索内容")
        self.kw_edit.setPlaceholderText("输入关键字，支持空格分隔多词...")

        self.hint = QLabel("专业可能有缩写，建议搜索简写（如'人工'）以及完整称呼（如'人工智能'），进行两次搜索以确保结果准确")
        self.hint.setStyleSheet(f"color: {M_BLUE}; font-size: 12px; font-weight: 500; margin-left: 75px;")
        cl.addWidget(self.hint)

        btn_box = QHBoxLayout(); btn_box.setContentsMargins(0, 10, 0, 0); btn_box.setSpacing(12)
        self.search_btn = QPushButton(" 开始分析搜索"); self.search_btn.setFixedSize(160, 42); self.search_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.search_btn.clicked.connect(self._toggle_search)
        
        self.export_btn = QPushButton(" 导出汇总表格"); self.export_btn.setFixedSize(150, 42); self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self._export)

        self.theme_btn = QPushButton(" 切换深色模式"); self.theme_btn.setFixedSize(130, 42); self.theme_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.theme_btn.clicked.connect(self._toggle_theme)

        contact_btn = QPushButton(" 联系作者 / 打赏"); contact_btn.setFixedSize(150, 42); contact_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        contact_btn.clicked.connect(lambda: self.overlay.show_card())
        self.contact_btn = contact_btn

        btn_box.addWidget(self.search_btn); btn_box.addWidget(self.export_btn); btn_box.addWidget(self.theme_btn); btn_box.addStretch(); btn_box.addWidget(contact_btn)
        self.main_layout.addLayout(btn_box)

        # 结果过滤行
        filter_box = QHBoxLayout()
        self.filter_edit = QLineEdit(); self.filter_edit.setPlaceholderText("🔍 在结果中快速过滤..."); self.filter_edit.setFixedHeight(32)
        self.filter_edit.textChanged.connect(self._filter_table)
        filter_box.addStretch(); filter_box.addWidget(self.filter_edit, 0)
        self.main_layout.addLayout(filter_box)

        self.stack = QFrame(); self.stack.setStyleSheet("border: none;")
        sl = QVBoxLayout(self.stack); sl.setContentsMargins(0, 0, 0, 0)
        # 移除最后的“操作”列以提升大批量数据渲染性能
        self.table = QTableWidget(); self.table.setColumnCount(5); self.table.setHorizontalHeaderLabels(["文件名", "工作表", "姓名", "学号", "班级"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers); self.table.setAlternatingRowColors(True)
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self.table.itemDoubleClicked.connect(lambda item: self._open_file(self.results_data[item.row()]["filepath"]))
        
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setDefaultSectionSize(36)
        
        # 绑定回车搜索
        self.kw_edit.returnPressed.connect(self._toggle_search)
        self.path_edit.lineEdit().returnPressed.connect(self._toggle_search)
        
        # 为路径编辑框增加右键菜单清理历史
        self.path_edit.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.path_edit.customContextMenuRequested.connect(self._path_context_menu)

        self.empty_widget = QWidget()
        el = QVBoxLayout(self.empty_widget); el.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.folder_icon = QLabel("📁"); self.folder_icon.setStyleSheet("font-size: 64px; border: none;")
        self.t1 = QLabel("暂无搜索结果"); self.t1.setStyleSheet(f"font-size: 16px; font-weight: 700; border: none;")
        self.t2 = QLabel("输入关键字并选择文件夹后，点击搜索按钮开始"); self.t2.setStyleSheet(f"font-size: 13px; border: none;")
        el.addWidget(self.folder_icon, 0, Qt.AlignmentFlag.AlignCenter); el.addWidget(self.t1, 0, Qt.AlignmentFlag.AlignCenter); el.addWidget(self.t2, 0, Qt.AlignmentFlag.AlignCenter)
        sl.addWidget(self.table); sl.addWidget(self.empty_widget); self.table.hide()
        self.main_layout.addWidget(self.stack)

        self.prog = QProgressBar(); self.prog.setFixedHeight(4); self.prog.setTextVisible(False)
        self.status = QLabel("就绪")
        self.status.setStyleSheet(f"font-size: 12px;")
        self.main_layout.addWidget(self.prog); self.main_layout.addWidget(self.status)

    def _apply_theme(self):
        bg = M_DARK if self.dark_mode else CANVAS
        bg_soft = "#1E293B" if self.dark_mode else CANVAS_SOFT
        text = "#F8FAFC" if self.dark_mode else TEXT_INK
        text_mute = "#94A3B8" if self.dark_mode else TEXT_MUTE
        border = "#334155" if self.dark_mode else BORDER
        
        self.body.setStyleSheet(f"background: {bg};")
        self.card.setStyleSheet(f"background: {bg}; border: 1px solid {border}; border-radius: 6px;")
        self.path_label.setStyleSheet(f"color: {text}; font-weight: 700;")
        self.kw_label.setStyleSheet(f"color: {text}; font-weight: 700;")
        self.path_edit.setStyleSheet(f"QComboBox {{ border: 1px solid {border}; border-radius: 4px; padding: 0 12px; background: {bg}; color: {text}; }} QComboBox::drop-down {{ border: none; }}")
        self.kw_edit.setStyleSheet(f"border: 1px solid {border}; border-radius: 4px; padding: 0 12px; background: {bg}; color: {text};")
        self.path_browse_btn.setStyleSheet(f"background: {bg}; border: 1px solid {border}; border-radius: 4px; color: {text};")
        
        btn_style_base = f"QPushButton {{ background: white; border: 1px solid {border}; border-radius: 6px; color: {M_DARK}; }} QPushButton:hover {{ background: {CANVAS_SOFT}; }}"
        if self.dark_mode:
            btn_style_base = f"QPushButton {{ background: #1E293B; border: 1px solid {border}; border-radius: 6px; color: white; }} QPushButton:hover {{ background: #334155; }}"
        
        self.theme_btn.setStyleSheet(btn_style_base)
        self.contact_btn.setStyleSheet(btn_style_base)
        self.export_btn.setStyleSheet(f"QPushButton {{ background: {bg_soft}; border: 1px solid {border}; color: {text_mute}; font-weight: 700; border-radius: 6px; }} QPushButton:enabled {{ background: {bg}; color: {text}; }}")
        
        self.filter_edit.setStyleSheet(f"border: 1px solid {border}; border-radius: 16px; padding: 0 15px; background: {bg_soft}; color: {text}; font-size: 12px;")
        
        self.table.setStyleSheet(f"""
            QTableWidget {{ background: {bg}; border: 1px solid {border}; border-radius: 6px; gridline-color: transparent; selection-background-color: {M_BLUE}; selection-color: white; color: {text}; }} 
            QHeaderView::section {{ background: {bg_soft}; padding: 10px; font-weight: 700; border: none; color: {text}; }}
            QScrollBar:vertical {{ border: none; background: {bg_soft}; width: 8px; margin: 0px; border-radius: 4px; }}
            QScrollBar::handle:vertical {{ background: {border}; min-height: 20px; border-radius: 4px; }}
            QScrollBar::handle:vertical:hover {{ background: {M_BLUE}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        """)
        
        self.t1.setStyleSheet(f"font-size: 16px; font-weight: 700; color: {text_mute}; border: none;")
        self.t2.setStyleSheet(f"font-size: 13px; color: {text_mute}; border: none;")
        self.status.setStyleSheet(f"color: {text_mute}; font-size: 12px;")
        self.prog.setStyleSheet(f"QProgressBar {{ background: {bg_soft}; border: none; }} QProgressBar::chunk {{ background: {M_RED}; }}")
        
        self.theme_btn.setText(" 切换浅色模式" if self.dark_mode else " 切换深色模式")
        self.settings.setValue("dark_mode", "true" if self.dark_mode else "false")

    def _toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self._apply_theme()

    def _load_config(self):
        history = self.settings.value("path_history", [])
        if history:
            self.path_edit.addItems(history)
            self.path_edit.setCurrentText(history[0])
        self.kw_edit.setText(self.settings.value("last_kw", ""))

    def _save_history(self, path):
        history = self.settings.value("path_history", [])
        if path in history: history.remove(path)
        history.insert(0, path)
        history = history[:5]
        self.settings.setValue("path_history", history)
        self.path_edit.clear()
        self.path_edit.addItems(history)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls(): e.accept()
    def dropEvent(self, e):
        urls = e.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if os.path.isfile(path): path = os.path.dirname(path)
            self.path_edit.setCurrentText(path)

    def _browse(self):
        p = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if p: self.path_edit.setCurrentText(p)

    def _filter_table(self, text):
        text = text.lower()
        for i in range(self.table.rowCount()):
            match = False
            for j in range(self.table.columnCount()):
                item = self.table.item(i, j)
                if item and text in item.text().lower():
                    match = True; break
            self.table.setRowHidden(i, not match)

    def _show_context_menu(self, pos):
        item = self.table.itemAt(pos)
        if not item: return
        row = item.row()
        file_path = self.results_data[row]["filepath"]
        menu = QMenu(self)
        menu.setStyleSheet(f"QMenu {{ background: white; border: 1px solid {BORDER}; }} QMenu::item:selected {{ background: {CANVAS_SOFT}; color: {M_BLUE}; }}")
        act_open = QAction("打开文件", self)
        act_open.triggered.connect(lambda: self._open_file(file_path))
        act_folder = QAction("打开所在文件夹", self)
        act_folder.triggered.connect(lambda: self._open_folder(file_path))
        menu.addAction(act_open); menu.addAction(act_folder)
        menu.exec(self.table.viewport().mapToGlobal(pos))

    def _path_context_menu(self, pos):
        menu = QMenu(self)
        menu.setStyleSheet(f"QMenu {{ background: white; border: 1px solid {BORDER}; }} QMenu::item:selected {{ background: {CANVAS_SOFT}; color: {M_RED}; }}")
        clear_act = QAction("清空历史记录", self)
        clear_act.triggered.connect(self._clear_history)
        menu.addAction(clear_act)
        menu.exec(self.path_edit.mapToGlobal(pos))

    def _clear_history(self):
        self.settings.setValue("path_history", [])
        self.path_edit.clear()
        QMessageBox.information(self, "提示", "历史记录已清空。")

    def _toggle_search(self):
        if self.searching:
            self.stop_requested = True; return
        path = self.path_edit.currentText().strip(); kw = self.kw_edit.text().strip()
        if not path or not kw or not os.path.isdir(path): return
        self._save_history(path)
        self.settings.setValue("last_kw", kw)
        self.searching = True; self.stop_requested = False
        self.start_time = time.time()
        self.search_btn.setText(" 停止分析搜索"); self.search_btn.setStyleSheet(f"background: {M_RED}; color: white; font-weight: 700; border-radius: 6px;")
        self.table.hide(); self.empty_widget.hide(); self.table.setRowCount(0); self.export_btn.setEnabled(False)
        threading.Thread(target=self._run_engine, args=(path, kw), daemon=True).start()

    def _run_engine(self, path, kw_str):
        keywords = [k.lower() for k in re.split(r'[,，\s]+', kw_str) if k]
        files = []
        for r, _, fs in os.walk(path):
            for f in fs:
                if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$"):
                    files.append(os.path.join(r, f))
        self.total_files = len(files)
        if not files: self.finished_sig.emit([]); return
        final = []
        with ProcessPoolExecutor(max_workers=min(os.cpu_count(), 8)) as exe:
            futures = {exe.submit(fast_worker_rust, f, keywords): f for f in files}
            count = 0
            for fut in as_completed(futures):
                if self.stop_requested: exe.shutdown(wait=False, cancel_futures=True); break
                count += 1
                try:
                    res = fut.result()
                    if res: final.extend(res)
                except: pass
                if count % 10 == 0 or count == len(files): self.progress_sig.emit(count, len(files), f"分析中: {count}/{len(files)}")
        self.finished_sig.emit(final)

    def _on_progress(self, c, t, msg):
        self.prog.setMaximum(t); self.prog.setValue(c); self.status.setText(msg)

    def _on_finished(self, results):
        self.searching = False; self.search_btn.setText(" 开始分析搜索")
        self.search_btn.setStyleSheet(f"background: {M_BLUE}; color: white; font-weight: 700; border-radius: 6px;")
        self.results_data = results
        duration = time.time() - self.start_time
        if not results:
            self.table.hide(); self.empty_widget.show()
        else:
            self.empty_widget.hide()
            self._render_table(results)
            self.table.show()
            self.export_btn.setEnabled(True)
        self.status.setText(f"完成 | 耗时 {duration:.2f}s | 扫描 {self.total_files} 文件 | 找到 {len(results)} 条结果")

    def _render_table(self, data):
        self.table.setUpdatesEnabled(False)
        self.table.setRowCount(len(data))
        for i, r in enumerate(data):
            self.table.setItem(i, 0, QTableWidgetItem(r["filename"]))
            self.table.setItem(i, 1, QTableWidgetItem(r["sheet"]))
            self.table.setItem(i, 1, QTableWidgetItem(r["sheet"]))
            self.table.setItem(i, 2, QTableWidgetItem(r["name"]))
            self.table.setItem(i, 3, QTableWidgetItem(r["id"]))
            self.table.setItem(i, 4, QTableWidgetItem(r["class"]))
        self.table.setUpdatesEnabled(True)

    def _open_file(self, path):
        if os.name == 'nt': os.startfile(path)
        else: subprocess.run(['open', path])

    def _open_folder(self, path):
        if os.name == 'nt': subprocess.run(['explorer', '/select,', os.path.normpath(path)])
        else: subprocess.run(['open', os.path.dirname(path)])

    def _export(self):
        path, _ = QFileDialog.getSaveFileName(self, "导出结果", "Results.xlsx", "Excel (*.xlsx)")
        if path:
            import openpyxl
            wb = openpyxl.Workbook(); ws = wb.active; ws.append(["文件名", "路径", "工作表", "姓名", "学号", "班级"])
            for r in self.results_data: ws.append([r["filename"], r["filepath"], r["sheet"], r["name"], r["id"], r["class"]])
            wb.save(path); QMessageBox.information(self, "成功", "结果已成功导出。")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'overlay') and self.overlay.isVisible():
            self.overlay.setFixedSize(self.size())

if __name__ == "__main__":
    app = QApplication(sys.argv)
    splash = SplashScreen()
    def start_main():
        main = ExcelSearchApp(); main.show()
    splash.finished.connect(start_main)
    splash.show()
    sys.exit(app.exec())
