import openpyxl
import os
import sys
import threading
import subprocess
import re
from concurrent.futures import ThreadPoolExecutor

# PySide6
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QLineEdit, QFrame, QProgressBar,
    QFileDialog, QMessageBox, QScrollArea, QSizeGrip,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView
)
from PySide6.QtCore import Qt, QPoint, QTimer, Signal, QSettings
from PySide6.QtGui import QFont, QMouseEvent, QIcon, QColor, QPalette, QPainter

# ============================================================
# BMW Design System — Color Tokens
# https://getdesign.md/bmw/design-md
# ============================================================
BMW_BLUE         = "#1c69d4"
BMW_BLUE_ACTIVE  = "#0653b6"
BMW_INK          = "#262626"
BMW_BODY         = "#3c3c3c"
BMW_MUTED        = "#6b6b6b"
BMW_MUTED_SOFT   = "#9a9a9a"
BMW_HAIRLINE     = "#e6e6e6"
BMW_HAIRLINE_STR = "#cccccc"
BMW_CANVAS       = "#ffffff"
BMW_SURFACE_SOFT = "#f7f7f7"
BMW_SURFACE_CARD = "#fafafa"
BMW_SURFACE_DARK = "#1a2129"
BMW_ERROR        = "#dc2626"
BMW_SUCCESS      = "#22c55e"
BMW_ON_PRIMARY   = "#ffffff"
BMW_ON_DARK      = "#ffffff"
BMW_ON_DARK_SOFT = "#bbbbbb"

FONT_FAMILY = "'Inter', 'Microsoft YaHei UI', system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif"

# ============================================================
# 资源路径处理
# ============================================================
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ============================================================
# 启发式列识别与搜索核心逻辑
# ============================================================
def detect_columns_heuristically(rows_sample):
    if not rows_sample:
        return {"姓名": -1, "学号": -1, "班级": -1}

    num_cols = max(len(r) for r in rows_sample)
    score_name = [0] * num_cols
    score_id = [0] * num_cols
    score_class = [0] * num_cols

    for row in rows_sample:
        for idx, val in enumerate(row):
            if val is None:
                continue
            val_str = str(val).strip()
            if not val_str:
                continue

            if 7 <= len(val_str) <= 15 and val_str.isalnum() and any(c.isdigit() for c in val_str):
                score_id[idx] += 1

            if "班" in val_str or re.search(r'[一-龥]+\d+', val_str):
                score_class[idx] += 1

            if re.match(r'^[一-龥]{2,4}$', val_str):
                score_name[idx] += 1

    col_map = {"姓名": -1, "学号": -1, "班级": -1}
    assigned = set()

    best_id_idx = -1
    best_id_score = 0
    for idx in range(num_cols):
        if score_id[idx] > best_id_score:
            best_id_score = score_id[idx]
            best_id_idx = idx
    if best_id_idx != -1:
        col_map["学号"] = best_id_idx
        assigned.add(best_id_idx)

    best_class_idx = -1
    best_class_score = 0
    for idx in range(num_cols):
        if idx in assigned:
            continue
        if score_class[idx] > best_class_score:
            best_class_score = score_class[idx]
            best_class_idx = idx
    if best_class_idx != -1:
        col_map["班级"] = best_class_idx
        assigned.add(best_class_idx)

    best_name_idx = -1
    best_name_score = 0
    for idx in range(num_cols):
        if idx in assigned:
            continue
        if score_name[idx] > best_name_score:
            best_name_score = score_name[idx]
            best_name_idx = idx
    if best_name_idx != -1:
        col_map["姓名"] = best_name_idx
        assigned.add(best_name_idx)

    return col_map

def search_xlsx_professional(directory, search_string, progress_callback=None, is_cancelled=None, error_callback=None):
    results = []
    keywords = [k.strip().lower() for k in re.split(r'[,，\s]+', search_string) if k.strip()]
    if not keywords: return []

    all_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if (file.endswith(".xlsx") or file.endswith(".xlsm")) and not file.startswith("~$"):
                all_files.append(os.path.join(root, file))

    total_files = len(all_files)
    if total_files == 0:
        return []

    completed_files = 0
    lock = threading.Lock()

    def process_file(file_path):
        nonlocal completed_files
        fname = os.path.basename(file_path)
        file_results = []

        if is_cancelled and is_cancelled():
            return []

        try:
            workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True, keep_links=False)
            try:
                for sheet_name in workbook.sheetnames:
                    if is_cancelled and is_cancelled():
                        break
                    sheet = workbook[sheet_name]

                    row_iter = sheet.iter_rows(min_row=1, values_only=True)
                    sample_rows = []
                    for _ in range(10):
                        r = next(row_iter, None)
                        if r is None:
                            break
                        sample_rows.append(r)

                    if not sample_rows:
                        continue

                    col_map = {"姓名": -1, "学号": -1, "班级": -1}
                    first_row = sample_rows[0]
                    first_row_str = " ".join(str(v) for v in first_row if v is not None).lower()

                    is_header = any(k in first_row_str for k in ["姓名", "学生姓名", "name", "学号", "id", "学籍号", "班级", "class"])

                    if is_header:
                        for idx, val in enumerate(first_row):
                            v_str = str(val).lower() if val is not None else ""
                            if any(k in v_str for k in ["姓名", "学生姓名", "name"]): col_map["姓名"] = idx
                            elif any(k in v_str for k in ["学号", "id", "学籍号"]): col_map["学号"] = idx
                            elif any(k in v_str for k in ["班级", "class"]): col_map["班级"] = idx

                    if -1 in col_map.values():
                        heuristic_map = detect_columns_heuristically(sample_rows)
                        for k, v in col_map.items():
                            if v == -1:
                                col_map[k] = heuristic_map[k]

                    data_rows = sample_rows[1:] if is_header else sample_rows
                    consecutive_empty = 0

                    def process_row_tuple(row_tuple):
                        nonlocal consecutive_empty
                        if is_cancelled and is_cancelled():
                            return False

                        if not row_tuple or all(v is None or str(v).strip() == "" for v in row_tuple):
                            consecutive_empty += 1
                            if consecutive_empty >= 50:
                                return False
                            return True
                        else:
                            consecutive_empty = 0

                        row_values = [str(v).strip() if v is not None else "" for v in row_tuple]
                        row_str_combined = " ".join(row_values).lower()

                        if all(k in row_str_combined for k in keywords):
                            file_results.append({
                                "filename": fname,
                                "filepath": file_path,
                                "sheet": sheet_name,
                                "name": row_values[col_map["姓名"]] if col_map["姓名"] != -1 and col_map["姓名"] < len(row_values) else "未知",
                                "id": row_values[col_map["学号"]] if col_map["学号"] != -1 and col_map["学号"] < len(row_values) else "未记录",
                                "class": row_values[col_map["班级"]] if col_map["班级"] != -1 and col_map["班级"] < len(row_values) else "未记录"
                            })
                        return True

                    for r_t in data_rows:
                        if not process_row_tuple(r_t):
                            break
                    else:
                        for r_t in row_iter:
                            if not process_row_tuple(r_t):
                                break
            finally:
                workbook.close()
        except Exception as e:
            if error_callback:
                error_callback(file_path, str(e))
            else:
                print(f"Error reading {file_path}: {e}")

        with lock:
            completed_files += 1
            if progress_callback:
                progress_callback(completed_files, total_files, f"正在分析 ({completed_files}/{total_files}): {fname}")

        return file_results

    max_workers = min(8, os.cpu_count() or 4)
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(process_file, f) for f in all_files]
        for future in futures:
            if is_cancelled and is_cancelled():
                break
            results.extend(future.result())

    return results

# ============================================================
# 启动屏 — BMW Dark Hero
# ============================================================
class SplashScreen(QWidget):
    finished = Signal()

    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setFixedSize(450, 220)
        self.setStyleSheet(f"background: {BMW_SURFACE_DARK};")

        screen = QApplication.primaryScreen().geometry()
        self.move((screen.width() - 450) // 2, (screen.height() - 220) // 2)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(32, 32, 32, 32)
        layout.setSpacing(8)

        title = QLabel("XLSX 智能搜索助手")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet(f"font-family: {FONT_FAMILY}; font-size: 20px; font-weight: 700; color: {BMW_ON_DARK}; border: none;")

        author = QLabel("本软件由软工24-3 沈宇 开发完成")
        author.setAlignment(Qt.AlignmentFlag.AlignCenter)
        author.setStyleSheet(f"font-family: {FONT_FAMILY}; font-size: 13px; font-weight: 300; color: {BMW_ON_DARK_SOFT}; border: none;")

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFixedHeight(4)
        self.progress.setTextVisible(False)
        self.progress.setStyleSheet(f"""
            QProgressBar {{ background: rgba(255,255,255,0.12); border: none; }}
            QProgressBar::chunk {{ background: {BMW_BLUE}; border: none; }}
        """)

        layout.addStretch()
        layout.addWidget(title)
        layout.addWidget(author)
        layout.addSpacing(20)
        layout.addWidget(self.progress)
        layout.addStretch()

        self._tick_count = 0
        self._timer = QTimer()
        self._timer.timeout.connect(self._tick)
        self._timer.start(20)

    def _tick(self):
        self._tick_count += 1
        self.progress.setValue(self._tick_count)
        if self._tick_count >= 100:
            self._timer.stop()
            self.finished.emit()
            self.hide()
            self.deleteLater()

# ============================================================
# 自定义标题栏 — BMW Dark Surface
# ============================================================
class TitleBar(QFrame):
    def __init__(self, window):
        super().__init__()
        self.window = window
        self._drag_pos = None
        self.setFixedHeight(48)
        self.setStyleSheet(f"background: {BMW_SURFACE_DARK}; border: none;")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(20, 0, 0, 0)
        layout.setSpacing(0)

        title = QLabel("Excel 智能全文检索系统  By 沈宇")
        title.setStyleSheet(f"font-family: {FONT_FAMILY}; font-size: 14px; font-weight: 700; color: {BMW_ON_DARK}; border: none;")
        layout.addWidget(title)
        layout.addStretch()

        for text, action in [("—", "min"), ("□", "max"), ("✕", "close")]:
            btn = QPushButton(text)
            btn.setFixedSize(48, 32)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setStyleSheet(f"""
                QPushButton {{ background: transparent; color: {BMW_ON_DARK}; border: none; font-size: 14px; }}
                QPushButton:hover {{ background: {'#e22718' if action == 'close' else 'rgba(255,255,255,0.12)'}; }}
            """)
            if action == "min": btn.clicked.connect(window.showMinimized)
            elif action == "max": btn.clicked.connect(self._toggle_max)
            elif action == "close": btn.clicked.connect(window.close)
            layout.addWidget(btn)

    def _toggle_max(self):
        if self.window.isMaximized(): self.window.showNormal()
        else: self.window.showMaximized()

    def mousePressEvent(self, e: QMouseEvent):
        if e.button() == Qt.MouseButton.LeftButton:
            self._drag_pos = e.position().toPoint()

    def mouseMoveEvent(self, e: QMouseEvent):
        if self._drag_pos is not None and e.buttons() == Qt.MouseButton.LeftButton:
            self.window.move(e.globalPosition().toPoint() - self._drag_pos)

    def mouseReleaseEvent(self, e: QMouseEvent):
        self._drag_pos = None

# ============================================================
# Button Factories — BMW Rectangular 0px Radius
# ============================================================
def primary_button(text, height=48):
    btn = QPushButton(text)
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    btn.setFixedHeight(height)
    btn.setStyleSheet(f"""
        QPushButton {{
            font-family: {FONT_FAMILY};
            background: {BMW_BLUE}; color: {BMW_ON_PRIMARY}; border: none;
            border-radius: 0px; font-size: 14px; font-weight: 700;
            padding: 0 32px;
        }}
        QPushButton:hover {{ background: {BMW_BLUE_ACTIVE}; }}
        QPushButton:disabled {{ background: {BMW_HAIRLINE_STR}; color: {BMW_MUTED}; }}
    """)
    return btn

def secondary_button(text, height=48):
    btn = QPushButton(text)
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    btn.setFixedHeight(height)
    btn.setStyleSheet(f"""
        QPushButton {{
            font-family: {FONT_FAMILY};
            background: {BMW_CANVAS}; color: {BMW_INK}; border: 1px solid {BMW_HAIRLINE_STR};
            border-radius: 0px; font-size: 14px; font-weight: 700;
            padding: 0 24px;
        }}
        QPushButton:hover {{ border-color: {BMW_BLUE}; color: {BMW_BLUE}; }}
    """)
    return btn

def success_button(text, height=48):
    btn = QPushButton(text)
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    btn.setFixedHeight(height)
    btn.setStyleSheet(f"""
        QPushButton {{
            font-family: {FONT_FAMILY};
            background: {BMW_SUCCESS}; color: {BMW_CANVAS}; border: none;
            border-radius: 0px; font-size: 14px; font-weight: 700;
            padding: 0 32px;
        }}
        QPushButton:hover {{ background: #1ea84e; }}
        QPushButton:disabled {{ background: {BMW_HAIRLINE_STR}; color: {BMW_MUTED}; }}
    """)
    return btn

# ============================================================
# 联系作者 模态遮罩层 — BMW Style
# ============================================================
class ContactOverlay(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setVisible(False)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground, True)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)

        main_layout = QVBoxLayout(self)
        main_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.setContentsMargins(0, 0, 0, 0)

        card = QFrame()
        card.setFixedSize(340, 360)
        card.setStyleSheet(f"""
            QFrame {{
                background-color: {BMW_CANVAS};
                border: 1px solid {BMW_HAIRLINE};
            }}
            QLabel {{
                border: none;
                background: transparent;
            }}
        """)

        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(24, 24, 24, 24)
        card_layout.setSpacing(16)

        header_layout = QHBoxLayout()
        title = QLabel("联系作者")
        title.setStyleSheet(f"font-family: {FONT_FAMILY}; font-size: 16px; font-weight: 700; color: {BMW_INK};")
        close_btn = QPushButton("✕")
        close_btn.setFixedSize(28, 28)
        close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        close_btn.setStyleSheet(f"""
            QPushButton {{ background: transparent; color: {BMW_MUTED}; border: none; font-size: 16px; font-weight: 700; }}
            QPushButton:hover {{ color: {BMW_ERROR}; }}
        """)
        close_btn.clicked.connect(self.hide)
        header_layout.addWidget(title)
        header_layout.addStretch()
        header_layout.addWidget(close_btn)
        card_layout.addLayout(header_layout)

        desc = QLabel("如有问题或建议，欢迎扫码添加作者微信（沈宇）：")
        desc.setWordWrap(True)
        desc.setStyleSheet(f"font-family: {FONT_FAMILY}; font-size: 14px; font-weight: 300; color: {BMW_BODY}; line-height: 1.55;")
        card_layout.addWidget(desc)

        lbl = QLabel()
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        pix = QIcon(resource_path("2.png")).pixmap(200, 200)
        lbl.setPixmap(pix)
        card_layout.addWidget(lbl)

        card_layout.addStretch()
        main_layout.addWidget(card)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor(0, 0, 0, 100))
        super().paintEvent(event)

    def resize_to_parent(self):
        if self.parent:
            self.setGeometry(0, 0, self.parent.width(), self.parent.height())

# ============================================================
# 主窗口
# ============================================================
class ExcelSearchApp(QWidget):
    search_progress_sig = Signal(int, int, str)
    search_finished_sig = Signal(list)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel 智能搜索助手 — By 沈宇")
        self.resize(1020, 820)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False)

        self.last_results = []
        self.searching = False
        self.cancel_flag = False
        self.failed_files = []

        self._setup_ui()
        self.load_settings()

        self.contact_overlay = ContactOverlay(self)

        self.search_progress_sig.connect(self._on_progress)
        self.search_finished_sig.connect(self._on_finished)

    def load_settings(self):
        settings = QSettings("ShenYu", "ExcelSearchApp")
        saved_dir = settings.value("search_directory", "")
        saved_kw = settings.value("search_keyword", "")
        if saved_dir:
            self.dir_path.setText(saved_dir)
        if saved_kw:
            self.search_input.setText(saved_kw)

    def save_settings(self):
        settings = QSettings("ShenYu", "ExcelSearchApp")
        settings.setValue("search_directory", self.dir_path.text().strip())
        settings.setValue("search_keyword", self.search_input.text().strip())

    def _setup_ui(self):
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        self.title_bar = TitleBar(self)
        root_layout.addWidget(self.title_bar)

        body = QWidget()
        body.setStyleSheet(f"background: {BMW_SURFACE_SOFT};")
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(32, 32, 32, 32)
        body_layout.setSpacing(24)

        # ---- 搜索卡片 ----
        card1 = QFrame()
        card1.setStyleSheet(f"background: {BMW_CANVAS}; border: 1px solid {BMW_HAIRLINE};")
        c1 = QVBoxLayout(card1)
        c1.setContentsMargins(24, 24, 24, 24)
        c1.setSpacing(16)

        path_row = QHBoxLayout()
        path_row.setSpacing(12)
        path_label = QLabel("搜索路径")
        path_label.setStyleSheet(f"font-family: {FONT_FAMILY}; font-size: 14px; font-weight: 700; color: {BMW_INK}; border: none; background: transparent;")
        path_label.setFixedWidth(64)
        self.dir_path = QLineEdit()
        self.dir_path.setPlaceholderText("选择包含 .xlsx 文件的文件夹...")
        self.dir_path.setFixedHeight(48)
        self.dir_path.setStyleSheet(f"""
            font-family: {FONT_FAMILY}; font-size: 14px; font-weight: 300;
            color: {BMW_INK}; background: {BMW_CANVAS};
            border: 1px solid {BMW_HAIRLINE}; border-radius: 0px;
            padding: 0 16px;
        """)
        browse_btn = secondary_button("浏览")
        path_row.addWidget(path_label)
        path_row.addWidget(self.dir_path)
        path_row.addWidget(browse_btn)
        c1.addLayout(path_row)

        kw_row = QHBoxLayout()
        kw_row.setSpacing(12)
        kw_label = QLabel("搜索内容")
        kw_label.setStyleSheet(f"font-family: {FONT_FAMILY}; font-size: 14px; font-weight: 700; color: {BMW_INK}; border: none; background: transparent;")
        kw_label.setFixedWidth(64)
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("输入关键字，多个关键字用空格或逗号分隔...")
        self.search_input.setFixedHeight(48)
        self.search_input.setStyleSheet(f"""
            font-family: {FONT_FAMILY}; font-size: 14px; font-weight: 300;
            color: {BMW_INK}; background: {BMW_CANVAS};
            border: 1px solid {BMW_HAIRLINE}; border-radius: 0px;
            padding: 0 16px;
        """)
        self.search_input.returnPressed.connect(self.start_search)
        kw_row.addWidget(kw_label)
        kw_row.addWidget(self.search_input)
        c1.addLayout(kw_row)

        hint = QLabel("专业可能有缩写，建议搜索简写以及完整称呼进行两次搜索以确保准确")
        hint.setStyleSheet(f"font-family: {FONT_FAMILY}; font-size: 12px; font-weight: 400; color: {BMW_MUTED}; border: none; background: transparent;")
        c1.addWidget(hint)
        body_layout.addWidget(card1)

        # ---- 按钮栏 ----
        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)
        self.search_btn = primary_button("开始分析搜索", height=48)
        self.search_btn.clicked.connect(self.start_search)
        self.export_btn = success_button("导出汇总表格", height=48)
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export_results)

        btn_row.addWidget(self.search_btn)
        btn_row.addWidget(self.export_btn)
        contact_btn = secondary_button("联系作者", height=48)
        contact_btn.clicked.connect(self.show_contact)
        btn_row.addWidget(contact_btn)
        btn_row.addStretch()
        body_layout.addLayout(btn_row)

        # ---- 进度 ----
        self.prog_widget = QWidget()
        self.prog_widget.setStyleSheet("background: transparent; border: none;")
        pl = QVBoxLayout(self.prog_widget)
        pl.setContentsMargins(0, 0, 0, 0)
        pl.setSpacing(8)
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(4)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{ background: {BMW_CANVAS}; border: none; }}
            QProgressBar::chunk {{ background: {BMW_BLUE}; border: none; }}
        """)
        self.prog_label = QLabel("等待开始...")
        self.prog_label.setStyleSheet(f"font-family: {FONT_FAMILY}; font-size: 12px; font-weight: 400; color: {BMW_MUTED}; border: none; background: transparent;")
        pl.addWidget(self.progress_bar)
        pl.addWidget(self.prog_label)
        self.prog_widget.hide()
        body_layout.addWidget(self.prog_widget)

        # ---- 结果卡片 ----
        card2 = QFrame()
        card2.setStyleSheet(f"background: {BMW_CANVAS}; border: 1px solid {BMW_HAIRLINE};")
        c2 = QVBoxLayout(card2)
        c2.setContentsMargins(1, 1, 1, 1)
        c2.setSpacing(0)

        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["文件名", "姓名", "学号", "班级", "操作"])
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setShowGrid(False)

        self.table.setStyleSheet(f"""
            QTableWidget {{
                font-family: {FONT_FAMILY};
                background-color: {BMW_CANVAS};
                alternate-background-color: {BMW_SURFACE_SOFT};
                border: none;
                color: {BMW_INK};
                font-size: 14px;
                font-weight: 300;
            }}
            QTableWidget::item {{
                padding: 10px 12px;
                border-bottom: 1px solid {BMW_HAIRLINE};
            }}
            QTableWidget::item:selected {{
                background-color: #e8eef6;
                color: {BMW_INK};
            }}
            QHeaderView::section {{
                font-family: {FONT_FAMILY};
                background-color: {BMW_SURFACE_SOFT};
                color: {BMW_INK};
                font-weight: 700;
                font-size: 13px;
                border: none;
                border-bottom: 2px solid {BMW_HAIRLINE_STR};
                height: 44px;
                padding-left: 12px;
                text-transform: uppercase;
                letter-spacing: 1px;
            }}
            QScrollBar:vertical {{
                border: none;
                background: {BMW_SURFACE_SOFT};
                width: 8px;
            }}
            QScrollBar::handle:vertical {{
                background: {BMW_HAIRLINE_STR};
                min-height: 20px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: {BMW_MUTED};
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                border: none;
                background: none;
                height: 0px;
            }}
        """)

        table_header = self.table.horizontalHeader()
        table_header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        table_header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        table_header.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        table_header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        table_header.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(1, 100)
        self.table.setColumnWidth(2, 140)
        self.table.setColumnWidth(3, 140)
        self.table.setColumnWidth(4, 90)
        table_header.setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.table.verticalHeader().setVisible(False)
        self.table.cellDoubleClicked.connect(self._on_cell_double_clicked)

        # 空状态
        self.empty_widget = QWidget()
        self.empty_widget.setStyleSheet(f"border: none; background: {BMW_CANVAS};")
        el = QVBoxLayout(self.empty_widget)
        el.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon = QLabel("—")
        icon.setStyleSheet(f"font-family: {FONT_FAMILY}; font-size: 48px; font-weight: 300; color: {BMW_HAIRLINE_STR}; border: none; background: transparent;")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        txt = QLabel("暂无搜索结果")
        txt.setAlignment(Qt.AlignmentFlag.AlignCenter)
        txt.setStyleSheet(f"font-family: {FONT_FAMILY}; font-size: 14px; font-weight: 300; color: {BMW_MUTED}; border: none; background: transparent;")
        el.addWidget(icon)
        el.addWidget(txt)

        c2.addWidget(self.table)
        c2.addWidget(self.empty_widget)

        self.table.hide()
        self.empty_widget.show()

        body_layout.addWidget(card2, 1)
        root_layout.addWidget(body, 1)

        # ---- 状态栏 ----
        status = QFrame()
        status.setFixedHeight(36)
        status.setStyleSheet(f"background: {BMW_CANVAS}; border: none; border-top: 1px solid {BMW_HAIRLINE};")
        sl = QHBoxLayout(status)
        sl.setContentsMargins(20, 0, 20, 0)
        self.status_label = QLabel("系统就绪")
        self.status_label.setStyleSheet(f"font-family: {FONT_FAMILY}; font-size: 12px; font-weight: 400; color: {BMW_MUTED}; border: none; background: transparent;")
        sl.addWidget(self.status_label)
        sl.addStretch()

        sizegrip = QSizeGrip(self)
        sizegrip.setStyleSheet("background: transparent;")
        sl.addWidget(sizegrip)

        root_layout.addWidget(status)

    def browse(self):
        d = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if d: self.dir_path.setText(d)

    def start_search(self):
        if self.searching:
            self._cancel_search()
            return

        folder = self.dir_path.text().strip()
        kw = self.search_input.text().strip()
        if not folder or not kw: return

        if not os.path.isdir(folder):
            QMessageBox.warning(self, "错误", "填写的搜索路径不是有效的文件夹目录！")
            return

        self.save_settings()

        self.searching = True
        self.cancel_flag = False
        self.failed_files = []

        self.table.setRowCount(0)
        self.table.hide()
        self.empty_widget.hide()

        self.search_btn.setText("停止搜索")
        self.search_btn.setStyleSheet(f"""
            QPushButton {{
                font-family: {FONT_FAMILY};
                background: {BMW_ERROR}; color: {BMW_CANVAS}; border: none;
                border-radius: 0px; font-size: 14px; font-weight: 700;
                padding: 0 32px;
            }}
            QPushButton:hover {{ background: #b91c1c; }}
        """)

        self.export_btn.setEnabled(False)
        self.prog_widget.show()
        self.status_label.setText("正在准备搜索...")

        threading.Thread(target=self._worker, args=(folder, kw), daemon=True).start()

    def _cancel_search(self):
        self.cancel_flag = True
        self.status_label.setText("正在停止搜索...")
        self.search_btn.setEnabled(False)

    def _worker(self, folder, kw):
        def cb(c, t, m): self.search_progress_sig.emit(c, t, m)
        def cancel_check(): return self.cancel_flag
        def error_cb(path, err): self.failed_files.append((path, err))

        res = search_xlsx_professional(folder, kw, cb, cancel_check, error_cb)
        self.search_finished_sig.emit(res)

    def _on_progress(self, c, t, msg):
        self.progress_bar.setMaximum(t)
        self.progress_bar.setValue(c)
        self.prog_label.setText(msg)

    def _on_finished(self, results):
        self.searching = False
        self.search_btn.setEnabled(True)

        self.search_btn.setText("开始分析搜索")
        self.search_btn.setStyleSheet(f"""
            QPushButton {{
                font-family: {FONT_FAMILY};
                background: {BMW_BLUE}; color: {BMW_ON_PRIMARY}; border: none;
                border-radius: 0px; font-size: 14px; font-weight: 700;
                padding: 0 32px;
            }}
            QPushButton:hover {{ background: {BMW_BLUE_ACTIVE}; }}
            QPushButton:disabled {{ background: {BMW_HAIRLINE_STR}; color: {BMW_MUTED}; }}
        """)

        self.prog_widget.hide()

        if self.cancel_flag:
            self.status_label.setText(f"搜索已中止 | 找到 {len(results)} 条结果")
        else:
            self.status_label.setText(f"完成 | 找到 {len(results)} 条结果")

        self.last_results = results

        if results:
            self.export_btn.setEnabled(True)
            self._render_results()
        else:
            self.table.hide()
            self.empty_widget.show()

        if self.failed_files:
            err_msg = f"搜寻完成，但有 {len(self.failed_files)} 个文件读取失败：\n"
            for fpath, err in self.failed_files[:5]:
                err_msg += f"- {os.path.basename(fpath)}: {err}\n"
            if len(self.failed_files) > 5:
                err_msg += "... 等等\n"
            QMessageBox.warning(self, "部分文件读取失败", err_msg)

        if self.cancel_flag:
            QMessageBox.information(self, "已中止", f"搜索已中止！共保留 {len(results)} 条结果。")
        else:
            QMessageBox.information(self, "完成", f"搜索完成！共找到 {len(results)} 条结果。")

    def _render_results(self):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        self.table.show()
        self.empty_widget.hide()

        limit = 1000
        results_to_show = self.last_results[:limit]
        self.table.setRowCount(len(results_to_show))

        for idx, item in enumerate(results_to_show):
            file_item = QTableWidgetItem(item['filename'])
            file_item.setToolTip(item['filepath'])
            self.table.setItem(idx, 0, file_item)

            self.table.setItem(idx, 1, QTableWidgetItem(item['name']))
            self.table.setItem(idx, 2, QTableWidgetItem(item['id']))
            self.table.setItem(idx, 3, QTableWidgetItem(item['class']))

            view_btn = QPushButton("查看")
            view_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            view_btn.setFixedHeight(28)
            view_btn.setStyleSheet(f"""
                QPushButton {{
                    font-family: {FONT_FAMILY};
                    background: {BMW_BLUE}; color: {BMW_ON_PRIMARY}; border: none;
                    border-radius: 0px; font-size: 12px; font-weight: 700;
                    padding: 0 16px;
                }}
                QPushButton:hover {{ background: {BMW_BLUE_ACTIVE}; }}
            """)
            view_btn.clicked.connect(lambda checked=False, p=item['filepath']: self._open_file(p))

            container = QWidget()
            container.setStyleSheet("background: transparent;")
            layout = QHBoxLayout(container)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(view_btn)
            self.table.setCellWidget(idx, 4, container)

        self.table.setSortingEnabled(True)

    def _on_cell_double_clicked(self, row, column):
        if row >= 0 and row < len(self.last_results):
            file_path = self.last_results[row]['filepath']
            self._open_file(file_path)

    def _open_file(self, filepath):
        if os.path.exists(filepath):
            if os.name == 'nt':
                os.startfile(filepath)
            else:
                subprocess.call(['open', filepath])
        else:
            QMessageBox.warning(self, "错误", f"文件不存在或已被移动：\n{filepath}")

    def show_contact(self):
        self.contact_overlay.resize_to_parent()
        self.contact_overlay.show()
        self.contact_overlay.raise_()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "contact_overlay"):
            self.contact_overlay.resize_to_parent()

    def export_results(self):
        path, _ = QFileDialog.getSaveFileName(self, "导出结果", "汇总.xlsx", "Excel (*.xlsx)")
        if path:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["文件名", "路径", "姓名", "学号", "班级"])
            for i in self.last_results: ws.append([i['filename'], i['filepath'], i['name'], i['id'], i['class']])
            wb.save(path)
            QMessageBox.information(self, "成功", "导出完成")



if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(f"QWidget {{ font-family: '{FONT_FAMILY}'; color: {BMW_INK}; }} QLineEdit {{ font-family: '{FONT_FAMILY}'; border: 1px solid {BMW_HAIRLINE}; border-radius: 0px; padding: 0 16px; }}")
    window = ExcelSearchApp()
    splash = SplashScreen()
    splash.finished.connect(window.show)
    splash.show()
    sys.exit(app.exec())
