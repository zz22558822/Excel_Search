# -*- coding: utf-8 -*-
"""
套件依賴安裝
pip install PyQt6 pandas openpyxl xlrd
"""
from __future__ import annotations

import html
import json
import os
import re
import sys
import time
from dataclasses import dataclass
from functools import reduce
from pathlib import Path

import pandas as pd
from PyQt6.QtCore import (
    QEasingCurve, QPropertyAnimation, QRectF, QSettings, QThread, QTimer, QUrl,
    Qt, pyqtProperty, pyqtSignal,
)
from PyQt6.QtGui import QColor, QDesktopServices, QFont, QPainter
from PyQt6.QtWidgets import (
    QAbstractButton, QAbstractItemView, QApplication, QComboBox, QFileDialog,
    QFrame, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QListWidget,
    QMainWindow, QMessageBox, QProgressBar, QPushButton, QScrollArea,
    QTableWidget, QTableWidgetItem, QTabWidget, QTextEdit, QVBoxLayout, QWidget,
)

APP_NAME = "Excel 關鍵字搜尋"
DEFAULT_SOURCE = r"C:\Users\User\Desktop"
DEFAULT_OUTPUT = str(Path(__file__).resolve().parent / "Excel_Search_Results.xlsx")
DEFAULT_KEYWORDS = ["Google Chrome", "Microsoft Edge"]
ALL_EXTS = (".csv", ".xlsx", ".xls")

# ───────────────────────────── 主題 ─────────────────────────────
THEMES = {
    "dark": dict(
        bg="#0E1016", card="#171A23", input="#1E222D", border="#2B3040",
        text="#E8EAF0", muted="#8B93A7", accent="#5B7CFA",
        accent_hover="#7490FF", accent_press="#4A69E6",
        sel="rgba(91,124,250,0.20)", hover="rgba(255,255,255,0.05)",
        ok="#3DD68C", err="#FF6B6B", warn="#F5B94D", switch_off="#3A4052",
    ),
    "light": dict(
        bg="#F3F5FA", card="#FFFFFF", input="#F1F3F9", border="#E1E5EF",
        text="#1B1F33", muted="#6B7390", accent="#4A66E8",
        accent_hover="#5E79F0", accent_press="#3B55D1",
        sel="rgba(74,102,232,0.14)", hover="rgba(0,0,0,0.04)",
        ok="#1FA866", err="#E5484D", warn="#C98A12", switch_off="#C9CFDD",
    ),
}
_CURRENT = {"t": THEMES["dark"]}


def current_theme() -> dict:
    return _CURRENT["t"]


QSS = """
QWidget { color: %(text)s; font-size: 13px; }
QMainWindow, QWidget#Root { background: %(bg)s; }
QScrollArea { border: none; background: transparent; }
QScrollArea > QWidget > QWidget { background: transparent; }
QLabel { background: transparent; }
QLabel#Title { font-size: 24px; font-weight: 700; }
QLabel#Subtitle, QLabel#Muted { color: %(muted)s; }
QLabel#CardTitle { font-size: 14px; font-weight: 700; }
QLabel#RowTitle { font-weight: 600; }
QLabel#StatValue { font-size: 24px; font-weight: 700; }
QLabel#Path { color: %(muted)s; font-size: 12px; }

QFrame#Card { background: %(card)s; border: 1px solid %(border)s; border-radius: 14px; }
QFrame#Divider { background: %(border)s; max-height: 1px; min-height: 1px; border: none; }

QLineEdit, QComboBox {
    background: %(input)s; border: 1px solid %(border)s; border-radius: 10px;
    padding: 9px 12px; selection-background-color: %(accent)s; selection-color: white;
}
QLineEdit:focus, QComboBox:focus { border: 1px solid %(accent)s; }
QLineEdit:disabled, QComboBox:disabled { color: %(muted)s; }
QComboBox::drop-down { border: none; width: 28px; }
QComboBox QAbstractItemView {
    background: %(card)s; border: 1px solid %(border)s; border-radius: 8px; padding: 4px;
    selection-background-color: %(sel)s; selection-color: %(text)s; outline: 0;
}

QPushButton {
    background: %(input)s; border: 1px solid %(border)s; border-radius: 10px;
    padding: 9px 16px; font-weight: 600;
}
QPushButton:hover { border-color: %(accent)s; }
QPushButton:pressed { background: %(sel)s; }
QPushButton:disabled { color: %(muted)s; border-color: %(border)s; }
QPushButton#Primary {
    background: %(accent)s; border: none; color: white; padding: 13px 20px; font-size: 14px;
}
QPushButton#Primary:hover { background: %(accent_hover)s; }
QPushButton#Primary:pressed { background: %(accent_press)s; }
QPushButton#Primary:disabled { background: %(border)s; color: %(muted)s; }
QPushButton#Danger { color: %(err)s; padding: 13px 20px; font-size: 14px; }
QPushButton#Danger:hover { border-color: %(err)s; }
QPushButton#Danger:disabled { color: %(muted)s; }
QPushButton#Ghost { background: transparent; }
QPushButton#Chip {
    background: %(input)s; border: 1px solid %(border)s; border-radius: 16px;
    padding: 6px 16px; font-weight: 600; color: %(muted)s;
}
QPushButton#Chip:checked { background: %(sel)s; border-color: %(accent)s; color: %(accent)s; }

QListWidget {
    background: %(input)s; border: 1px solid %(border)s; border-radius: 10px;
    padding: 4px; outline: 0;
}
QListWidget::item { padding: 7px 10px; border-radius: 7px; }
QListWidget::item:hover { background: %(hover)s; }
QListWidget::item:selected { background: %(sel)s; color: %(text)s; }

QProgressBar {
    background: %(input)s; border: none; border-radius: 5px;
    min-height: 10px; max-height: 10px; color: transparent;
}
QProgressBar::chunk { background: %(accent)s; border-radius: 5px; }

QTabWidget::pane { border: none; }
QTabBar::tab {
    background: transparent; color: %(muted)s; padding: 8px 16px; margin-right: 4px;
    border-bottom: 2px solid transparent; font-weight: 600;
}
QTabBar::tab:selected { color: %(text)s; border-bottom: 2px solid %(accent)s; }
QTabBar::tab:hover { color: %(text)s; }

QTextEdit, QTableWidget {
    background: %(input)s; border: 1px solid %(border)s; border-radius: 10px;
    padding: 6px; gridline-color: %(border)s; selection-background-color: %(sel)s;
    selection-color: %(text)s;
}
QTableWidget { alternate-background-color: %(hover)s; }
QTableWidget::item { padding: 4px 6px; }
QHeaderView::section {
    background: %(card)s; color: %(muted)s; border: none;
    border-bottom: 1px solid %(border)s; padding: 8px; font-weight: 600;
}
QTableCornerButton::section { background: %(card)s; border: none; }

QScrollBar:vertical { background: transparent; width: 10px; margin: 2px; }
QScrollBar::handle:vertical { background: %(border)s; border-radius: 4px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: %(muted)s; }
QScrollBar:horizontal { background: transparent; height: 10px; margin: 2px; }
QScrollBar::handle:horizontal { background: %(border)s; border-radius: 4px; min-width: 30px; }
QScrollBar::handle:horizontal:hover { background: %(muted)s; }
QScrollBar::add-line, QScrollBar::sub-line { width: 0; height: 0; }
QScrollBar::add-page, QScrollBar::sub-page { background: transparent; }

QToolTip {
    background: %(card)s; color: %(text)s; border: 1px solid %(border)s;
    border-radius: 6px; padding: 6px 8px;
}
"""


# ───────────────────────────── 自訂元件 ─────────────────────────────
class Switch(QAbstractButton):
    """帶動畫的開關。"""

    def __init__(self, checked: bool = False, parent=None):
        super().__init__(parent)
        self._pos = 1.0 if checked else 0.0
        self.setCheckable(True)
        self.setChecked(checked)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedSize(46, 26)
        self._anim = QPropertyAnimation(self, b"position", self)
        self._anim.setDuration(160)
        self._anim.setEasingCurve(QEasingCurve.Type.InOutCubic)
        self.toggled.connect(self._animate)

    def _animate(self, on: bool):
        self._anim.stop()
        self._anim.setStartValue(self._pos)
        self._anim.setEndValue(1.0 if on else 0.0)
        self._anim.start()

    def getPosition(self) -> float:
        return self._pos

    def setPosition(self, v: float):
        self._pos = v
        self.update()

    position = pyqtProperty(float, getPosition, setPosition)

    def paintEvent(self, _):
        t = current_theme()
        off, on = QColor(t["switch_off"]), QColor(t["accent"])
        k = self._pos
        c = QColor(
            int(off.red() + (on.red() - off.red()) * k),
            int(off.green() + (on.green() - off.green()) * k),
            int(off.blue() + (on.blue() - off.blue()) * k),
        )
        if not self.isEnabled():
            c.setAlpha(110)
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        w, h = self.width(), self.height()
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(c)
        p.drawRoundedRect(QRectF(0, 0, w, h), h / 2, h / 2)
        d = h - 6
        x = 3 + (w - d - 6) * k
        p.setBrush(QColor("#FFFFFF"))
        p.drawEllipse(QRectF(x, 3, d, d))


class DropLineEdit(QLineEdit):
    """可把資料夾 / 檔案直接拖進來的輸入框。"""

    def __init__(self, mode: str, parent=None):
        super().__init__(parent)
        self.mode = mode  # "dir" 或 "file"
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        urls = e.mimeData().urls()
        if not urls:
            return
        p = urls[0].toLocalFile()
        if self.mode == "dir" and os.path.isfile(p):
            p = os.path.dirname(p)
        self.setText(os.path.normpath(p))


def make_card(title: str, subtitle: str | None = None):
    card = QFrame()
    card.setObjectName("Card")
    lay = QVBoxLayout(card)
    lay.setContentsMargins(18, 16, 18, 18)
    lay.setSpacing(12)
    head = QVBoxLayout()
    head.setSpacing(2)
    t = QLabel(title)
    t.setObjectName("CardTitle")
    head.addWidget(t)
    if subtitle:
        s = QLabel(subtitle)
        s.setObjectName("Muted")
        s.setWordWrap(True)
        head.addWidget(s)
    lay.addLayout(head)
    return card, lay


class SettingRow(QWidget):
    def __init__(self, title: str, desc: str, control: QWidget):
        super().__init__()
        h = QHBoxLayout(self)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(12)
        v = QVBoxLayout()
        v.setSpacing(2)
        a = QLabel(title)
        a.setObjectName("RowTitle")
        b = QLabel(desc)
        b.setObjectName("Muted")
        b.setWordWrap(True)
        v.addWidget(a)
        v.addWidget(b)
        h.addLayout(v, 1)
        h.addWidget(control, 0, Qt.AlignmentFlag.AlignVCenter)
        self.control = control


class StatCard(QFrame):
    def __init__(self, title: str, value: str = "—"):
        super().__init__()
        self.setObjectName("Card")
        v = QVBoxLayout(self)
        v.setContentsMargins(16, 14, 16, 14)
        v.setSpacing(2)
        t = QLabel(title)
        t.setObjectName("Muted")
        self.value = QLabel(value)
        self.value.setObjectName("StatValue")
        v.addWidget(t)
        v.addWidget(self.value)

    def set(self, text: str):
        self.value.setText(text)


# ───────────────────────────── 搜尋邏輯 ─────────────────────────────
@dataclass
class SearchConfig:
    source: str
    output: str
    keywords: list
    case_sensitive: bool
    whole: bool
    whole_mode: str      # "cell" 整個儲存格相同 / "word" 完整單字
    match_all: bool      # True: 全部關鍵字都要出現;False: 任一個即可
    recursive: bool
    exts: tuple
    skip_hidden: bool


def fmt_seconds(sec: float) -> str:
    if sec < 60:
        return f"{sec:.1f} 秒"
    m, s = divmod(int(sec), 60)
    return f"{m} 分 {s} 秒"


class SearchWorker(QThread):
    progress = pyqtSignal(int, int, str)
    stats = pyqtSignal(int, int, int)   # 已掃描檔案, 命中檔案, 命中筆數
    log = pyqtSignal(str, str)          # 等級, 訊息
    done = pyqtSignal(dict)
    failed = pyqtSignal(str)

    def __init__(self, cfg: SearchConfig):
        super().__init__()
        self.cfg = cfg
        self._cancel = False

    def cancel(self):
        self._cancel = True

    # --- 檔案蒐集 ---
    def _collect(self) -> list:
        cfg = self.cfg
        out_abs = os.path.abspath(cfg.output).lower()
        files = []
        for root, dirs, names in os.walk(cfg.source):
            if cfg.skip_hidden:
                dirs[:] = [d for d in dirs if not d.startswith(".")]
            for n in names:
                ext = os.path.splitext(n)[1].lower()
                if ext not in cfg.exts or n.startswith("~$"):
                    continue
                if cfg.skip_hidden and n.startswith("."):
                    continue
                full = os.path.join(root, n)
                if os.path.abspath(full).lower() == out_abs:
                    continue  # 不要掃到自己的輸出檔
                files.append(full)
            if not cfg.recursive:
                break
        return sorted(files)

    # --- 讀檔 (回傳 [(工作表名稱或 None, DataFrame)]) ---
    @staticmethod
    def _read(path: str) -> list:
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            last = None
            for enc in ("utf-8-sig", "cp950", "gb18030"):
                try:
                    return [(None, pd.read_csv(path, dtype=str, encoding=enc))]
                except UnicodeDecodeError as e:
                    last = e
            raise last
        xl = pd.ExcelFile(path)
        return [(s, pd.read_excel(xl, sheet_name=s, dtype=str)) for s in xl.sheet_names]

    # --- 比對 ---
    def _row_mask(self, df: pd.DataFrame) -> pd.Series:
        cfg = self.cfg
        clean = df.fillna("").astype(str)
        flags = 0 if cfg.case_sensitive else re.IGNORECASE
        masks = []

        if cfg.whole and cfg.whole_mode == "cell":
            cells = clean.apply(lambda c: c.str.strip())
            if not cfg.case_sensitive:
                cells = cells.apply(lambda c: c.str.casefold())
            for kw in cfg.keywords:
                k = kw.strip() if cfg.case_sensitive else kw.strip().casefold()
                masks.append((cells == k).any(axis=1))
        else:
            # 把整列串成一段文字再搜尋
            text = clean.iloc[:, 0]
            for i in range(1, clean.shape[1]):
                text = text + " " + clean.iloc[:, i]
            for kw in cfg.keywords:
                esc = re.escape(kw)
                pat = rf"(?<!\w){esc}(?!\w)" if cfg.whole else esc
                masks.append(text.str.contains(pat, flags=flags, regex=True, na=False))

        op = (lambda a, b: a & b) if cfg.match_all else (lambda a, b: a | b)
        return reduce(op, masks)

    def run(self):
        try:
            t0 = time.time()
            cfg = self.cfg
            files = self._collect()
            total = len(files)
            self.progress.emit(0, total, "")
            if total == 0:
                self.log.emit("warn", "資料夾內沒有符合副檔名的檔案。")
                self.done.emit(dict(cancelled=False, scanned=0, matched_files=0, rows=0,
                                    elapsed=time.time() - t0, output=None, preview=None))
                return
            self.log.emit("info", f"找到 {total} 個檔案,開始搜尋…")

            matches, scanned, matched_files, rows = [], 0, 0, 0
            for i, path in enumerate(files, 1):
                if self._cancel:
                    break
                self.progress.emit(i - 1, total, path)
                try:
                    hit_parts = []
                    for sheet, df in self._read(path):
                        if df.empty or df.shape[1] == 0:
                            continue
                        part = df[self._row_mask(df)].copy()
                        if part.empty:
                            continue
                        if sheet is not None:
                            part["來源工作表"] = sheet
                        hit_parts.append(part)
                    if hit_parts:
                        m = pd.concat(hit_parts, ignore_index=True)
                        m.insert(0, "檔案路徑", path)
                        m.insert(1, "檔案名稱", os.path.basename(path))
                        matches.append(m)
                        matched_files += 1
                        rows += len(m)
                        self.log.emit("ok", f"找到符合項目  {path}  ({len(m)} 筆)")
                except Exception as e:  # noqa: BLE001
                    self.log.emit("err", f"讀取失敗  {path}: {e}")
                scanned += 1
                self.stats.emit(scanned, matched_files, rows)
            self.progress.emit(scanned, total, "")

            if self._cancel:
                self.log.emit("warn", "已取消,未寫入結果。")
                self.done.emit(dict(cancelled=True, scanned=scanned, matched_files=matched_files,
                                    rows=rows, elapsed=time.time() - t0, output=None, preview=None))
                return

            if not matches:
                self.log.emit("warn", "未找到包含指定關鍵字的資料。")
                self.done.emit(dict(cancelled=False, scanned=scanned, matched_files=0, rows=0,
                                    elapsed=time.time() - t0, output=None, preview=None))
                return

            final = pd.concat(matches, ignore_index=True)
            out = cfg.output
            os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
            try:
                if out.lower().endswith(".csv"):
                    final.to_csv(out, index=False, encoding="utf-8-sig")
                else:
                    final.to_excel(out, index=False)
            except PermissionError:
                self.failed.emit(f"無法寫入「{out}」。\n檔案可能正被 Excel 開啟,請先關閉後再試。")
                return
            self.log.emit("ok", f"搜尋完成,結果已寫入:{out}")
            self.done.emit(dict(cancelled=False, scanned=scanned, matched_files=matched_files,
                                rows=rows, elapsed=time.time() - t0, output=out,
                                preview=final.head(500)))
        except Exception as e:  # noqa: BLE001
            self.failed.emit(f"發生未預期的錯誤:{e}")


# ───────────────────────────── 主視窗 ─────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self, app: QApplication):
        super().__init__()
        self.app = app
        self.settings = QSettings("ExcelSearch", "GUI")
        self.worker: SearchWorker | None = None
        self.result_path: str | None = None
        self.t0 = 0.0
        self.theme_name = "dark"

        self.setWindowTitle(APP_NAME)
        self.resize(1180, 780)
        self.setMinimumSize(1020, 680)

        self._build_ui()
        self._load_settings()
        self.apply_theme(self.theme_name)

        self.timer = QTimer(self)
        self.timer.setInterval(200)
        self.timer.timeout.connect(self._tick)

    # ----- UI -----
    def _build_ui(self):
        root = QWidget()
        root.setObjectName("Root")
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(28, 22, 28, 24)
        outer.setSpacing(18)

        # 標題列
        head = QHBoxLayout()
        tv = QVBoxLayout()
        tv.setSpacing(2)
        title = QLabel(APP_NAME)
        title.setObjectName("Title")
        sub = QLabel("掃描資料夾中的 Excel / CSV,把含有關鍵字的整列資料匯出成一份結果")
        sub.setObjectName("Subtitle")
        tv.addWidget(title)
        tv.addWidget(sub)
        head.addLayout(tv, 1)
        self.theme_btn = QPushButton()
        self.theme_btn.setObjectName("Ghost")
        self.theme_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.theme_btn.clicked.connect(self.toggle_theme)
        head.addWidget(self.theme_btn, 0, Qt.AlignmentFlag.AlignTop)
        outer.addLayout(head)

        body = QHBoxLayout()
        body.setSpacing(18)
        outer.addLayout(body, 1)

        # ── 左側:設定 ──
        left = QVBoxLayout()
        left.setSpacing(14)
        body.addLayout(left, 0)

        self.settings_area = QScrollArea()
        self.settings_area.setWidgetResizable(True)
        self.settings_area.setFixedWidth(450)
        self.settings_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        inner = QWidget()
        sv = QVBoxLayout(inner)
        sv.setContentsMargins(0, 0, 10, 0)
        sv.setSpacing(14)
        self.settings_area.setWidget(inner)
        left.addWidget(self.settings_area, 1)

        # 卡片 1:路徑
        c, l = make_card("資料來源與輸出", "可直接把資料夾拖進輸入框")
        l.addWidget(self._field_label("來源資料夾  SOURCE_FOLDER"))
        row = QHBoxLayout()
        self.src_edit = DropLineEdit("dir")
        self.src_edit.setPlaceholderText("選擇要搜尋的資料夾")
        b = QPushButton("瀏覽")
        b.clicked.connect(self._pick_source)
        row.addWidget(self.src_edit, 1)
        row.addWidget(b)
        l.addLayout(row)
        l.addWidget(self._field_label("輸出檔案  OUTPUT_FILE"))
        row = QHBoxLayout()
        self.out_edit = DropLineEdit("file")
        self.out_edit.setPlaceholderText("結果檔案路徑 (.xlsx 或 .csv)")
        b = QPushButton("另存")
        b.clicked.connect(self._pick_output)
        row.addWidget(self.out_edit, 1)
        row.addWidget(b)
        l.addLayout(row)
        sv.addWidget(c)

        # 卡片 2:關鍵字
        c, l = make_card("搜尋關鍵字", "輸入後按 Enter 新增;雙擊清單項目可移除")
        row = QHBoxLayout()
        self.kw_edit = QLineEdit()
        self.kw_edit.setPlaceholderText("輸入關鍵字…")
        self.kw_edit.returnPressed.connect(self._add_keyword)
        b = QPushButton("新增")
        b.clicked.connect(self._add_keyword)
        row.addWidget(self.kw_edit, 1)
        row.addWidget(b)
        l.addLayout(row)
        self.kw_list = QListWidget()
        self.kw_list.setFixedHeight(130)
        self.kw_list.itemDoubleClicked.connect(lambda it: self._remove_keyword(it))
        l.addWidget(self.kw_list)
        row = QHBoxLayout()
        b1 = QPushButton("移除選取")
        b1.clicked.connect(self._remove_selected)
        b2 = QPushButton("全部清除")
        b2.clicked.connect(self.kw_list.clear)
        row.addWidget(b1)
        row.addWidget(b2)
        row.addStretch(1)
        l.addLayout(row)
        sv.addWidget(c)

        # 卡片 3:比對規則
        c, l = make_card("比對規則")
        self.sw_case = Switch(False)
        l.addWidget(SettingRow("區分大小寫", "開啟後 Chrome 與 chrome 視為不同", self.sw_case))
        l.addWidget(self._divider())
        self.sw_whole = Switch(False)
        l.addWidget(SettingRow("完整比對", "關閉時,只要內容「包含」關鍵字就算符合", self.sw_whole))
        self.cb_whole_mode = QComboBox()
        self.cb_whole_mode.addItem("整個儲存格完全相同", "cell")
        self.cb_whole_mode.addItem("完整單字 (前後不可連著文字)", "word")
        self.cb_whole_mode.setToolTip(
            "整個儲存格完全相同:儲存格內容必須恰好等於關鍵字\n"
            "完整單字:關鍵字前後不能緊接英數字或中文字")
        l.addWidget(SettingRow("完整比對方式", "只在「完整比對」開啟時有效", self.cb_whole_mode))
        self.sw_whole.toggled.connect(self.cb_whole_mode.setEnabled)
        self.cb_whole_mode.setEnabled(False)
        self.cb_whole_mode.setMinimumWidth(190)
        l.addWidget(self._divider())
        self.cb_logic = QComboBox()
        self.cb_logic.addItem("符合任一關鍵字", False)
        self.cb_logic.addItem("全部關鍵字都要出現", True)
        l.addWidget(SettingRow("多個關鍵字", "決定一列資料要符合幾個關鍵字", self.cb_logic))
        sv.addWidget(c)

        # 卡片 4:檔案範圍
        c, l = make_card("檔案範圍")
        self.sw_recursive = Switch(True)
        l.addWidget(SettingRow("包含子資料夾", "遞迴搜尋所有下層資料夾", self.sw_recursive))
        l.addWidget(self._divider())
        self.sw_hidden = Switch(True)
        l.addWidget(SettingRow("略過隱藏項目", "忽略以 . 開頭的資料夾與檔案", self.sw_hidden))
        l.addWidget(self._divider())
        l.addWidget(self._field_label("要搜尋的檔案類型"))
        row = QHBoxLayout()
        self.ext_chips = {}
        for ext in ALL_EXTS:
            ch = QPushButton(ext)
            ch.setObjectName("Chip")
            ch.setCheckable(True)
            ch.setChecked(True)
            ch.setCursor(Qt.CursorShape.PointingHandCursor)
            self.ext_chips[ext] = ch
            row.addWidget(ch)
        row.addStretch(1)
        l.addLayout(row)
        sv.addWidget(c)

        # 卡片 5:完成後
        c, l = make_card("完成後")
        self.sw_open = Switch(False)
        l.addWidget(SettingRow("自動開啟結果檔案", "搜尋完成後直接用預設程式開啟", self.sw_open))
        sv.addWidget(c)
        sv.addStretch(1)

        # 底部按鈕
        foot = QHBoxLayout()
        foot.setSpacing(10)
        self.start_btn = QPushButton("開始搜尋")
        self.start_btn.setObjectName("Primary")
        self.start_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.start_btn.clicked.connect(self.start)
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.setObjectName("Danger")
        self.cancel_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self.cancel)
        foot.addWidget(self.start_btn, 1)
        foot.addWidget(self.cancel_btn)
        left.addLayout(foot)

        # ── 右側:進度與結果 ──
        right = QVBoxLayout()
        right.setSpacing(14)
        body.addLayout(right, 1)

        stats = QHBoxLayout()
        stats.setSpacing(12)
        self.st_scanned = StatCard("已掃描檔案", "0")
        self.st_files = StatCard("命中檔案", "0")
        self.st_rows = StatCard("符合筆數", "0")
        self.st_time = StatCard("耗時", "—")
        for s in (self.st_scanned, self.st_files, self.st_rows, self.st_time):
            stats.addWidget(s, 1)
        right.addLayout(stats)

        pc, pl = make_card("進度")
        self.status_lbl = QLabel("準備就緒,設定好左側選項後按「開始搜尋」")
        self.status_lbl.setObjectName("Path")
        self.status_lbl.setWordWrap(True)
        self.bar = QProgressBar()
        self.bar.setRange(0, 1)
        self.bar.setValue(0)
        pl.addWidget(self.bar)
        pl.addWidget(self.status_lbl)
        right.addWidget(pc)

        rc, rl = make_card("結果")
        self.tabs = QTabWidget()
        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.table = QTableWidget()
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.tabs.addTab(self.log_view, "執行紀錄")
        self.tabs.addTab(self.table, "結果預覽")
        rl.addWidget(self.tabs, 1)
        row = QHBoxLayout()
        self.preview_note = QLabel("")
        self.preview_note.setObjectName("Muted")
        self.open_file_btn = QPushButton("開啟結果檔案")
        self.open_dir_btn = QPushButton("開啟所在資料夾")
        self.open_file_btn.setEnabled(False)
        self.open_dir_btn.setEnabled(False)
        self.open_file_btn.clicked.connect(self._open_result)
        self.open_dir_btn.clicked.connect(self._open_result_dir)
        row.addWidget(self.preview_note, 1)
        row.addWidget(self.open_file_btn)
        row.addWidget(self.open_dir_btn)
        rl.addLayout(row)
        right.addWidget(rc, 1)

    @staticmethod
    def _field_label(text: str) -> QLabel:
        lb = QLabel(text)
        lb.setObjectName("Muted")
        return lb

    @staticmethod
    def _divider() -> QFrame:
        d = QFrame()
        d.setObjectName("Divider")
        return d

    # ----- 主題 -----
    def apply_theme(self, name: str):
        self.theme_name = name
        _CURRENT["t"] = THEMES[name]
        self.app.setStyleSheet(QSS % THEMES[name])
        self.theme_btn.setText("☀  淺色" if name == "dark" else "🌙  深色")
        for sw in self.findChildren(Switch):
            sw.update()

    def toggle_theme(self):
        self.apply_theme("light" if self.theme_name == "dark" else "dark")

    # ----- 路徑 / 關鍵字 -----
    def _pick_source(self):
        d = QFileDialog.getExistingDirectory(self, "選擇來源資料夾", self.src_edit.text() or str(Path.home()))
        if d:
            self.src_edit.setText(os.path.normpath(d))

    def _pick_output(self):
        p, _ = QFileDialog.getSaveFileName(
            self, "選擇輸出檔案", self.out_edit.text() or DEFAULT_OUTPUT,
            "Excel 活頁簿 (*.xlsx);;CSV 檔案 (*.csv)")
        if p:
            self.out_edit.setText(os.path.normpath(p))

    def _keywords(self) -> list:
        return [self.kw_list.item(i).text() for i in range(self.kw_list.count())]

    def _add_keyword(self):
        text = self.kw_edit.text().strip()
        if not text:
            return
        if text.casefold() in (k.casefold() for k in self._keywords()):
            self.kw_edit.clear()
            return
        self.kw_list.addItem(text)
        self.kw_edit.clear()

    def _remove_keyword(self, item):
        self.kw_list.takeItem(self.kw_list.row(item))

    def _remove_selected(self):
        for it in self.kw_list.selectedItems():
            self.kw_list.takeItem(self.kw_list.row(it))

    # ----- 設定存取 -----
    def _load_settings(self):
        s = self.settings
        self.src_edit.setText(s.value("source", DEFAULT_SOURCE))
        self.out_edit.setText(s.value("output", DEFAULT_OUTPUT))
        try:
            kws = json.loads(s.value("keywords", ""))
        except (TypeError, ValueError):
            kws = DEFAULT_KEYWORDS
        self.kw_list.addItems(kws)
        self.sw_case.setChecked(s.value("case", False, type=bool))
        self.sw_whole.setChecked(s.value("whole", False, type=bool))
        self.cb_whole_mode.setCurrentIndex(max(0, self.cb_whole_mode.findData(s.value("whole_mode", "cell"))))
        self.cb_logic.setCurrentIndex(1 if s.value("match_all", False, type=bool) else 0)
        self.sw_recursive.setChecked(s.value("recursive", True, type=bool))
        self.sw_hidden.setChecked(s.value("skip_hidden", True, type=bool))
        self.sw_open.setChecked(s.value("auto_open", False, type=bool))
        saved = s.value("exts", None)
        if isinstance(saved, str):
            saved = [saved]
        if saved:
            for ext, ch in self.ext_chips.items():
                ch.setChecked(ext in saved)
        theme = s.value("theme", None)
        if theme not in THEMES:
            theme = "dark"
            try:
                if self.app.styleHints().colorScheme() == Qt.ColorScheme.Light:
                    theme = "light"
            except AttributeError:
                pass
        self.theme_name = theme

    def _save_settings(self):
        s = self.settings
        s.setValue("source", self.src_edit.text())
        s.setValue("output", self.out_edit.text())
        s.setValue("keywords", json.dumps(self._keywords(), ensure_ascii=False))
        s.setValue("case", self.sw_case.isChecked())
        s.setValue("whole", self.sw_whole.isChecked())
        s.setValue("whole_mode", self.cb_whole_mode.currentData())
        s.setValue("match_all", self.cb_logic.currentData())
        s.setValue("recursive", self.sw_recursive.isChecked())
        s.setValue("skip_hidden", self.sw_hidden.isChecked())
        s.setValue("auto_open", self.sw_open.isChecked())
        s.setValue("exts", [e for e, c in self.ext_chips.items() if c.isChecked()])
        s.setValue("theme", self.theme_name)

    # ----- 執行 -----
    def _config(self) -> SearchConfig:
        return SearchConfig(
            source=self.src_edit.text().strip(),
            output=self.out_edit.text().strip(),
            keywords=self._keywords(),
            case_sensitive=self.sw_case.isChecked(),
            whole=self.sw_whole.isChecked(),
            whole_mode=self.cb_whole_mode.currentData(),
            match_all=bool(self.cb_logic.currentData()),
            recursive=self.sw_recursive.isChecked(),
            exts=tuple(e for e, c in self.ext_chips.items() if c.isChecked()),
            skip_hidden=self.sw_hidden.isChecked(),
        )

    def _validate(self, cfg: SearchConfig) -> str | None:
        if not cfg.source or not os.path.isdir(cfg.source):
            return "來源資料夾不存在,請重新選擇。"
        if not cfg.output or not cfg.output.lower().endswith((".xlsx", ".csv")):
            return "輸出檔案請使用 .xlsx 或 .csv 副檔名。"
        if not cfg.keywords:
            return "請至少新增一個搜尋關鍵字。"
        if not cfg.exts:
            return "請至少選擇一種檔案類型。"
        return None

    def start(self):
        cfg = self._config()
        err = self._validate(cfg)
        if err:
            QMessageBox.warning(self, APP_NAME, err)
            return
        self._save_settings()
        self.result_path = None
        self.log_view.clear()
        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.preview_note.setText("")
        self.open_file_btn.setEnabled(False)
        self.open_dir_btn.setEnabled(False)
        for s in (self.st_scanned, self.st_files, self.st_rows):
            s.set("0")
        self.st_time.set("0.0 秒")
        self.bar.setRange(0, 0)  # 忙碌動畫,等取得總數
        self.tabs.setCurrentIndex(0)
        self._set_running(True)

        self.worker = SearchWorker(cfg)
        self.worker.progress.connect(self._on_progress)
        self.worker.stats.connect(self._on_stats)
        self.worker.log.connect(self._on_log)
        self.worker.done.connect(self._on_done)
        self.worker.failed.connect(self._on_failed)
        self.t0 = time.time()
        self.timer.start()
        self.worker.start()

    def cancel(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.cancel_btn.setEnabled(False)
            self.status_lbl.setText("正在取消…")

    def _set_running(self, running: bool):
        self.settings_area.setEnabled(not running)
        self.start_btn.setEnabled(not running)
        self.cancel_btn.setEnabled(running)
        self.start_btn.setText("搜尋中…" if running else "開始搜尋")

    def _tick(self):
        self.st_time.set(fmt_seconds(time.time() - self.t0))

    # ----- 工作執行緒回呼 -----
    def _on_progress(self, cur: int, total: int, path: str):
        if total > 0:
            self.bar.setRange(0, total)
            self.bar.setValue(cur)
        if path:
            short = path if len(path) <= 80 else "…" + path[-78:]
            self.status_lbl.setText(f"正在處理 ({cur + 1}/{total})  {short}")

    def _on_stats(self, scanned: int, files: int, rows: int):
        self.st_scanned.set(f"{scanned:,}")
        self.st_files.set(f"{files:,}")
        self.st_rows.set(f"{rows:,}")

    def _on_log(self, level: str, msg: str):
        t = current_theme()
        color = {"ok": t["ok"], "err": t["err"], "warn": t["warn"]}.get(level, t["text"])
        stamp = time.strftime("%H:%M:%S")
        self.log_view.append(
            f'<span style="color:{t["muted"]}">{stamp}</span>&nbsp;&nbsp;'
            f'<span style="color:{color}">{html.escape(msg)}</span>')

    def _on_done(self, r: dict):
        self.timer.stop()
        self._set_running(False)
        self.st_time.set(fmt_seconds(r["elapsed"]))
        self._on_stats(r["scanned"], r["matched_files"], r["rows"])
        if self.bar.maximum() == 0:
            self.bar.setRange(0, 1)
        if r["cancelled"]:
            self.status_lbl.setText("已取消")
            return
        if r["output"]:
            self.result_path = r["output"]
            self.status_lbl.setText(f"完成!結果已寫入 {r['output']}")
            self.open_file_btn.setEnabled(True)
            self.open_dir_btn.setEnabled(True)
            self._fill_preview(r["preview"], r["rows"])
            self.tabs.setCurrentIndex(1)
            if self.sw_open.isChecked():
                self._open_result()
        else:
            self.status_lbl.setText("完成,但沒有找到符合的資料")

    def _on_failed(self, msg: str):
        self.timer.stop()
        self._set_running(False)
        self.bar.setRange(0, 1)
        self.status_lbl.setText("執行失敗")
        self._on_log("err", msg.replace("\n", " "))
        QMessageBox.critical(self, APP_NAME, msg)

    def _fill_preview(self, df: pd.DataFrame, total_rows: int):
        df = df.fillna("").astype(str)
        self.table.setColumnCount(df.shape[1])
        self.table.setRowCount(df.shape[0])
        self.table.setHorizontalHeaderLabels([str(c) for c in df.columns])
        for r, row in enumerate(df.itertuples(index=False)):
            for c, v in enumerate(row):
                self.table.setItem(r, c, QTableWidgetItem(v))
        self.table.resizeColumnsToContents()
        for c in range(self.table.columnCount()):
            if self.table.columnWidth(c) > 260:
                self.table.setColumnWidth(c, 260)
        shown = df.shape[0]
        self.preview_note.setText(
            f"預覽前 {shown} 筆 (共 {total_rows:,} 筆)" if total_rows > shown
            else f"共 {total_rows:,} 筆")

    def _open_result(self):
        if self.result_path and os.path.exists(self.result_path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.result_path))

    def _open_result_dir(self):
        if self.result_path:
            QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(os.path.abspath(self.result_path))))

    def closeEvent(self, e):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.worker.wait(3000)
        self._save_settings()
        super().closeEvent(e)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    font = QFont()
    font.setFamilies(["Segoe UI", "Microsoft JhengHei UI", "PingFang TC", "Noto Sans TC", "sans-serif"])
    font.setPointSize(10)
    app.setFont(font)
    win = MainWindow(app)
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
