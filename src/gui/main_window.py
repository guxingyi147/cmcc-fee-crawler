# -*- coding: utf-8 -*-
"""
主窗口
"""

import threading
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QTableWidget, QTableWidgetItem,
    QTextEdit, QFileDialog, QSizePolicy, QMessageBox, QAbstractItemView,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QDoubleValidator

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

from ..constants import (
    TARGET_URL, COLUMNS, COL_COUNT, COL_PLAN_ID, COL_NAME, COL_PRICE,
)
from ..crawler.crawler_thread import CrawlerThread, WorkerSignals
from ..driver import create_driver
from ..export.excel_writer import write_excel, parse_price

# 样式表常量
GLOBAL_STYLESHEET = """
    QMainWindow{background:#f5f5f5;}
    QPushButton{background:#fff; border:1px solid #d9d9d9;
        border-radius:4px; padding:6px 16px; font-size:13px; color:#333;}
    QPushButton:hover{border-color:#40a9ff; color:#40a9ff;}
    QPushButton:pressed{background:#e6f7ff; border-color:#1890ff;}
    QPushButton:disabled{color:#bfbfbf; border-color:#d9d9d9;
        background:#f5f5f5;}
    QLineEdit{border:1px solid #d9d9d9; border-radius:4px;
        padding:5px 10px; font-size:13px; background:#fff;}
    QLineEdit:focus{border-color:#40a9ff;}
    QTableWidget{background:#fff; border:1px solid #e8e8e8;
        border-radius:4px; gridline-color:#e8e8e8; font-size:12px;}
    QTableWidget::item{padding:4px 6px;}
    QTableWidget::item:alternate{background:#fafafa;}
    QHeaderView::section{background:#fafafa; border:1px solid #e8e8e8;
        padding:5px 8px; font-weight:bold; font-size:12px; color:#333;}
    QLabel{font-size:13px; color:#333;}
"""

LOG_AREA_STYLE = (
    "QTextEdit{background:#1e1e1e; color:#d4d4d4;"
    "font-family:'Consolas','Microsoft YaHei',monospace;"
    "font-size:12px; border:1px solid #444; border-radius:4px;}"
)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("中国移动资费公示爬取工具")
        self.setMinimumSize(1200, 750)
        self.resize(1400, 850)

        self.driver = None
        self.crawler_thread = None
        self.stop_event = threading.Event()
        self.all_data = []
        self.seen_ids = set()
        self.is_crawling = False
        self.signals = WorkerSignals()
        
        # 日志视图状态
        self.log_expanded = False
        
        # 抓取计时（仅用于结束时计算耗时）
        self.crawl_start_time = None

        self._connect_signals()
        self._build_ui()
        self.append_log("[信息] 应用已启动，请点击「启动浏览器」开始")

    # ---------- 信号连接 ----------

    def _connect_signals(self):
        self.signals.log.connect(self.append_log)
        self.signals.status.connect(self.update_status)
        self.signals.new_data.connect(self._on_new_data)
        self.signals.type_switched.connect(self._on_type_switched)
        self.signals.error.connect(self._on_error)
        self.signals.finished.connect(self._on_crawl_finished)

    # ---------- UI 构建 ----------

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        # ---- 1. 顶部按钮 ----
        self.btn_row_widget = QWidget()
        btn_row = QHBoxLayout(self.btn_row_widget)
        btn_row.setContentsMargins(0, 0, 0, 0)
        btn_row.setSpacing(10)

        self.btn_launch = QPushButton("启动浏览器")
        self.btn_start = QPushButton("开始抓取")
        self.btn_stop = QPushButton("停止抓取")
        self.btn_clear = QPushButton("清空数据")
        self.btn_export = QPushButton("导出Excel")

        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(False)

        for b in (self.btn_launch, self.btn_start, self.btn_stop,
                  self.btn_clear, self.btn_export):
            b.setFixedHeight(36)
            b.setMinimumWidth(110)
            btn_row.addWidget(b)
        btn_row.addStretch()

        self.btn_launch.clicked.connect(self.launch_browser)
        self.btn_start.clicked.connect(self.start_crawl)
        self.btn_stop.clicked.connect(self.stop_crawl)
        self.btn_clear.clicked.connect(self.clear_data)
        self.btn_export.clicked.connect(self.export_excel)
        root.addWidget(self.btn_row_widget)

        # ---- 2. 搜索筛选 ----
        self.filter_row_widget = QWidget()
        filter_row = QHBoxLayout(self.filter_row_widget)
        filter_row.setContentsMargins(0, 0, 0, 0)
        filter_row.setSpacing(12)

        filter_row.addWidget(QLabel("关键字搜索:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(
            "按套餐名称 / 方案编号模糊搜索")
        self.search_input.setFixedWidth(280)
        self.search_input.textChanged.connect(self.apply_filter)
        filter_row.addWidget(self.search_input)

        filter_row.addWidget(QLabel("价格区间:"))
        self.price_min = QLineEdit()
        self.price_min.setPlaceholderText("最低")
        self.price_min.setFixedWidth(80)
        self.price_min.setValidator(QDoubleValidator(0, 99999, 2))
        filter_row.addWidget(self.price_min)

        filter_row.addWidget(QLabel("—"))

        self.price_max = QLineEdit()
        self.price_max.setPlaceholderText("最高")
        self.price_max.setFixedWidth(80)
        self.price_max.setValidator(QDoubleValidator(0, 99999, 2))
        filter_row.addWidget(self.price_max)

        btn_f = QPushButton("筛选")
        btn_f.setFixedWidth(60)
        btn_f.clicked.connect(self.apply_filter)
        filter_row.addWidget(btn_f)

        btn_r = QPushButton("重置")
        btn_r.setFixedWidth(60)
        btn_r.clicked.connect(self.reset_filter)
        filter_row.addWidget(btn_r)

        filter_row.addStretch()

        self.label_count = QLabel("共 0 条数据")
        self.label_count.setStyleSheet(
            "font-weight:bold; color:#1890ff; font-size:13px;")
        filter_row.addWidget(self.label_count)
        root.addWidget(self.filter_row_widget)

        # ---- 3. 数据表格 ----
        self.table = QTableWidget()
        self.table.setColumnCount(COL_COUNT)
        self.table.setHorizontalHeaderLabels(COLUMNS)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(
            QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setDefaultSectionSize(28)
        self.table.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        hdr = self.table.horizontalHeader()
        widths = {
            0: 110, 1: 220, 2: 100, 3: 80, 4: 80,
            5: 100, 6: 120, 7: 100, 8: 100, 9: 140,
            10: 80, 11: 100, 12: 110, 13: 100, 14: 80,
            15: 160, 16: 160, 17: 200,
        }
        for c, w in widths.items():
            hdr.resizeSection(c, w)
        root.addWidget(self.table, stretch=1)

        # ---- 4. 底部状态栏 ----
        bot = QVBoxLayout()
        bot.setSpacing(4)

        status_row = QHBoxLayout()
        self.label_status = QLabel("状态: 就绪")
        self.label_status.setStyleSheet(
            "color:#52c41a; font-weight:bold;")
        status_row.addWidget(self.label_status)

        self.label_current_type = QLabel("")
        self.label_current_type.setStyleSheet(
            "color:#722ed1; font-weight:bold; font-size:12px;")
        status_row.addWidget(self.label_current_type)

        status_row.addStretch()
        
        # 日志放大/缩小按钮
        self.btn_expand_log = QPushButton("放大日志")
        self.btn_expand_log.setFixedSize(80, 28)
        self.btn_expand_log.setStyleSheet(
            "QPushButton{font-size:11px; padding:2px 8px;}"
            "QPushButton:hover{background:#e6f7ff;}")
        self.btn_expand_log.clicked.connect(self.toggle_log_size)
        status_row.addWidget(self.btn_expand_log)
        
        bot.addLayout(status_row)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(130)
        self.log_text.setStyleSheet(LOG_AREA_STYLE)
        bot.addWidget(self.log_text)
        root.addLayout(bot)

        self.setStyleSheet(GLOBAL_STYLESHEET)

    # ========== 浏览器控制 ==========

    def launch_browser(self):
        if self.driver:
            self.append_log("[警告] 浏览器已在运行中")
            return
        try:
            self.append_log("[信息] 正在启动浏览器...")
            self.update_status("启动浏览器中...")

            self.driver = create_driver()

            # 反检测
            self.driver.execute_cdp_cmd(
                "Page.addScriptToEvaluateOnNewDocument", {"source": """
                    Object.defineProperty(navigator,'webdriver',{
                        get:()=>undefined});
                """})

            self.driver.get(TARGET_URL)
            self.append_log("[信息] 浏览器已启动，已打开资费公示页面")
            self.append_log(
                "[提示] 请在浏览器中手动选择省份（如江西省）")
            self.append_log(
                "[提示] 选择完毕后，点击「开始抓取」按钮")
            self.append_log(
                "[提示] 程序将自动遍历所有资费类型并抓取数据")

            self.btn_launch.setEnabled(False)
            self.btn_start.setEnabled(True)
            self.update_status("等待用户操作浏览器...")

        except Exception as e:
            self.append_log(f"[错误] 启动浏览器失败: {str(e)}")
            self.update_status("浏览器启动失败")
            self.driver = None
            self.btn_launch.setEnabled(True)

    def start_crawl(self):
        if not self.driver:
            self.append_log("[错误] 请先启动浏览器")
            return
        try:
            _ = self.driver.current_url
        except Exception:
            self.append_log("[错误] 浏览器已关闭，请重新启动")
            self._reset_browser_state()
            return

        self.is_crawling = True
        self.stop_event.clear()
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.btn_launch.setEnabled(False)
        
        # 记录开始时间（用于结束时计算耗时）
        self.crawl_start_time = datetime.now()

        self.crawler_thread = CrawlerThread(
            self.driver, self.signals, self.stop_event,
            traverse_types=True)
        self.crawler_thread.start()

    def stop_crawl(self):
        self.append_log("[信息] 正在停止抓取...")
        self.stop_event.set()
        self.btn_stop.setEnabled(False)

    # ========== 数据接收 ==========

    def _on_new_data(self, batch):
        if not batch:
            return
        self.all_data.extend(batch)
        self.append_log(f"[信息] 新增 {len(batch)} 条，"
                        f"累计 {len(self.all_data)} 条")
        self._append_rows(batch)

    def _append_rows(self, rows):
        if (self.search_input.text().strip()
                or self.price_min.text().strip()
                or self.price_max.text().strip()):
            self.apply_filter()
            return

        row_count = self.table.rowCount()
        self.table.setRowCount(row_count + len(rows))
        for i, rd in enumerate(rows):
            for j, v in enumerate(rd):
                item = QTableWidgetItem(str(v) if v else "")
                item.setTextAlignment(
                    Qt.AlignmentFlag.AlignLeft
                    | Qt.AlignmentFlag.AlignVCenter)
                self.table.setItem(row_count + i, j, item)

        self.label_count.setText(f"共 {len(self.all_data)} 条数据")

    # ========== 筛选 ==========

    def apply_filter(self):
        keyword = self.search_input.text().strip().lower()
        pmin = self.price_min.text().strip()
        pmax = self.price_max.text().strip()
        price_lo = float(pmin) if pmin else 0
        price_hi = float(pmax) if pmax else float('inf')

        filtered = []
        for row in self.all_data:
            if keyword:
                if not (keyword in row[COL_PLAN_ID].lower()
                        or keyword in row[COL_NAME].lower()):
                    continue
            price = parse_price(row[COL_PRICE])
            if price is not None and not (price_lo <= price <= price_hi):
                continue
            filtered.append(row)

        self.table.setRowCount(len(filtered))
        for i, rd in enumerate(filtered):
            for j, v in enumerate(rd):
                item = QTableWidgetItem(str(v) if v else "")
                item.setTextAlignment(
                    Qt.AlignmentFlag.AlignLeft
                    | Qt.AlignmentFlag.AlignVCenter)
                self.table.setItem(i, j, item)

        total = len(self.all_data)
        shown = len(filtered)
        self.label_count.setText(
            f"共 {total} 条数据"
            + (f"（筛选显示 {shown} 条）" if total != shown else ""))

    def reset_filter(self):
        self.search_input.clear()
        self.price_min.clear()
        self.price_max.clear()
        self.apply_filter()

    # ========== 数据操作 ==========

    def clear_data(self):
        if self.is_crawling:
            reply = QMessageBox.question(
                self, "确认", "正在抓取数据，确定要清空吗？",
                QMessageBox.StandardButton.Yes
                | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No)
        else:
            reply = QMessageBox.StandardButton.Yes

        if reply == QMessageBox.StandardButton.Yes:
            self.all_data.clear()
            self.seen_ids.clear()
            self.table.setRowCount(0)
            self.label_count.setText("共 0 条数据")
            self.append_log("[信息] 数据已清空")

    def export_excel(self):
        if not self.all_data:
            QMessageBox.warning(self, "提示", "没有数据可导出")
            return
        if not HAS_OPENPYXL:
            QMessageBox.critical(
                self, "错误",
                "缺少 openpyxl，请执行: pip install openpyxl")
            return

        keyword = self.search_input.text().strip().lower()
        pmin = self.price_min.text().strip()
        pmax = self.price_max.text().strip()
        price_lo = float(pmin) if pmin else 0
        price_hi = float(pmax) if pmax else float('inf')

        export = []
        for row in self.all_data:
            if keyword:
                if not (keyword in row[COL_PLAN_ID].lower()
                        or keyword in row[COL_NAME].lower()):
                    continue
            price = parse_price(row[COL_PRICE])
            if price is not None and not (
                    price_lo <= price <= price_hi):
                continue
            export.append(row)

        if not export:
            QMessageBox.warning(self, "提示", "筛选后没有数据可导出")
            return

        default_name = (
            f"中国移动资费数据_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        path, _ = QFileDialog.getSaveFileName(
            self, "导出Excel", default_name, "Excel文件 (*.xlsx)")
        if not path:
            return

        try:
            self.append_log(f"[信息] 正在导出 {len(export)} 条数据...")
            write_excel(path, export)
            self.append_log(f"[信息] 导出成功: {path}")
            QMessageBox.information(
                self, "导出成功",
                f"已导出 {len(export)} 条数据到:\n{path}")
        except Exception as e:
            self.append_log(f"[错误] 导出失败: {str(e)}")
            QMessageBox.critical(self, "导出失败", str(e))

    # ========== UI 辅助 ==========

    def append_log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{ts}] {msg}")
        sb = self.log_text.verticalScrollBar()
        sb.setValue(sb.maximum())

    def update_status(self, text):
        self.label_status.setText(f"状态: {text}")
        if "抓取中" in text:
            color = "#1890ff"
        elif "完成" in text or "成功" in text:
            color = "#52c41a"
        elif "错误" in text or "失败" in text:
            color = "#ff4d4f"
        else:
            color = "#faad14"
        self.label_status.setStyleSheet(
            f"color:{color}; font-weight:bold;")

    def _on_type_switched(self, info):
        self.label_current_type.setText(f"当前: {info}")

    def _on_error(self, msg):
        self.append_log(msg)
        self.update_status("出错")

    def _on_crawl_finished(self):
        self.is_crawling = False
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        
        # 显示最终耗时
        if self.crawl_start_time:
            elapsed = datetime.now() - self.crawl_start_time
            total_seconds = int(elapsed.total_seconds())
            hours, remainder = divmod(total_seconds, 3600)
            minutes, seconds = divmod(remainder, 60)
            if hours > 0:
                time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
            else:
                time_str = f"{minutes:02d}:{seconds:02d}"
            self.append_log(
                f"[信息] 抓取结束，共获取 {len(self.all_data)} 条数据，耗时 {time_str}")
        else:
            self.append_log(
                f"[信息] 抓取结束，共获取 {len(self.all_data)} 条数据")

    def _reset_browser_state(self):
        self.driver = None
        self.btn_launch.setEnabled(True)
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(False)
        self.is_crawling = False
        self.update_status("浏览器未启动")
        
    def toggle_log_size(self):
        """切换日志视图大小 - 全屏覆盖模式"""
        if self.log_expanded:
            # 缩小：显示被隐藏的控件
            self.btn_row_widget.setVisible(True)
            self.filter_row_widget.setVisible(True)
            self.table.setVisible(True)
            self.log_text.setMaximumHeight(130)
            self.btn_expand_log.setText("放大日志")
            self.log_expanded = False
        else:
            # 放大：隐藏表格和筛选区，日志占满窗口
            self.btn_row_widget.setVisible(False)
            self.filter_row_widget.setVisible(False)
            self.table.setVisible(False)
            self.log_text.setMaximumHeight(16777215)  # QWIDGETSIZE_MAX
            self.btn_expand_log.setText("缩小日志")
            self.log_expanded = True

    # ========== 关闭 ==========

    def closeEvent(self, event):
        if self.is_crawling:
            if QMessageBox.question(
                self, "确认退出", "正在抓取数据，确定要退出吗？",
                QMessageBox.StandardButton.Yes
                | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            ) == QMessageBox.StandardButton.No:
                event.ignore()
                return
            self.stop_event.set()

        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

        event.accept()