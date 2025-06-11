import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QFileDialog, QMessageBox, QProgressBar, QComboBox, QTableWidget, QTableWidgetItem, QHeaderView, QTextEdit, QTabWidget, QInputDialog, QSizePolicy, QMenu
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from ui_signals import continue_signal
from ui_worker import BatchSearchWorker

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("企业备案批量工具 by心海")
        self.setWindowIcon(QIcon("icpsearch.png"))
        self.resize(1200, 800)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.init_ui(central_widget)
        self.results = []
        continue_signal.ask_continue.connect(self.on_ask_continue)

    def init_ui(self, central_widget):
        layout = QVBoxLayout(central_widget)
        import_layout = QHBoxLayout()
        import_layout.setSpacing(8)
        import_layout.addWidget(QLabel("企业名称:"))
        self.manual_input = QTextEdit()
        self.manual_input.setFixedHeight(30)
        self.manual_input.setAlignment(Qt.AlignVCenter)
        self.manual_input.setStyleSheet("QTextEdit { text-align: center; }")
        import_layout.addWidget(self.manual_input, stretch=1)
        self.import_btn = QPushButton("导入")
        self.import_btn.setFixedWidth(80)
        self.import_btn.clicked.connect(self.import_file)
        import_layout.addWidget(self.import_btn)
        layout.addLayout(import_layout)
        proxy_layout = QHBoxLayout()
        proxy_layout.setSpacing(8)
        proxy_layout.addWidget(QLabel("代理模式:"))
        self.proxy_mode_combo = QComboBox()
        self.proxy_mode_combo.addItems(["无代理", "代理"])
        self.proxy_mode_combo.setFixedWidth(90)
        self.proxy_mode_combo.currentIndexChanged.connect(self.proxy_mode_changed)
        proxy_layout.addWidget(self.proxy_mode_combo)
        self.proxy_input = QLineEdit()
        self.proxy_input.setPlaceholderText("代理地址")
        self.proxy_input.setAlignment(Qt.AlignCenter)
        proxy_layout.addWidget(self.proxy_input, stretch=1)
        proxy_layout.addWidget(QLabel("查询间隔(秒):"))
        self.delay_input = QLineEdit("0")
        self.delay_input.setFixedWidth(60)
        self.delay_input.setAlignment(Qt.AlignCenter)
        proxy_layout.addWidget(self.delay_input)
        layout.addLayout(proxy_layout)
        search_layout = QHBoxLayout()
        search_layout.setSpacing(8)
        search_layout.addWidget(QLabel("查找企业名:"))
        self.search_input = QLineEdit()
        self.search_input.setAlignment(Qt.AlignCenter)
        search_layout.addWidget(self.search_input, stretch=1)
        self.search_btn = QPushButton("查找")
        self.search_btn.setFixedWidth(60)
        self.search_btn.clicked.connect(self.find_company)
        search_layout.addWidget(self.search_btn)
        self.down_btn = QPushButton("下一个")
        self.down_btn.setFixedWidth(110)
        self.down_btn.clicked.connect(self.find_company)
        search_layout.addWidget(self.down_btn)
        self.adjust_btn = QPushButton("自动调整行列")
        self.adjust_btn.setFixedWidth(120)
        self.adjust_btn.clicked.connect(self.auto_adjust)
        search_layout.addWidget(self.adjust_btn)
        layout.addLayout(search_layout)
        self.search_btn_main = QPushButton("开始查询")
        self.search_btn_main.setStyleSheet("QPushButton { background-color: #0078d7; color: white; font-weight: bold; }QPushButton:hover { background-color: #005fa1; }")
        self.search_btn_main.setFixedHeight(40)
        layout.addWidget(self.search_btn_main)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(True)
        layout.addWidget(self.progress_bar)
        self.tabs = QTabWidget()
        self.tabs.setTabsClosable(True)
        self.tabs.tabBarDoubleClicked.connect(self.rename_tab)
        self.tabs.tabCloseRequested.connect(self.close_tab)
        self.tabs.currentChanged.connect(self.on_tab_changed)
        layout.addWidget(self.tabs)
        btn_widget = QWidget()
        btn_layout = QHBoxLayout(btn_widget)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(0)
        self.add_tab_btn = QPushButton("+")
        self.add_tab_btn.setFixedWidth(30)
        self.add_tab_btn.clicked.connect(self.add_tab)
        btn_layout.addWidget(self.add_tab_btn)
        self.del_tab_btn = QPushButton("-")
        self.del_tab_btn.setFixedWidth(30)
        self.del_tab_btn.clicked.connect(self.del_tab)
        btn_layout.addWidget(self.del_tab_btn)
        btn_layout.addStretch()
        self.tabs.setCornerWidget(btn_widget, Qt.TopRightCorner)
        self.export_btn = QPushButton("将此标签页导出-result.xlsx")
        self.export_btn.clicked.connect(self.export_results)
        self.export_btn.setEnabled(False)
        layout.addWidget(self.export_btn)
        self.add_tab()
        self.search_btn_main.clicked.connect(self.start_search)
        self.search_btn.clicked.connect(self.find_company)
        self.adjust_btn.clicked.connect(self.auto_adjust)

    def proxy_mode_changed(self):
        mode = self.proxy_mode_combo.currentText()
        if mode == "代理":
            self.proxy_input.setPlaceholderText("代理地址，如http://127.0.0.1:8080")
        else:
            self.proxy_input.setPlaceholderText("")

    def batch_icp(self):
        pass

    def start_search(self):
        company_list = self.get_company_list()
        if not company_list:
            QMessageBox.warning(self, "警告", "请导入企业名单或手动输入企业名")
            return
        proxy_mode = self.proxy_mode_combo.currentText()
        proxy_value = self.proxy_input.text().strip()
        if proxy_mode == "代理" and not proxy_value:
            QMessageBox.warning(self, "警告", "请输入代理地址")
            return
        if proxy_mode == "代理池" and not proxy_value:
            QMessageBox.warning(self, "警告", "请选择代理池文件")
            return
        try:
            delay = float(self.delay_input.text())
        except Exception:
            delay = 0
        self.current_results = []
        self.worker = BatchSearchWorker(
            company_list, proxy_mode, proxy_value, delay
        )
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.row_result.connect(self.append_row)
        self.worker.finished.connect(self.search_finished)
        self.worker.error.connect(self.show_error)
        self.worker.start()

    def append_row(self, op_tuple):
        op, row_idx, row = op_tuple
        idx = self.tabs.currentIndex()
        tab_widget = self.tabs.widget(idx)
        if tab_widget is None:
            return
        table = tab_widget.findChild(QTableWidget, "result_table")
        if table is None:
            return
        found_row = None
        for i in range(table.rowCount()):
            cell = table.item(i, 0)
            if cell and cell.text() == row.get("企业名称", ""):
                found_row = i
                break
        if found_row is not None:
            for col, key in enumerate(["企业名称", "备案域名", "备案IP", "微信小程序", "备案APP"]):
                table.setItem(found_row, col, QTableWidgetItem(row.get(key, "")))
        else:
            row_idx = table.rowCount()
            table.insertRow(row_idx)
            table.setItem(row_idx, 0, QTableWidgetItem(row.get("企业名称", "")))
            table.setItem(row_idx, 1, QTableWidgetItem(row.get("备案域名", "")))
            table.setItem(row_idx, 2, QTableWidgetItem(row.get("备案IP", "")))
            table.setItem(row_idx, 3, QTableWidgetItem(row.get("微信小程序", "")))
            table.setItem(row_idx, 4, QTableWidgetItem(row.get("备案APP", "")))
            table.setRowHeight(row_idx, 30)
        def is_success(r):
            for key in ["备案域名", "备案IP", "微信小程序", "备案APP"]:
                val = r.get(key, "")
                if not val or "查询异常" in val or "企业名称错误" in val:
                    return False
            return True
        if hasattr(self, "tab_results"):
            if idx not in self.tab_results:
                self.tab_results[idx] = []
            old = None
            for r in self.tab_results[idx]:
                if r["企业名称"] == row["企业名称"]:
                    old = r
                    break
            if old:
                old_success = is_success(old)
                new_success = is_success(row)
                if not old_success and new_success:
                    self.tab_results[idx] = [r for r in self.tab_results[idx] if r["企业名称"] != row["企业名称"]]
                    self.tab_results[idx].append(row)
                elif old_success and new_success:
                    self.tab_results[idx] = [r for r in self.tab_results[idx] if r["企业名称"] != row["企业名称"]]
                    self.tab_results[idx].append(row)
                elif old_success and not new_success:
                    return
                elif not old_success and not new_success:
                    self.tab_results[idx] = [r for r in self.tab_results[idx] if r["企业名称"] != row["企业名称"]]
                    self.tab_results[idx].append(row)
            else:
                self.tab_results[idx].append(row)

    def search_finished(self):
        self.search_btn_main.setEnabled(True)
        self.export_btn.setEnabled(True)

    def show_error(self, msg):
        QMessageBox.critical(self, "错误", msg)
        self.search_btn_main.setEnabled(True)

    def export_results(self):
        idx = self.tabs.currentIndex()
        if not hasattr(self, "tab_results") or idx not in self.tab_results or not self.tab_results[idx]:
            QMessageBox.warning(self, "警告", "没有可导出的结果")
            return
        try:
            df = pd.DataFrame(self.tab_results[idx])
            columns = ["企业名称", "备案域名", "备案IP", "微信小程序", "备案APP"]
            df = df.reindex(columns=columns)
            first_name = self.search_input.text().strip()
            if not first_name:
                first_name = "result"
            import time
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            filename = f"{first_name.split(',')[0]}_{timestamp}.xlsx"
            import importlib
            icpmod = importlib.import_module("icpsearch_icp")
            icpmod.save_results(df.to_dict(orient="records"), filename, show_message=False)
            QMessageBox.information(self, "成功", f"结果已导出到 {filename}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")

    def find_company(self):
        name = self.search_input.text().strip()
        if not name:
            return
        idx = self.tabs.currentIndex()
        tab_widget = self.tabs.widget(idx)
        if tab_widget is None:
            return
        table = tab_widget.findChild(QTableWidget, "result_table")
        if table is None:
            return
        if not hasattr(self, '_search_matches') or self._search_name != name or self._search_tab != idx:
            self._search_matches = [row for row in range(table.rowCount()) if table.item(row, 0) and name in table.item(row, 0).text()]
            self._search_idx = 0
            self._search_name = name
            self._search_tab = idx
        if not self._search_matches:
            QMessageBox.information(self, "查找结果", f"未找到企业名：{name}")
            return
        sender = self.sender()
        if sender == self.down_btn:
            self._search_idx = (self._search_idx + 1) % len(self._search_matches)
        row = self._search_matches[self._search_idx]
        table.selectRow(row)
        table.scrollToItem(table.item(row, 0))

    def auto_adjust(self):
        idx = self.tabs.currentIndex()
        tab_widget = self.tabs.widget(idx)
        if tab_widget is not None:
            table = tab_widget.findChild(QTableWidget, "result_table")
            if table is not None:
                table.resizeColumnsToContents()
                table.resizeRowsToContents()
                table.setColumnWidth(0, 320)
                for i in range(1, 5):
                    table.setColumnWidth(i, 160)

    def add_tab(self):
        table = QTableWidget(0, 5)
        table.setObjectName("result_table")
        table.setHorizontalHeaderLabels(["企业名称", "备案域名", "备案IP", "微信小程序", "备案APP"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        table.horizontalHeader().setStretchLastSection(True)
        table.setColumnWidth(0, 320)
        for i in range(1, 5):
            table.setColumnWidth(i, 160)
        table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(self.show_table_context_menu)
        tab = QWidget()
        vbox = QVBoxLayout(tab)
        vbox.setContentsMargins(0, 0, 0, 0)
        vbox.addWidget(table)
        tab.setLayout(vbox)
        tab_idx = self.tabs.addTab(tab, f"标签{self.tabs.count() + 1}")
        self.tabs.setCurrentIndex(tab_idx)
        if not hasattr(self, "tab_results"):
            self.tab_results = {}
        self.tab_results[tab_idx] = []
        if hasattr(self, "manual_input"):
            self.manual_input.clear()
        if hasattr(self, "file_label"):
            self.file_label.setText("未选择文件")
        if hasattr(self, "input_file"):
            self.input_file = ""
        if hasattr(self, "input_type"):
            self.input_type = ""

    def show_table_context_menu(self, pos):
        table = self.sender()
        menu = QMenu()
        del_action = menu.addAction("删除此企业行")
        copy_action = menu.addAction("复制此单元格")
        fofa_action = menu.addAction("生成FOFA语法")
        hunter_action = menu.addAction("生成Hunter语法")
        action = menu.exec_(table.viewport().mapToGlobal(pos))
        if action == del_action:
            row = table.currentRow()
            if row >= 0:
                table.removeRow(row)
        elif action == copy_action:
            row = table.currentRow()
            col = table.currentColumn()
            if row >= 0 and col >= 0:
                item = table.item(row, col)
                if item:
                    QApplication.clipboard().setText(item.text())
        elif action == fofa_action:
            row = table.currentRow()
            if row < 0:
                QMessageBox.warning(self, "提示", "请先选中一行")
                return
            domain_item = table.item(row, 1)
            ip_item = table.item(row, 2)
            domains = set()
            ips = set()
            if domain_item:
                for d in domain_item.text().split("\n"):
                    d = d.strip()
                    if d:
                        domains.add(d)
            if ip_item:
                for ip in ip_item.text().split("\n"):
                    ip = ip.strip()
                    if ip:
                        ips.add(ip)
            fofa_parts = []
            for d in domains:
                fofa_parts.append(f'domain="{d}"')
            for ip in ips:
                fofa_parts.append(f'ip="{ip}"')
            fofa_str = " || ".join(fofa_parts)
            QApplication.clipboard().setText(fofa_str)
            QMessageBox.information(self, "FOFA语法", "FOFA语法已复制到剪贴板！\n\n" + fofa_str)
        elif action == hunter_action:
            row = table.currentRow()
            if row < 0:
                QMessageBox.warning(self, "提示", "请先选中一行")
                return
            name_item = table.item(row, 0)
            domain_item = table.item(row, 1)
            ip_item = table.item(row, 2)
            names = set()
            domains = set()
            ips = set()
            if name_item:
                for n in name_item.text().split("\n"):
                    n = n.strip()
                    if n:
                        names.add(n)
            if domain_item:
                for d in domain_item.text().split("\n"):
                    d = d.strip()
                    if d:
                        domains.add(d)
            if ip_item:
                for ip in ip_item.text().split("\n"):
                    ip = ip.strip()
                    if ip:
                        ips.add(ip)
            hunter_parts = []
            for d in domains:
                hunter_parts.append(f'domain="{d}"')
            for ip in ips:
                hunter_parts.append(f'ip="{ip}"')
            for n in names:
                hunter_parts.append(f'icp.name="{n}"')
            hunter_str = " || ".join(hunter_parts)
            QApplication.clipboard().setText(hunter_str)
            QMessageBox.information(self, "Hunter语法", "Hunter语法已复制到剪贴板！\n\n" + hunter_str)

    def del_tab(self):
        idx = self.tabs.currentIndex()
        if self.tabs.count() > 1 and idx != -1:
            self.tabs.removeTab(idx)
            if hasattr(self, "tab_results"):
                self.tab_results.pop(idx, None)
        else:
            tab_widget = self.tabs.widget(0)
            if tab_widget is not None:
                table = tab_widget.findChild(QTableWidget, "result_table")
                if table is not None:
                    table.setRowCount(0)
            self.tabs.setTabText(0, "标签1")
            if hasattr(self, "tab_results"):
                self.tab_results[0] = []
            if hasattr(self, "manual_input"):
                self.manual_input.clear()
            if hasattr(self, "file_label"):
                self.file_label.setText("未选择文件")
            if hasattr(self, "input_file"):
                self.input_file = ""
            if hasattr(self, "input_type"):
                self.input_type = ""

    def rename_tab(self, idx):
        if idx == -1:
            return
        old_name = self.tabs.tabText(idx)
        new_name, ok = QInputDialog.getText(self, "重命名", "新标签名：", text=old_name)
        if ok and new_name.strip():
            self.tabs.setTabText(idx, new_name.strip())

    def on_tab_changed(self, idx):
        if hasattr(self, "manual_input"):
            self.manual_input.clear()
        if hasattr(self, "tab_results"):
            self.current_results = self.tab_results.get(idx, [])

    def close_tab(self, idx):
        self.del_tab()

    def import_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "选择企业名单txt文件", "", "Text Files (*.txt)")
        if not file_name:
            return
        try:
            with open(file_name, "r", encoding="utf-8") as f:
                lines = [line.strip() for line in f if line.strip()]
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取TXT失败：{str(e)}")
            return
        valid_lines = [line for line in lines if line]
        if not valid_lines:
            QMessageBox.warning(self, "导入格式错误", "导入格式错误，请检查txt文件内容，每行应为一个企业名称")
            return
        manual = self.manual_input.toPlainText().replace('，', ',').strip()
        manual_names = [x.strip() for x in manual.split(",") if x.strip()] if manual else []
        all_names = manual_names + valid_lines
        unique_names = list(dict.fromkeys(all_names))
        self.manual_input.setPlainText(",".join(unique_names))

    def get_company_list(self):
        if hasattr(self, "manual_input"):
            manual_text = self.manual_input.toPlainText().replace('，', ',').strip()
            manual = manual_text.split(",")
            manual = [x.strip() for x in manual if x.strip()]
        else:
            manual = []
        imported = []
        if hasattr(self, "input_file") and self.input_file:
            try:
                with open(self.input_file, "r", encoding="utf-8") as f:
                    imported = [line.strip() for line in f if line.strip()]
            except Exception:
                pass
        all_names = manual + imported
        seen = set()
        result = []
        for name in all_names:
            if name and name not in seen:
                seen.add(name)
                result.append(name)
        return result

    def on_ask_continue(self, msg):
        reply = QMessageBox.question(self, "是否继续查询", msg, QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            continue_signal.continue_result.emit('y')
        else:
            continue_signal.continue_result.emit('n') 