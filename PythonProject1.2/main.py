#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
装备维修器材出入库管理系统（合并稳定版）
所有界面、对话框均在本文件内，避免模块拆分造成的关闭崩溃
"""

import sys, os, uuid
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QPushButton, QLineEdit, QLabel,
    QComboBox, QFormLayout, QDialog, QMessageBox,
    QTabWidget, QGroupBox, QDateEdit, QDoubleSpinBox, QSpinBox,
    QDialogButtonBox, QFileDialog, QFrame, QInputDialog,
    QRadioButton
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# ==================== 导入数据库和单据模块 ====================
from database import Database
from receipt_exporter import ReceiptExporter

# ==================== 常量 ====================
UNITS = ['个', '台', '套', '件', '米', '千克', '升', '卷', '包', '箱']
QUALITY_LEVELS = ['新品', '堪用', '待修', '报废']
SOURCE_TYPES = ['本级采购', '上级采购', '航材库领取', '其他单位调拨', '厂修返回',
                '故障件拆机', '到寿件拆机', '装备升级拆机', '厂家赠与', '其它']
OUT_DIRECTIONS_ORG = ['下拨中队', '调拨其他单位', '送厂维修', '融通公司报废', '其它']
OUT_DIRECTIONS_UNIT = ['装机使用', '调拨其他单位', '送厂维修', '上交机关', '其它']

# 机关管理员账号
ORG_ACCOUNTS = {
    "焦汉学": "654321",
    "刘致岐": "266811"
}
# ==================== 登录对话框 ====================
class LoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("用户登录")
        self.setFixedSize(450, 400)
        self.role = None
        self.belong_unit = None
        self.operator = ''
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        title = QLabel("装备维修器材出入库管理系统")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size:18px;font-weight:bold;padding:15px;")
        layout.addWidget(title)

        form = QFormLayout()
        self.role_combo = QComboBox()
        self.role_combo.addItems(['机关器材管理员', '中队器材管理员'])
        self.role_combo.currentTextChanged.connect(self.on_role_changed)
        form.addRow("选择角色:", self.role_combo)

        self.unit_input = QLineEdit()
        self.unit_input.setPlaceholderText("请输入中队名称")
        self.unit_input.setEnabled(False)
        form.addRow("所属中队:", self.unit_input)

        self.account_input = QLineEdit()
        self.account_input.setPlaceholderText("机关管理员账号")
        form.addRow("账号:", self.account_input)

        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("密码")
        self.password_input.setEchoMode(QLineEdit.Password)
        form.addRow("密码:", self.password_input)

        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("请输入您的姓名（操作人）")
        form.addRow("操作人:", self.name_input)

        layout.addLayout(form)
        layout.addStretch()

        btn = QPushButton("进 入 系 统")
        btn.setMinimumHeight(40)
        btn.setStyleSheet("font-size:14px;font-weight:bold;background:#2196F3;color:white;")
        btn.clicked.connect(self.do_login)
        layout.addWidget(btn)

    def on_role_changed(self, text):
        is_org = (text == '机关器材管理员')
        self.unit_input.setEnabled(not is_org)
        self.account_input.setEnabled(is_org)
        self.password_input.setEnabled(is_org)

    def do_login(self):
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "提示", "请输入操作人姓名")
            return
        self.operator = name

        if self.role_combo.currentText() == '机关器材管理员':
            account = self.account_input.text().strip()
            pwd = self.password_input.text().strip()
            if account not in ORG_ACCOUNTS or ORG_ACCOUNTS[account] != pwd:
                QMessageBox.warning(self, "提示", "账号或密码错误")
                return
            self.role = 'org'
            self.belong_unit = ''
        else:
            unit = self.unit_input.text().strip()
            if not unit:
                QMessageBox.warning(self, "提示", "请输入所属中队名称")
                return
            self.role = 'unit'
            self.belong_unit = unit

        self.accept()


# ==================== 器材选择对话框 ====================
class EquipmentSelectDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.selected_equipment = None
        self.setWindowTitle("选择器材")
        self.setMinimumWidth(750)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        # 搜索栏
        sl = QHBoxLayout()
        sl.addWidget(QLabel("关键词:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("所属单位/器材名称/规格型号/编号")
        sl.addWidget(self.search_input)
        sl.addWidget(QLabel("质量:"))
        self.quality_combo = QComboBox()
        self.quality_combo.addItems(['全部'] + QUALITY_LEVELS)
        sl.addWidget(self.quality_combo)
        btn = QPushButton("搜索")
        btn.clicked.connect(self.do_search)
        sl.addWidget(btn)
        layout.addLayout(sl)

        # 器材列表表格
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(
            ['器材编号', '器材名称', '规格型号', '所属单位', '单价', '库存'])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)

        # 按钮
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.select)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

        self.do_search()

    def do_search(self):
        kw = self.search_input.text().strip()
        q = self.quality_combo.currentText()
        if q == '全部':
            q = ''

        sql = """
              SELECT e.*,
                     COALESCE(si.total_in, 0) - COALESCE(so.total_out, 0) as current_stock
              FROM equipment e
                       LEFT JOIN (SELECT equipment_id, SUM(quantity) as total_in
                                  FROM org_stock_in
                                  GROUP BY equipment_id) si
                                 ON e.id = si.equipment_id
                       LEFT JOIN (SELECT equipment_id, SUM(quantity) as total_out
                                  FROM org_stock_out
                                  GROUP BY equipment_id) so
                                 ON e.id = so.equipment_id
              WHERE 1 = 1
              """
        params = []
        if kw:
            sql += " AND (e.belong_unit LIKE ? OR e.name LIKE ? OR e.spec_model LIKE ? OR e.serial_no LIKE ?)"
            kw_pattern = f"%{kw}%"
            params.extend([kw_pattern, kw_pattern, kw_pattern, kw_pattern])
        if q:
            # 改为模糊匹配，解决“堪用品”搜索“堪用”找不到的问题
            sql += " AND e.quality_level LIKE ?"
            params.append(f"%{q}%")

        sql += " ORDER BY e.updated_at DESC LIMIT 500"

        self.db.cursor.execute(sql, params)
        records = self.db.cursor.fetchall()

        self.table.setRowCount(len(records))
        for i, r in enumerate(records):
            self.table.setItem(i, 0, QTableWidgetItem(r['serial_no'] or ''))
            self.table.setItem(i, 1, QTableWidgetItem(r['name'] or ''))
            self.table.setItem(i, 2, QTableWidgetItem(r['spec_model'] or ''))
            self.table.setItem(i, 3, QTableWidgetItem(r['belong_unit'] or ''))
            self.table.setItem(i, 4, QTableWidgetItem(
                f"¥{r['unit_price']:.2f}" if r['unit_price'] else ''))
            self.table.setItem(i, 5, QTableWidgetItem(str(int(r['current_stock']) or 0)))

    def select(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "提示", "请选择一种器材")
            return

        name = self.table.item(row, 1).text() if self.table.item(row, 1) else ''
        spec = self.table.item(row, 2).text() if self.table.item(row, 2) else ''

        records = self.db.search_equipment(keyword=name)
        found = None
        for rec in records:
            if rec['name'] == name and (rec['spec_model'] or '') == spec:
                found = rec
                break
        if not found and records:
            found = records[0]

        if found:
            self.selected_equipment = found
            self.accept()
        else:
            QMessageBox.warning(self, "提示", "未找到该器材的完整信息，请重新搜索选择")
# ==================== 器材编辑对话框 ====================
class EquipmentDialog(QDialog):
    def __init__(self, db, equip_id=None, parent=None):
        super().__init__(parent)
        self.db = db
        self.equip_id = equip_id
        self.setWindowTitle("编辑器材" if equip_id else "新增器材")
        self.setMinimumWidth(550)
        self.init_ui()
        if equip_id:
            self.load_data()

    def init_ui(self):
        layout = QFormLayout(self)

        self.serial_no = QLineEdit()
        self.serial_no.setPlaceholderText("每件器材唯一编号")
        layout.addRow("器材编号:", self.serial_no)

        self.name = QLineEdit()
        self.name.setPlaceholderText("(必填)")
        layout.addRow("器材名称*:", self.name)

        self.spec_model = QLineEdit()
        self.spec_model.setPlaceholderText("(必填)")
        layout.addRow("规格型号*:", self.spec_model)

        self.category_code = QLineEdit()
        self.category_code.setPlaceholderText("国军标/行业标准编码")
        layout.addRow("品种标识码:", self.category_code)

        self.unit = QComboBox()
        self.unit.addItems(UNITS)
        self.unit.setEditable(True)
        layout.addRow("计量单位:", self.unit)

        self.quality_level = QComboBox()
        self.quality_level.addItems(QUALITY_LEVELS)
        layout.addRow("质量等级:", self.quality_level)

        self.manufacturer = QLineEdit()
        layout.addRow("生产厂家:", self.manufacturer)

        self.unit_price = QDoubleSpinBox()
        self.unit_price.setRange(0, 99999999.99)
        self.unit_price.setDecimals(2)
        self.unit_price.setPrefix("¥ ")
        layout.addRow("单价:", self.unit_price)

        self.production_date = QDateEdit()
        self.production_date.setCalendarPopup(True)
        self.production_date.setDate(QDate.currentDate())
        layout.addRow("生产日期:", self.production_date)

        self.storage_life = QSpinBox()
        self.storage_life.setRange(0, 9999)
        self.storage_life.setSuffix(" 个月")
        self.storage_life.setSpecialValueText("无")
        layout.addRow("存储寿命:", self.storage_life)

        self.source_type = QComboBox()
        self.source_type.addItems([''] + SOURCE_TYPES)
        self.source_type.setEditable(True)
        layout.addRow("来源类别:", self.source_type)

        self.pricing_method = QLineEdit()
        layout.addRow("计价方法:", self.pricing_method)

        self.contract_no = QLineEdit()
        layout.addRow("合同编号:", self.contract_no)

        self.belong_unit = QLineEdit()
        self.belong_unit.setPlaceholderText("请输入所属单位名称")
        layout.addRow("所属单位:", self.belong_unit)

        self.belong_equipment = QLineEdit()
        layout.addRow("所属装备:", self.belong_equipment)

        bl = QHBoxLayout()
        btn_save = QPushButton("保存")
        btn_save.clicked.connect(self.save)
        btn_cancel = QPushButton("取消")
        btn_cancel.clicked.connect(self.reject)
        bl.addStretch()
        bl.addWidget(btn_save)
        bl.addWidget(btn_cancel)
        layout.addRow(bl)

    def load_data(self):
        equip = self.db.get_equipment_by_id(self.equip_id)
        if not equip:
            return
        self.serial_no.setText(equip['serial_no'] or '')
        self.name.setText(equip['name'] or '')
        self.spec_model.setText(equip['spec_model'] or '')
        self.category_code.setText(equip['category_code'] or '')
        self.unit.setCurrentText(equip['unit'] or '')
        self.quality_level.setCurrentText(equip['quality_level'] or '新品')
        self.manufacturer.setText(equip['manufacturer'] or '')
        self.unit_price.setValue(equip['unit_price'] or 0)
        if equip['production_date']:
            try:
                d = datetime.strptime(equip['production_date'], '%Y-%m-%d')
                self.production_date.setDate(QDate(d.year, d.month, d.day))
            except:
                pass
        self.storage_life.setValue(equip['storage_life'] or 0)
        self.source_type.setCurrentText(equip['source_type'] or '')
        self.pricing_method.setText(equip['pricing_method'] or '')
        self.contract_no.setText(equip['contract_no'] or '')
        self.belong_unit.setText(equip['belong_unit'] or '')
        self.belong_equipment.setText(equip['belong_equipment'] or '')

    def save(self):
        name = self.name.text().strip()
        spec_model = self.spec_model.text().strip()
        if not name or not spec_model:
            QMessageBox.warning(self, "提示", "器材名称和规格型号为必填项")
            return
        data = {
            'serial_no': self.serial_no.text().strip(),
            'name': name,
            'spec_model': spec_model,
            'category_code': self.category_code.text().strip(),
            'unit': self.unit.currentText().strip(),
            'quality_level': self.quality_level.currentText(),
            'manufacturer': self.manufacturer.text().strip(),
            'unit_price': self.unit_price.value(),
            'production_date': self.production_date.date().toString('yyyy-MM-dd'),
            'storage_life': self.storage_life.value() if self.storage_life.value() > 0 else None,
            'source_type': self.source_type.currentText().strip(),
            'pricing_method': self.pricing_method.text().strip(),
            'contract_no': self.contract_no.text().strip(),
            'belong_unit': self.belong_unit.text().strip(),
            'belong_equipment': self.belong_equipment.text().strip(),
        }
        try:
            if self.equip_id:
                self.db.update_equipment(self.equip_id, data)
            else:
                self.db.add_equipment(data)
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存失败：{str(e)}")


# ==================== 入库单分组选择对话框 ====================
class ReceiptGroupDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择入库单生成方式")
        self.setFixedSize(400, 200)
        self.group_type = 'factory'
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("请选择入库单生成方式："))
        self.btn_factory = QRadioButton("按生产厂家分组生成（每个厂家一份入库单）")
        self.btn_factory.setChecked(True)
        self.btn_date = QRadioButton("按入库日期生成（同一天入库合并为一份入库单）")
        layout.addWidget(self.btn_factory)
        layout.addWidget(self.btn_date)
        bl = QHBoxLayout()
        btn_ok = QPushButton("确定")
        btn_ok.clicked.connect(self.do_ok)
        btn_cancel = QPushButton("取消")
        btn_cancel.clicked.connect(self.reject)
        bl.addWidget(btn_ok)
        bl.addWidget(btn_cancel)
        layout.addLayout(bl)

    def do_ok(self):
        self.group_type = 'factory' if self.btn_factory.isChecked() else 'date'
        self.accept()


# ==================== 出库单分组选择对话框 ====================
class OutReceiptGroupDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择出库单生成方式")
        self.setFixedSize(400, 200)
        self.group_type = 'equipment'
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("请选择出库单生成方式："))
        self.btn_equip = QRadioButton("按所属装备分组生成")
        self.btn_equip.setChecked(True)
        self.btn_date = QRadioButton("按出库日期生成")
        layout.addWidget(self.btn_equip)
        layout.addWidget(self.btn_date)
        bl = QHBoxLayout()
        btn_ok = QPushButton("确定")
        btn_ok.clicked.connect(self.do_ok)
        btn_cancel = QPushButton("取消")
        btn_cancel.clicked.connect(self.reject)
        bl.addWidget(btn_ok)
        bl.addWidget(btn_cancel)
        layout.addLayout(bl)

    def do_ok(self):
        self.group_type = 'equipment' if self.btn_equip.isChecked() else 'date'
        self.accept()

# ==================== 机关入库登记对话框 ====================
class OrgStockInDialog(QDialog):
    """
    机关器材入库登记对话框
    支持添加多条器材明细，数量由 QDoubleSpinBox 输入，自动计算总价
    保存时写入数据库，记录详细日志，并可生成本次入库的 Excel 单据
    """
    def __init__(self, db, operator_name='', parent=None):
        super().__init__(parent)
        self.db = db                        # 数据库操作对象
        self.operator_name = operator_name  # 当前操作人姓名
        self.setWindowTitle("机关入库登记")
        self.setMinimumSize(850, 650)
        self.init_ui()

    def init_ui(self):
        """构建界面布局"""
        layout = QVBoxLayout(self)

        # ---------- 基本信息区域 ----------
        info_group = QGroupBox("入库基本信息")
        info_form = QFormLayout(info_group)

        # 入库日期（默认今天）
        self.in_time = QDateEdit()
        self.in_time.setCalendarPopup(True)        # 点击弹出日历
        self.in_time.setDate(QDate.currentDate())
        info_form.addRow("入库日期:", self.in_time)

        # 入库人（自动填充操作人姓名）
        self.stock_in_person = QLineEdit()
        self.stock_in_person.setText(self.operator_name)
        info_form.addRow("入库人:", self.stock_in_person)

        # 经办人（自动填充操作人姓名）
        self.operator = QLineEdit()
        self.operator.setText(self.operator_name)
        info_form.addRow("经办人:", self.operator)

        layout.addWidget(info_group)

        # ---------- 器材明细区域 ----------
        detail_group = QGroupBox("器材明细(可一次录入多条)")
        detail_layout = QVBoxLayout(detail_group)

        # 添加/删除行按钮
        bl = QHBoxLayout()
        btn_add = QPushButton("＋ 添加一行器材")
        btn_add.clicked.connect(self.add_item)
        btn_del = QPushButton("－ 删除选中行")
        btn_del.clicked.connect(self.del_item)
        bl.addWidget(btn_add)
        bl.addWidget(btn_del)
        bl.addStretch()                         # 弹性空间，让按钮靠左
        detail_layout.addLayout(bl)

        # 器材明细表格：7列
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(
            ['器材编号', '器材名称', '规格型号', '数量', '单价', '总价', '备注'])
        self.table.horizontalHeader().setStretchLastSection(True)  # 备注列自动拉伸
        self.table.setSelectionBehavior(QTableWidget.SelectRows)   # 整行选择
        self.table.setMinimumHeight(250)
        detail_layout.addWidget(self.table)

        layout.addWidget(detail_group)

        # ---------- 确定/取消按钮 ----------
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.save_all)      # 点确定 -> 保存
        bb.rejected.connect(self.reject)        # 点取消 -> 关闭
        layout.addWidget(bb)

    def add_item(self):
        """
        添加一条器材到明细表格
        先弹出器材选择对话框，用户选择后自动填充器材编号、名称、规格型号、单价，
        数量由数字输入框（QDoubleSpinBox）提供，默认值为1，总价自动计算。
        """
        dlg = EquipmentSelectDialog(self.db, parent=self)
        if dlg.exec_() == QDialog.Accepted and dlg.selected_equipment:
            e = dlg.selected_equipment          # 获取选中的器材信息
            row = self.table.rowCount()         # 新行的索引
            self.table.insertRow(row)

            # 第1列：器材编号（隐藏存储器材ID，用于后续保存）
            item0 = QTableWidgetItem(e['serial_no'] or '')
            item0.setData(Qt.UserRole, e['id']) # Qt.UserRole 用于存储 ID
            self.table.setItem(row, 0, item0)

            # 第2列：器材名称
            self.table.setItem(row, 1, QTableWidgetItem(e['name']))

            # 第3列：规格型号
            self.table.setItem(row, 2, QTableWidgetItem(e['spec_model'] or ''))

            # 第4列：数量（数字输入框，嵌入表格）
            qty_spin = QSpinBox()
            qty_spin.setRange(1, 999999)     # 范围 1 ~ 999999
            #qty_spin.setDecimals(2)             # 保留两位小数
            qty_spin.setValue(1)                # 默认数量 1
            # 连接信号：数值变化时调用 calc_row 重新计算总价，lambda 捕获当前行号
            qty_spin.valueChanged.connect(lambda v, r=row: self.calc_row(r))
            self.table.setCellWidget(row, 3, qty_spin)

            # 第5列：单价（仅显示，实际数值存储在 UserRole 中）
            price_item = QTableWidgetItem(f"{e['unit_price']:.2f}")
            price_item.setData(Qt.UserRole, e['unit_price'] or 0)
            self.table.setItem(row, 4, price_item)

            # 第6列：总价（初始 = 单价 × 1）
            self.table.setItem(row, 5, QTableWidgetItem(f"{e['unit_price']:.2f}"))

            # 第7列：备注（可编辑）
            self.table.setItem(row, 6, QTableWidgetItem(''))

    def calc_row(self, row):
        """
        计算指定行的总价 = 数量 × 单价
        当数量输入框的值改变时触发
        """
        qty_widget = self.table.cellWidget(row, 3)    # 获取第4列的数量输入框
        if qty_widget and self.table.item(row, 4):    # 确保存在单价信息
            qty = qty_widget.value()
            price = self.table.item(row, 4).data(Qt.UserRole) or 0
            total = qty * price
            self.table.item(row, 5).setText(f"{total:.2f}")  # 更新总价显示

    def del_item(self):
        """删除当前选中的行"""
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

    def save_all(self):
        """
        保存所有入库记录到数据库
        从表格中逐行读取数据，写入 org_stock_in 表，
        同时生成入库单 Excel 文件，并记录详细日志
        """
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "提示", "请至少添加一条器材明细")
            return

        # 获取公共信息
        in_time = self.in_time.date().toString('yyyy-MM-dd')
        stock_in_person = self.stock_in_person.text().strip()
        operator = self.operator.text().strip()

        # 生成唯一批次号（时间戳 + 随机串）
        batch_no = datetime.now().strftime('%Y%m%d%H%M%S') + '_' + uuid.uuid4().hex[:6]

        all_items = []          # 存储所有入库明细，用于生成入库单

        try:
            # 遍历每一行
            for row in range(self.table.rowCount()):
                equip_id = self.table.item(row, 0).data(Qt.UserRole)   # 器材ID
                qty = self.table.cellWidget(row, 3).value()            # 数量
                unit_price = self.table.item(row, 4).data(Qt.UserRole) # 单价
                total_price = qty * unit_price
                remark = self.table.item(row, 6).text() if self.table.item(row, 6) else ''

                # 构造入库记录数据
                data = {
                    'equipment_id': equip_id,
                    'quantity': qty,
                    'total_price': total_price,
                    'in_time': in_time,
                    'stock_in_person': stock_in_person,
                    'operator': operator,
                    'remark': remark,
                    'batch_no': batch_no
                }
                self.db.add_org_stock_in(data)      # 保存到数据库

                # 获取器材完整信息（用于生成入库单）
                equip_full = self.db.get_equipment_by_id(equip_id)
                all_items.append({
                    'serial_no': equip_full['serial_no'] if equip_full else '',
                    'name': self.table.item(row, 1).text(),
                    'spec_model': self.table.item(row, 2).text(),
                    'belong_equipment': equip_full['belong_equipment'] if equip_full else '',
                    'quantity': qty,
                    'unit': equip_full['unit'] if equip_full else '',
                    'unit_price': unit_price,
                    'total_price': total_price,
                    'remark': remark,
                    'manufacturer': equip_full['manufacturer'] if equip_full else '',
                    'contract_no': equip_full['contract_no'] if equip_full else '',
                })

            # 记录详细日志（包含具体器材名称和数量）
            detail = "入库器材: " + ", ".join(
                f"{item['name']} {item['quantity']}{item['unit']}" for item in all_items
            )
            self.db.add_log('入库', '机关入库', None, detail,
                            operator=operator, role='org')

            # 弹出分组方式选择，并生成入库单
            dlg = ReceiptGroupDialog(self)
            if dlg.exec_() == QDialog.Accepted:
                self._generate_receipts(all_items, in_time, dlg.group_type)

            QMessageBox.information(self, "成功", f"已入库 {len(all_items)} 条器材记录")
            self.accept()       # 关闭对话框

        except Exception as e:
            QMessageBox.critical(self, "入库失败", f"保存入库记录时发生错误：{str(e)}")

    def _generate_receipts(self, all_items, in_date, group_type):
        """
        生成入库单 Excel 文件
        :param all_items: 入库明细列表
        :param in_date: 入库日期
        :param group_type: 分组方式 ('factory' 按厂家, 'date' 按日期)
        """
        # 输出目录：程序根目录下的“入库单”文件夹
        output_dir = os.path.join(os.path.dirname(__file__), '入库单')
        os.makedirs(output_dir, exist_ok=True)

        # 根据分组方式构建分组
        if group_type == 'factory':
            groups = {}                     # key: 厂家名, value: 该厂家下的器材列表
            for item in all_items:
                manufacturer = item.get('manufacturer', '未知厂家') or '未知厂家'
                contract_no = item.get('contract_no', '') or ''
                if manufacturer not in groups:
                    groups[manufacturer] = {'items': [], 'contract': contract_no}
                groups[manufacturer]['items'].append(item)
        else:   # 按日期分组
            contract_no = all_items[0].get('contract_no', '') if all_items else ''
            groups = {in_date: {'items': all_items, 'contract': contract_no}}

        total_files = 0
        for group_name, data in groups.items():
            try:
                files = ReceiptExporter.generate_in_receipt(
                    items=data['items'],
                    output_dir=output_dir,
                    group_name=group_name,
                    contract_no=data['contract'],
                    in_date=in_date,
                    stock_in_person='',
                    operator='',
                    group_type=group_type
                )
                total_files += len(files)
            except Exception as e:
                QMessageBox.critical(self, "生成入库单失败", str(e))
                return

        if total_files > 0:
            QMessageBox.information(self, "入库单已生成",
                                    f"已自动生成 {total_files} 个入库单文件\n保存位置：{output_dir}")

# ==================== 机关出库登记对话框 ====================
class OrgStockOutDialog(QDialog):
    """
    机关器材出库登记对话框（发放给中队）
    支持添加多条器材明细，显示当前库存，数量由 QDoubleSpinBox 输入，自动计算总价。
    保存时写入数据库，并自动为接收中队生成一条入库记录（来源：机关下拨），
    同时记录详细日志，并可生成本次出库的 Excel 单据。
    """
    def __init__(self, db, operator_name='', parent=None):
        super().__init__(parent)
        self.db = db
        self.operator_name = operator_name
        self.setWindowTitle("机关出库登记(发放给中队)")
        self.setMinimumSize(850, 650)
        self.init_ui()

    def init_ui(self):
        """构建界面布局"""
        layout = QVBoxLayout(self)

        # ---------- 基本信息区域 ----------
        info_group = QGroupBox("出库基本信息")
        info_form = QFormLayout(info_group)

        # 接收中队（必填）
        self.to_unit = QLineEdit()
        self.to_unit.setPlaceholderText("请输入接收中队名称")
        info_form.addRow("发放中队:", self.to_unit)

        # 出库去向（下拉列表，可自定义输入）
        self.out_direction = QComboBox()
        self.out_direction.addItems(OUT_DIRECTIONS_ORG)
        self.out_direction.setEditable(True)
        info_form.addRow("出库去向:", self.out_direction)

        # 出库日期（默认今天）
        self.out_time = QDateEdit()
        self.out_time.setCalendarPopup(True)
        self.out_time.setDate(QDate.currentDate())
        info_form.addRow("出库日期:", self.out_time)

        # 出库人（自动填充操作人姓名）
        self.stock_out_person = QLineEdit()
        self.stock_out_person.setText(self.operator_name)
        info_form.addRow("出库人:", self.stock_out_person)

        # 经办人（自动填充操作人姓名）
        self.operator = QLineEdit()
        self.operator.setText(self.operator_name)
        info_form.addRow("经办人:", self.operator)

        layout.addWidget(info_group)

        # ---------- 器材明细区域 ----------
        detail_group = QGroupBox("器材明细")
        detail_layout = QVBoxLayout(detail_group)

        # 添加/删除行按钮
        bl = QHBoxLayout()
        btn_add = QPushButton("＋ 添加一行器材")
        btn_add.clicked.connect(self.add_item)
        btn_del = QPushButton("－ 删除选中行")
        btn_del.clicked.connect(self.del_item)
        bl.addWidget(btn_add)
        bl.addWidget(btn_del)
        bl.addStretch()
        detail_layout.addLayout(bl)

        # 器材明细表格：8列（多了一列“库存”）
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels(
            ['器材编号', '器材名称', '规格型号', '库存', '出库数量', '单价', '总价', '备注'])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setMinimumHeight(250)
        detail_layout.addWidget(self.table)

        layout.addWidget(detail_group)

        # ---------- 确定/取消按钮 ----------
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.save_all)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

    def add_item(self):
        """
        添加一条器材明细。
        弹出器材选择对话框，用户选择后显示器材信息和当前机关库存，
        出库数量通过数字输入框提供，默认值为1，总价自动计算。
        """
        dlg = EquipmentSelectDialog(self.db, parent=self)
        if dlg.exec_() == QDialog.Accepted and dlg.selected_equipment:
            e = dlg.selected_equipment
            stock = self.db.get_org_current_stock(e['id'])   # 获取该器材的机关库存
            row = self.table.rowCount()
            self.table.insertRow(row)

            # 器材编号（隐藏存储器材ID）
            item0 = QTableWidgetItem(e['serial_no'] or '')
            item0.setData(Qt.UserRole, e['id'])
            self.table.setItem(row, 0, item0)

            # 器材名称
            self.table.setItem(row, 1, QTableWidgetItem(e['name']))

            # 规格型号
            self.table.setItem(row, 2, QTableWidgetItem(e['spec_model'] or ''))

            # 库存（只读显示）
            self.table.setItem(row, 3, QTableWidgetItem(str(int(stock))))

            # 出库数量（数字输入框）
            qty_spin = QSpinBox()
            qty_spin.setRange(1, 999999)    #范围
            #qty_spin.setDecimals(2)     #小数点
            qty_spin.setValue(1)
            qty_spin.valueChanged.connect(lambda v, r=row: self.calc_row(r))
            self.table.setCellWidget(row, 4, qty_spin)

            # 单价（仅显示，实际数值存储在 UserRole）
            price_item = QTableWidgetItem(f"{e['unit_price']:.2f}")
            price_item.setData(Qt.UserRole, e['unit_price'] or 0)
            self.table.setItem(row, 5, price_item)

            # 总价（初始 = 单价 × 1）
            self.table.setItem(row, 6, QTableWidgetItem(f"{e['unit_price']:.2f}"))

            # 备注（可编辑）
            self.table.setItem(row, 7, QTableWidgetItem(''))

    def calc_row(self, row):
        """计算指定行的总价 = 数量 × 单价"""
        qty_widget = self.table.cellWidget(row, 4)    # 出库数量在第5列
        if qty_widget and self.table.item(row, 5):
            qty = qty_widget.value()
            price = self.table.item(row, 5).data(Qt.UserRole) or 0
            total = qty * price
            self.table.item(row, 6).setText(f"{total:.2f}")

    def del_item(self):
        """删除当前选中的行"""
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

    def save_all(self):
        """
        保存所有出库记录到数据库。
        对每一行进行库存检查，若出库数量 > 当前库存则提示并中止。
        写入 org_stock_out 表，同时自动为接收中队添加 unit_stock_record (入库记录)。
        记录详细日志，并可生成本次出库的 Excel 单据。
        """
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "提示", "请至少添加一条器材明细")
            return

        to_unit = self.to_unit.text().strip()
        if not to_unit:
            QMessageBox.warning(self, "提示", "请输入接收中队名称")
            return

        out_direction = self.out_direction.currentText().strip()
        out_time = self.out_time.date().toString('yyyy-MM-dd')
        stock_out_person = self.stock_out_person.text().strip()
        operator = self.operator.text().strip()
        batch_no = datetime.now().strftime('%Y%m%d%H%M%S') + '_' + uuid.uuid4().hex[:6]

        all_items = []

        try:
            for row in range(self.table.rowCount()):
                equip_id = self.table.item(row, 0).data(Qt.UserRole)
                qty = self.table.cellWidget(row, 4).value()
                # 库存显示在表格中，取文本转为整数
                stock = int(self.table.item(row, 3).text())

                # 库存不足检查
                if qty > stock:
                    QMessageBox.warning(self, "库存不足",
                                        f"器材 {self.table.item(row, 1).text()} 库存仅 {stock}，无法出库 {qty}")
                    return

                unit_price = self.table.item(row, 5).data(Qt.UserRole)
                total_price = qty * unit_price
                remark = self.table.item(row, 7).text() if self.table.item(row, 7) else ''

                # 写入机关出库记录
                out_data = {
                    'equipment_id': equip_id,
                    'to_unit': to_unit,
                    'out_direction': out_direction,
                    'quantity': qty,
                    'total_price': total_price,
                    'out_time': out_time,
                    'stock_out_person': stock_out_person,
                    'operator': operator,
                    'remark': remark,
                    'batch_no': batch_no
                }
                self.db.add_org_stock_out(out_data)

                # 自动为接收中队添加一条入库记录（来源标记为“机关下拨”）
                unit_in_data = {
                    'equipment_id': equip_id,
                    'belong_unit': to_unit,
                    'record_type': 'in',
                    'quantity': qty,
                    'total_price': total_price,
                    'in_time': out_time,
                    'source': '机关下拨',
                    'stock_in_person': stock_out_person,
                    'operator': operator,
                    'remark': remark,
                    'batch_no': batch_no
                }
                self.db.add_unit_record(unit_in_data)

                # 收集用于生成单据的信息
                equip_full = self.db.get_equipment_by_id(equip_id)
                all_items.append({
                    'serial_no': equip_full['serial_no'] if equip_full else '',
                    'name': self.table.item(row, 1).text(),
                    'spec_model': self.table.item(row, 2).text(),
                    'belong_equipment': equip_full['belong_equipment'] if equip_full else '',
                    'quantity': qty,
                    'unit': equip_full['unit'] if equip_full else '',
                    'unit_price': unit_price,
                    'total_price': total_price,
                    'remark': remark,
                    'manufacturer': equip_full['manufacturer'] if equip_full else '',
                    'contract_no': equip_full['contract_no'] if equip_full else '',
                })

            # 记录详细日志
            detail = f"发放中队:{to_unit} 出库器材: " + ", ".join(
                f"{item['name']} {item['quantity']}{item['unit']}" for item in all_items
            )
            self.db.add_log('出库', '机关出库', None, detail,
                            operator=operator, role='org')

            # 弹出分组方式选择，并生成出库单
            dlg = OutReceiptGroupDialog(self)
            if dlg.exec_() == QDialog.Accepted:
                self._generate_exporters(all_items, out_time, dlg.group_type)

            QMessageBox.information(self, "成功", f"已出库 {len(all_items)} 条器材记录到 {to_unit}")
            self.accept()

        except Exception as e:
            QMessageBox.critical(self, "出库失败", f"保存出库记录时发生错误：{str(e)}")

    def _generate_exporters(self, all_items, out_date, group_type):
        """
        生成出库单 Excel 文件
        :param all_items: 出库明细列表
        :param out_date: 出库日期
        :param group_type: 分组方式 ('equipment' 按所属装备, 'date' 按日期)
        """
        output_dir = os.path.join(os.path.dirname(__file__), '出库单')
        os.makedirs(output_dir, exist_ok=True)

        if group_type == 'equipment':
            groups = {}
            for item in all_items:
                eq = item.get('belong_equipment', '未知装备') or '未知装备'
                if eq not in groups:
                    groups[eq] = []
                groups[eq].append(item)
        else:
            groups = {out_date: all_items}

        total_files = 0
        for group_name, items in groups.items():
            first = items[0]
            try:
                files = ReceiptExporter.generate_out_receipt(
                    items=items,
                    output_dir=output_dir,
                    group_name=group_name,
                    contract_no=first.get('contract_no', ''),
                    out_date=out_date,
                    stock_out_person='',
                    operator='',
                    group_type=group_type
                )
                total_files += len(files)
            except Exception as e:
                QMessageBox.critical(self, "生成出库单失败", str(e))
                return

        if total_files > 0:
            QMessageBox.information(self, "出库单已生成",
                                    f"已自动生成 {total_files} 个出库单文件\n保存位置：{output_dir}")

# ==================== 中队入库登记对话框 ====================
class UnitStockInDialog(QDialog):
    """
    中队器材入库登记对话框
    与机关入库类似，但入库来源可以自定义（如本级采购、上级调拨等），
    数量由 QDoubleSpinBox 输入，自动计算总价。
    保存时写入 unit_stock_record 表（record_type='in'），
    并生成入库单 Excel 文件。
    """
    def __init__(self, db, belong_unit, operator_name='', parent=None):
        super().__init__(parent)
        self.db = db
        self.belong_unit = belong_unit          # 当前中队名称
        self.operator_name = operator_name
        self.setWindowTitle(f"中队入库登记 - {belong_unit}")
        self.setMinimumSize(850, 650)
        self.init_ui()

    def init_ui(self):
        """构建界面布局"""
        layout = QVBoxLayout(self)

        # ---------- 基本信息区域 ----------
        info_group = QGroupBox("入库基本信息")
        info_form = QFormLayout(info_group)

        # 入库日期（默认今天）
        self.in_time = QDateEdit()
        self.in_time.setCalendarPopup(True)
        self.in_time.setDate(QDate.currentDate())
        info_form.addRow("入库日期:", self.in_time)

        # 来源类别（下拉列表 + 可自定义）
        self.source = QComboBox()
        self.source.addItems(SOURCE_TYPES)
        self.source.setEditable(True)
        info_form.addRow("来源类别:", self.source)

        # 入库人（自动填充）
        self.stock_in_person = QLineEdit()
        self.stock_in_person.setText(self.operator_name)
        info_form.addRow("入库人:", self.stock_in_person)

        # 经办人（自动填充）
        self.operator = QLineEdit()
        self.operator.setText(self.operator_name)
        info_form.addRow("经办人:", self.operator)

        layout.addWidget(info_group)

        # ---------- 器材明细区域 ----------
        detail_group = QGroupBox("器材明细(可一次录入多条)")
        detail_layout = QVBoxLayout(detail_group)

        bl = QHBoxLayout()
        btn_add = QPushButton("＋ 添加一行器材")
        btn_add.clicked.connect(self.add_item)
        btn_del = QPushButton("－ 删除选中行")
        btn_del.clicked.connect(self.del_item)
        bl.addWidget(btn_add)
        bl.addWidget(btn_del)
        bl.addStretch()
        detail_layout.addLayout(bl)

        # 器材明细表格：6列（没有备注列，可根据需要调整）
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(
            ['器材编号', '器材名称', '规格型号', '数量', '单价', '总价'])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setMinimumHeight(250)
        detail_layout.addWidget(self.table)

        layout.addWidget(detail_group)

        # ---------- 确定/取消按钮 ----------
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.save_all)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

    def add_item(self):
        """添加一条器材明细"""
        dlg = EquipmentSelectDialog(self.db, parent=self)
        if dlg.exec_() == QDialog.Accepted and dlg.selected_equipment:
            e = dlg.selected_equipment
            row = self.table.rowCount()
            self.table.insertRow(row)

            item0 = QTableWidgetItem(e['serial_no'] or '')
            item0.setData(Qt.UserRole, e['id'])
            self.table.setItem(row, 0, item0)
            self.table.setItem(row, 1, QTableWidgetItem(e['name']))
            self.table.setItem(row, 2, QTableWidgetItem(e['spec_model'] or ''))

            qty_spin = QDoubleSpinBox()
            qty_spin.setRange(1, 999999)
            #qty_spin.setDecimals(2)
            qty_spin.setValue(1)
            qty_spin.valueChanged.connect(lambda v, r=row: self.calc_row(r))
            self.table.setCellWidget(row, 3, qty_spin)

            price_item = QTableWidgetItem(f"{e['unit_price']:.2f}")
            price_item.setData(Qt.UserRole, e['unit_price'] or 0)
            self.table.setItem(row, 4, price_item)
            self.table.setItem(row, 5, QTableWidgetItem(f"{e['unit_price']:.2f}"))

    def calc_row(self, row):
        """计算总价 = 数量 × 单价"""
        qty_widget = self.table.cellWidget(row, 3)
        if qty_widget and self.table.item(row, 4):
            qty = qty_widget.value()
            price = self.table.item(row, 4).data(Qt.UserRole) or 0
            total = qty * price
            self.table.item(row, 5).setText(f"{total:.2f}")

    def del_item(self):
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

    def save_all(self):
        """保存所有入库记录"""
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "提示", "请至少添加一条器材明细")
            return

        in_time = self.in_time.date().toString('yyyy-MM-dd')
        source = self.source.currentText().strip()
        stock_in_person = self.stock_in_person.text().strip()
        operator = self.operator.text().strip()
        batch_no = datetime.now().strftime('%Y%m%d%H%M%S') + '_' + uuid.uuid4().hex[:6]

        all_items = []
        try:
            for row in range(self.table.rowCount()):
                equip_id = self.table.item(row, 0).data(Qt.UserRole)
                qty = self.table.cellWidget(row, 3).value()
                unit_price = self.table.item(row, 4).data(Qt.UserRole)
                total_price = qty * unit_price

                data = {
                    'equipment_id': equip_id,
                    'belong_unit': self.belong_unit,
                    'record_type': 'in',
                    'quantity': qty,
                    'total_price': total_price,
                    'in_time': in_time,
                    'source': source,
                    'stock_in_person': stock_in_person,
                    'operator': operator,
                    'remark': '',
                    'batch_no': batch_no
                }
                self.db.add_unit_record(data)

                equip_full = self.db.get_equipment_by_id(equip_id)
                all_items.append({
                    'serial_no': equip_full['serial_no'] if equip_full else '',
                    'name': self.table.item(row, 1).text(),
                    'spec_model': self.table.item(row, 2).text(),
                    'belong_equipment': equip_full['belong_equipment'] if equip_full else '',
                    'quantity': qty,
                    'unit': equip_full['unit'] if equip_full else '',
                    'unit_price': unit_price,
                    'total_price': total_price,
                    'remark': '',
                    'manufacturer': equip_full['manufacturer'] if equip_full else '',
                    'contract_no': equip_full['contract_no'] if equip_full else '',
                })

            detail = f"中队入库 {self.belong_unit}: " + ", ".join(
                f"{item['name']} {item['quantity']}{item['unit']}" for item in all_items
            )
            self.db.add_log('入库', '中队入库', None, detail,
                            operator=operator, role='unit', belong_unit=self.belong_unit)

            # 生成入库单
            dlg = ReceiptGroupDialog(self)
            if dlg.exec_() == QDialog.Accepted:
                self._generate_receipts(all_items, in_time, dlg.group_type)

            QMessageBox.information(self, "成功", f"中队入库 {len(all_items)} 条器材记录")
            self.accept()

        except Exception as e:
            QMessageBox.critical(self, "入库失败", f"保存入库记录时发生错误：{str(e)}")

    def _generate_receipts(self, all_items, in_date, group_type):
        """生成入库单（与机关入库相同逻辑）"""
        output_dir = os.path.join(os.path.dirname(__file__), '入库单')
        os.makedirs(output_dir, exist_ok=True)

        if group_type == 'factory':
            groups = {}
            for item in all_items:
                manufacturer = item.get('manufacturer', '未知厂家') or '未知厂家'
                contract_no = item.get('contract_no', '') or ''
                if manufacturer not in groups:
                    groups[manufacturer] = {'items': [], 'contract': contract_no}
                groups[manufacturer]['items'].append(item)
        else:
            contract_no = all_items[0].get('contract_no', '') if all_items else ''
            groups = {in_date: {'items': all_items, 'contract': contract_no}}

        total_files = 0
        for group_name, data in groups.items():
            try:
                files = ReceiptExporter.generate_in_receipt(
                    items=data['items'],
                    output_dir=output_dir,
                    group_name=group_name,
                    contract_no=data['contract'],
                    in_date=in_date,
                    stock_in_person='',
                    operator='',
                    group_type=group_type
                )
                total_files += len(files)
            except Exception as e:
                QMessageBox.critical(self, "生成入库单失败", str(e))
                return

        if total_files > 0:
            QMessageBox.information(self, "入库单已生成",
                                    f"已自动生成 {total_files} 个入库单文件\n保存位置：{output_dir}")


# ==================== 中队出库登记对话框 ====================
class UnitStockOutDialog(QDialog):
    """
    中队器材出库登记对话框
    与机关出库类似，显示当前中队库存，数量由 QDoubleSpinBox 输入，自动计算总价。
    保存时写入 unit_stock_record 表（record_type='out'），并进行库存不足检查。
    生成出库单 Excel 文件。
    """
    def __init__(self, db, belong_unit, operator_name='', parent=None):
        super().__init__(parent)
        self.db = db
        self.belong_unit = belong_unit          # 当前中队名称
        self.operator_name = operator_name
        self.setWindowTitle(f"中队出库登记 - {belong_unit}")
        self.setMinimumSize(850, 650)
        self.init_ui()

    def init_ui(self):
        """构建界面布局"""
        layout = QVBoxLayout(self)

        # ---------- 基本信息区域 ----------
        info_group = QGroupBox("出库基本信息")
        info_form = QFormLayout(info_group)

        # 出库去向（下拉列表 + 可自定义）
        self.out_direction = QComboBox()
        self.out_direction.addItems(OUT_DIRECTIONS_UNIT)
        self.out_direction.setEditable(True)
        info_form.addRow("出库去向:", self.out_direction)

        # 出库日期（默认今天）
        self.out_time = QDateEdit()
        self.out_time.setCalendarPopup(True)
        self.out_time.setDate(QDate.currentDate())
        info_form.addRow("出库日期:", self.out_time)

        # 出库人（自动填充）
        self.stock_out_person = QLineEdit()
        self.stock_out_person.setText(self.operator_name)
        info_form.addRow("出库人:", self.stock_out_person)

        # 经办人（自动填充）
        self.operator = QLineEdit()
        self.operator.setText(self.operator_name)
        info_form.addRow("经办人:", self.operator)

        layout.addWidget(info_group)

        # ---------- 器材明细区域 ----------
        detail_group = QGroupBox("器材明细")
        detail_layout = QVBoxLayout(detail_group)

        bl = QHBoxLayout()
        btn_add = QPushButton("＋ 添加一行器材")
        btn_add.clicked.connect(self.add_item)
        btn_del = QPushButton("－ 删除选中行")
        btn_del.clicked.connect(self.del_item)
        bl.addWidget(btn_add)
        bl.addWidget(btn_del)
        bl.addStretch()
        detail_layout.addLayout(bl)

        # 器材明细表格：7列（含库存）
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(
            ['器材编号', '器材名称', '规格型号', '库存', '出库数量', '单价', '总价'])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setMinimumHeight(250)
        detail_layout.addWidget(self.table)

        layout.addWidget(detail_group)

        # ---------- 确定/取消按钮 ----------
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.save_all)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

    def add_item(self):
        """添加一条器材明细，显示中队当前库存"""
        dlg = EquipmentSelectDialog(self.db, parent=self)
        if dlg.exec_() == QDialog.Accepted and dlg.selected_equipment:
            e = dlg.selected_equipment
            stock = self.db.get_unit_current_stock(e['id'], self.belong_unit)  # 获取中队库存
            row = self.table.rowCount()
            self.table.insertRow(row)

            item0 = QTableWidgetItem(e['serial_no'] or '')
            item0.setData(Qt.UserRole, e['id'])
            self.table.setItem(row, 0, item0)
            self.table.setItem(row, 1, QTableWidgetItem(e['name']))
            self.table.setItem(row, 2, QTableWidgetItem(e['spec_model'] or ''))

            # 库存（只读显示）
            self.table.setItem(row, 3, QTableWidgetItem(str(int(stock))))

            # 出库数量（数字输入框）
            qty_spin = QDoubleSpinBox()
            qty_spin.setRange(1, 999999)
            #qty_spin.setDecimals(2)
            qty_spin.setValue(1)
            qty_spin.valueChanged.connect(lambda v, r=row: self.calc_row(r))
            self.table.setCellWidget(row, 4, qty_spin)

            price_item = QTableWidgetItem(f"{e['unit_price']:.2f}")
            price_item.setData(Qt.UserRole, e['unit_price'] or 0)
            self.table.setItem(row, 5, price_item)
            self.table.setItem(row, 6, QTableWidgetItem(f"{e['unit_price']:.2f}"))

    def calc_row(self, row):
        """计算总价 = 数量 × 单价"""
        qty_widget = self.table.cellWidget(row, 4)
        if qty_widget and self.table.item(row, 5):
            qty = qty_widget.value()
            price = self.table.item(row, 5).data(Qt.UserRole) or 0
            total = qty * price
            self.table.item(row, 6).setText(f"{total:.2f}")

    def del_item(self):
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

    def save_all(self):
        """保存所有出库记录，并检查库存"""
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "提示", "请至少添加一条器材明细")
            return

        out_direction = self.out_direction.currentText().strip()
        out_time = self.out_time.date().toString('yyyy-MM-dd')
        stock_out_person = self.stock_out_person.text().strip()
        operator = self.operator.text().strip()
        batch_no = datetime.now().strftime('%Y%m%d%H%M%S') + '_' + uuid.uuid4().hex[:6]

        all_items = []
        try:
            for row in range(self.table.rowCount()):
                equip_id = self.table.item(row, 0).data(Qt.UserRole)
                qty = self.table.cellWidget(row, 4).value()
                stock = int(float(self.table.item(row, 3).text()))  # 库存文本转数字

                # 库存不足检查
                if qty > stock:
                    QMessageBox.warning(self, "库存不足",
                                        f"器材 {self.table.item(row, 1).text()} {self.belong_unit}库存仅 {stock}")
                    return

                unit_price = self.table.item(row, 5).data(Qt.UserRole)
                total_price = qty * unit_price

                data = {
                    'equipment_id': equip_id,
                    'belong_unit': self.belong_unit,
                    'record_type': 'out',
                    'quantity': qty,
                    'total_price': total_price,
                    'out_time': out_time,
                    'out_direction': out_direction,
                    'stock_out_person': stock_out_person,
                    'operator': operator,
                    'remark': '',
                    'batch_no': batch_no
                }
                self.db.add_unit_record(data)

                equip_full = self.db.get_equipment_by_id(equip_id)
                all_items.append({
                    'serial_no': equip_full['serial_no'] if equip_full else '',
                    'name': self.table.item(row, 1).text(),
                    'spec_model': self.table.item(row, 2).text(),
                    'belong_equipment': equip_full['belong_equipment'] if equip_full else '',
                    'quantity': qty,
                    'unit': equip_full['unit'] if equip_full else '',
                    'unit_price': unit_price,
                    'total_price': total_price,
                    'remark': '',
                    'manufacturer': equip_full['manufacturer'] if equip_full else '',
                    'contract_no': equip_full['contract_no'] if equip_full else '',
                })

            detail = f"中队出库 {self.belong_unit}: " + ", ".join(
                f"{item['name']} {item['quantity']}{item['unit']}" for item in all_items
            )
            self.db.add_log('出库', '中队出库', None, detail,
                            operator=operator, role='unit', belong_unit=self.belong_unit)

            # 生成出库单
            dlg = OutReceiptGroupDialog(self)
            if dlg.exec_() == QDialog.Accepted:
                self._generate_exporters(all_items, out_time, dlg.group_type)

            QMessageBox.information(self, "成功", f"中队出库 {len(all_items)} 条器材记录")
            self.accept()

        except Exception as e:
            QMessageBox.critical(self, "出库失败", f"保存出库记录时发生错误：{str(e)}")

    def _generate_exporters(self, all_items, out_date, group_type):
        """生成出库单（与机关出库相同逻辑）"""
        output_dir = os.path.join(os.path.dirname(__file__), '出库单')
        os.makedirs(output_dir, exist_ok=True)

        if group_type == 'equipment':
            groups = {}
            for item in all_items:
                eq = item.get('belong_equipment', '未知装备') or '未知装备'
                if eq not in groups:
                    groups[eq] = []
                groups[eq].append(item)
        else:
            groups = {out_date: all_items}

        total_files = 0
        for group_name, items in groups.items():
            first = items[0]
            try:
                files = ReceiptExporter.generate_out_receipt(
                    items=items,
                    output_dir=output_dir,
                    group_name=group_name,
                    contract_no=first.get('contract_no', ''),
                    out_date=out_date,
                    stock_out_person='',
                    operator='',
                    group_type=group_type
                )
                total_files += len(files)
            except Exception as e:
                QMessageBox.critical(self, "生成出库单失败", str(e))
                return

        if total_files > 0:
            QMessageBox.information(self, "出库单已生成",
                                    f"已自动生成 {total_files} 个出库单文件\n保存位置：{output_dir}")

# ==================== 主窗口 ====================
class MainWindow(QMainWindow):
    """
    应用程序主窗口
    根据登录角色（机关/中队）显示不同的选项卡界面，
    集成了器材目录、出入库管理、库存看板、操作日志、数据备份等功能。
    """
    def __init__(self, role, belong_unit, operator_name):
        super().__init__()
        self.db = Database()                    # 创建数据库连接
        self.role = role                        # 'org' 或 'unit'
        self.belong_unit = belong_unit          # 中队名称（机关为空字符串）
        self.operator_name = operator_name      # 操作人姓名

        # 状态栏显示当前用户信息
        self.status_bar = self.statusBar()
        role_text = '机关管理员' if self.role == 'org' else f'{self.belong_unit}管理员'
        self.status_bar.showMessage(f"当前用户: {self.operator_name} | 角色: {role_text}")

        # 窗口标题
        title = "装备维修器材出入库管理系统 - " + (
            '机关器材管理员' if role == 'org' else f'{belong_unit}管理员')
        self.setWindowTitle(title)
        self.setMinimumSize(1350, 820)

        self.init_ui()                          # 初始化界面
        self.check_expiry()                     # 检查即将到期的器材并提醒

        self.auto_backup()   # 启动时自动备份数据库

        # 记录登录日志
        self.db.add_log('登录', '系统', None, f'{title}登录',
                        operator=operator_name, role=role, belong_unit=belong_unit)



    def init_ui(self):
        """创建选项卡容器，并逐个添加选项卡"""
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # 器材目录选项卡（机关和中队共用）
        self.tab_equip = self.create_equipment_tab()
        self.tabs.addTab(self.tab_equip, "器材目录")

        # 根据角色添加不同的功能选项卡
        if self.role == 'org':
            self.tabs.addTab(self.create_org_in_tab(), "入库管理")
            self.tabs.addTab(self.create_org_out_tab(), "出库管理(发放中队)")
            self.tabs.addTab(self.create_org_dashboard_tab(), "📊 机关库存看板")
        else:
            self.tabs.addTab(self.create_unit_in_tab(), "中队入库")
            self.tabs.addTab(self.create_unit_out_tab(), "中队出库")
            self.tabs.addTab(self.create_unit_dashboard_tab(), "📊 中队库存看板")

        # 日志和备份选项卡（共用）
        self.tabs.addTab(self.create_log_tab(), "📝 操作日志")
        self.tabs.addTab(self.create_backup_tab(), "💾 数据备份")

    def auto_backup(self):
        """程序启动时自动备份数据库到 backups 目录"""
        backup_dir = os.path.join(os.path.dirname(__file__), 'backups')
        os.makedirs(backup_dir, exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_file = os.path.join(backup_dir, f"auto_backup_{timestamp}.bak")
        try:
            self.db.backup_database(backup_file)
            self.db.add_log('备份', '数据库', None, f'自动备份到:{backup_file}',
                            operator=self.operator_name, role=self.role,
                            belong_unit=self.belong_unit)
        except Exception as e:
            # 自动备份失败不影响主流程，仅控制台输出或记录日志
            print(f"自动备份失败：{e}")

    # ==================== 器材目录选项卡 ====================
    def create_equipment_tab(self):
        """
        创建器材目录选项卡
        显示所有器材信息，支持搜索筛选（关键词 + 质量等级）、
        新增、编辑、删除器材，并显示当前库存（机关或中队）。
        """
        w = QWidget()
        layout = QVBoxLayout(w)

        # ---------- 搜索筛选区域 ----------
        sg = QGroupBox("搜索筛选")
        sl = QHBoxLayout(sg)

        sl.addWidget(QLabel("关键词:"))
        self.eq_search = QLineEdit()
        self.eq_search.setPlaceholderText("所属单位/器材名称/规格型号/编号")
        self.eq_search.setMaximumWidth(220)
        sl.addWidget(self.eq_search)

        sl.addWidget(QLabel("质量:"))
        self.eq_quality = QComboBox()
        self.eq_quality.addItems(['全部'] + QUALITY_LEVELS)   # 质量等级下拉框
        sl.addWidget(self.eq_quality)

        btn_s = QPushButton("搜索")
        btn_s.clicked.connect(self.refresh_equipment_table)
        sl.addWidget(btn_s)

        btn_r = QPushButton("重置")
        btn_r.clicked.connect(lambda: (
            self.eq_search.clear(),
            self.eq_quality.setCurrentIndex(0),
            self.refresh_equipment_table()))
        sl.addWidget(btn_r)
        sl.addStretch()

        layout.addWidget(sg)

        # ---------- 操作按钮区域 ----------
        bl = QHBoxLayout()
        btn_add = QPushButton("＋ 新增器材")
        btn_add.clicked.connect(self.add_equipment)
        btn_edit = QPushButton("✎ 编辑")
        btn_edit.clicked.connect(self.edit_equipment)
        btn_del = QPushButton("✕ 删除")
        btn_del.clicked.connect(self.delete_equipment)
        bl.addWidget(btn_add)
        bl.addWidget(btn_edit)
        bl.addWidget(btn_del)
        bl.addStretch()
        layout.addLayout(bl)

        # ---------- 器材列表表格 ----------
        stock_label = '机关库存' if self.role == 'org' else f'{self.belong_unit}库存'
        cols = ['器材编号', '器材名称', '规格型号', '品种标识码', '计量单位',
                '质量等级', '单价', '生产日期', '存储寿命(月)', '来源类别',
                '所属单位', '所属装备', stock_label]

        self.eq_table = QTableWidget()
        self.eq_table.setColumnCount(len(cols))
        self.eq_table.setHorizontalHeaderLabels(cols)
        self.eq_table.setSelectionBehavior(QTableWidget.SelectRows)    # 整行选择
        self.eq_table.setEditTriggers(QTableWidget.NoEditTriggers)    # 不可编辑
        self.eq_table.setAlternatingRowColors(True)                   # 交替行颜色
        self.eq_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.eq_table, stretch=1)

        self.refresh_equipment_table()
        return w

    def refresh_equipment_table(self):
        """刷新器材列表，应用关键词和质量等级筛选"""
        kw = self.eq_search.text().strip()
        q = self.eq_quality.currentText()
        if q == '全部':
            q = ''

        records = self.db.search_equipment(keyword=kw, quality_level=q)
        self.eq_table.setRowCount(len(records))

        for i, r in enumerate(records):
            self.eq_table.setItem(i, 0, QTableWidgetItem(r['serial_no'] or ''))
            self.eq_table.setItem(i, 1, QTableWidgetItem(r['name'] or ''))
            self.eq_table.setItem(i, 2, QTableWidgetItem(r['spec_model'] or ''))
            self.eq_table.setItem(i, 3, QTableWidgetItem(r['category_code'] or ''))
            self.eq_table.setItem(i, 4, QTableWidgetItem(r['unit'] or ''))
            self.eq_table.setItem(i, 5, QTableWidgetItem(r['quality_level'] or ''))
            self.eq_table.setItem(i, 6, QTableWidgetItem(
                f"¥{r['unit_price']:.2f}" if r['unit_price'] else ''))
            self.eq_table.setItem(i, 7, QTableWidgetItem(r['production_date'] or ''))
            self.eq_table.setItem(i, 8, QTableWidgetItem(
                str(r['storage_life']) if r['storage_life'] else ''))
            self.eq_table.setItem(i, 9, QTableWidgetItem(r['source_type'] or ''))
            self.eq_table.setItem(i, 10, QTableWidgetItem(r['belong_unit'] or ''))
            self.eq_table.setItem(i, 11, QTableWidgetItem(r['belong_equipment'] or ''))

            # 根据角色显示不同库存
            if self.role == 'org':
                stock = self.db.get_org_current_stock(r['id'])
            else:
                stock = self.db.get_unit_current_stock(r['id'], self.belong_unit)

            si = QTableWidgetItem(str(stock))
            if stock <= 0:
                si.setForeground(Qt.red)        # 库存为0时显示红色
            self.eq_table.setItem(i, 12, si)

        self.status_bar.showMessage(f"器材目录：共 {len(records)} 条")

    def get_selected_equip_id(self):
        """获取当前选中的器材ID（根据名称和规格匹配）"""
        row = self.eq_table.currentRow()
        if row < 0:
            return None
        name = self.eq_table.item(row, 1).text()
        records = self.db.search_equipment(keyword=name)
        return records[0]['id'] if records else None

    def add_equipment(self):
        """打开新增器材对话框"""
        dlg = EquipmentDialog(self.db, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self.db.add_log('新增', '器材', None, '新增器材',
                            operator=self.operator_name, role=self.role,
                            belong_unit=self.belong_unit)
            self.refresh_equipment_table()

    def edit_equipment(self):
        """打开编辑器材对话框"""
        eid = self.get_selected_equip_id()
        if eid:
            dlg = EquipmentDialog(self.db, eid, parent=self)
            if dlg.exec_() == QDialog.Accepted:
                self.db.add_log('编辑', '器材', eid, '编辑器材',
                                operator=self.operator_name, role=self.role,
                                belong_unit=self.belong_unit)
                self.refresh_equipment_table()

    def delete_equipment(self):
        """删除选中的器材，同时安全删除关联的出入库记录"""
        eid = self.get_selected_equip_id()
        if eid is None:
            return
        # 获取器材信息用于确认
        equip = self.db.get_equipment_by_id(eid)
        if not equip:
            return
        name = equip['name']
        spec = equip['spec_model'] or ''
        msg = f"确定删除器材：{name} ({spec}) 吗？\n这将同时删除该器材的所有出入库记录！"
        if QMessageBox.question(self, "确认删除", msg,
                                QMessageBox.Yes | QMessageBox.No,
                                QMessageBox.No) != QMessageBox.Yes:
            return
        try:
            self.db.begin_transaction()
            # 删除关联的机关入库记录
            self.db.cursor.execute("DELETE FROM org_stock_in WHERE equipment_id=?", (eid,))
            # 删除关联的机关出库记录
            self.db.cursor.execute("DELETE FROM org_stock_out WHERE equipment_id=?", (eid,))
            # 删除关联的中队记录
            self.db.cursor.execute("DELETE FROM unit_stock_record WHERE equipment_id=?", (eid,))
            # 删除器材本身
            self.db.cursor.execute("DELETE FROM equipment WHERE id=?", (eid,))
            self.db.commit_transaction()
            self.db.add_log('删除', '器材', eid, f'删除器材 {name}',
                            operator=self.operator_name, role=self.role,
                            belong_unit=self.belong_unit)
            self.refresh_equipment_table()
        except Exception as e:
            self.db.rollback_transaction()
            QMessageBox.critical(self, "删除失败", f"无法删除器材：{str(e)}")

    # ==================== 机关入库记录选项卡 ====================
    def create_org_in_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)

        # ---------- 搜索筛选区域 ----------
        sg = QGroupBox("搜索筛选")
        sl = QHBoxLayout(sg)
        sl.addWidget(QLabel("关键词:"))
        self.org_in_search = QLineEdit()
        self.org_in_search.setPlaceholderText("器材名称/规格型号/编号")
        sl.addWidget(self.org_in_search)
        sl.addWidget(QLabel("质量:"))
        self.org_in_quality = QComboBox()
        self.org_in_quality.addItems(['全部'] + QUALITY_LEVELS)
        sl.addWidget(self.org_in_quality)
        btn_search = QPushButton("搜索")
        btn_search.clicked.connect(self.refresh_org_in_table)
        sl.addWidget(btn_search)
        btn_reset = QPushButton("重置")
        btn_reset.clicked.connect(lambda: (
            self.org_in_search.clear(),
            self.org_in_quality.setCurrentIndex(0),
            self.refresh_org_in_table()
        ))
        sl.addWidget(btn_reset)
        sl.addStretch()
        layout.addWidget(sg)

        # ---------- 操作按钮行 ----------
        bl = QHBoxLayout()
        btn_add = QPushButton("＋ 新增入库")
        btn_add.clicked.connect(self.add_org_in)
        bl.addWidget(btn_add)
        btn_import = QPushButton("📥 从Excel导入")
        btn_import.clicked.connect(self.import_org_in_excel)
        bl.addWidget(btn_import)
        btn_export = QPushButton("📤 导出Excel")
        btn_export.clicked.connect(self.export_org_in_excel)
        bl.addWidget(btn_export)
        btn_regen = QPushButton("📄 重新生成本日入库单")
        btn_regen.clicked.connect(self.regen_today_org_in)
        bl.addWidget(btn_regen)
        btn_delete = QPushButton("✕ 删除选中记录")
        btn_delete.clicked.connect(self.delete_org_in)
        bl.addWidget(btn_delete)
        bl.addStretch()
        btn_ref = QPushButton("↻ 刷新")
        btn_ref.clicked.connect(self.refresh_org_in_table)
        bl.addWidget(btn_ref)
        layout.addLayout(bl)

        # ---------- 入库记录表格 ----------
        self.org_in_table = QTableWidget()
        self.org_in_table.setColumnCount(13)
        self.org_in_table.setHorizontalHeaderLabels([
            'ID', '器材编号', '器材名称', '规格型号', '入库数量', '计量单位', '总价',
            '入库时间', '来源类别', '入库人', '经办人', '合同编号', '备注'
        ])
        self.org_in_table.setColumnHidden(0, True)
        self.org_in_table.horizontalHeader().setStretchLastSection(True)
        self.org_in_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.org_in_table.setAlternatingRowColors(True)
        self.org_in_table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.org_in_table, stretch=1)

        self.refresh_org_in_table()
        return w

    def refresh_org_in_table(self):
        """刷新机关入库记录表格，支持搜索筛选"""
        records = self.db.get_all_org_stock_in()
        # 获取筛选条件
        keyword = self.org_in_search.text().strip()
        q = self.org_in_quality.currentText()
        if q == '全部':
            q = ''
        # 过滤
        if keyword or q:
            filtered = []
            for r in records:
                if keyword:
                    kw = keyword.lower()
                    match = (kw in (r['name'] or '').lower() or
                             kw in (r['spec_model'] or '').lower() or
                             kw in (r['serial_no'] or '').lower())
                    if not match:
                        continue
                if q and (r['quality_level'] or '') != q:
                    continue
                filtered.append(r)
            records = filtered

        self.org_in_table.setRowCount(len(records))
        for i, r in enumerate(records):
            self.org_in_table.setItem(i, 0, QTableWidgetItem(str(r['id'])))
            self.org_in_table.setItem(i, 1, QTableWidgetItem(r['serial_no'] or ''))
            self.org_in_table.setItem(i, 2, QTableWidgetItem(r['name'] or ''))
            self.org_in_table.setItem(i, 3, QTableWidgetItem(r['spec_model'] or ''))
            self.org_in_table.setItem(i, 4, QTableWidgetItem(str(int(r['quantity']))))
            self.org_in_table.setItem(i, 5, QTableWidgetItem(r['unit'] or ''))
            self.org_in_table.setItem(i, 6, QTableWidgetItem(
                f"¥{r['total_price']:.2f}" if r['total_price'] else ''))
            self.org_in_table.setItem(i, 7, QTableWidgetItem(r['in_time'] or ''))
            self.org_in_table.setItem(i, 8, QTableWidgetItem(r['source_type'] or ''))
            self.org_in_table.setItem(i, 9, QTableWidgetItem(r['stock_in_person'] or ''))
            self.org_in_table.setItem(i, 10, QTableWidgetItem(r['operator'] or ''))
            self.org_in_table.setItem(i, 11, QTableWidgetItem(r['contract_no'] or ''))
            self.org_in_table.setItem(i, 12, QTableWidgetItem(r['remark'] or ''))

    def add_org_in(self):
        """打开机关入库登记对话框，成功后刷新表格和器材目录"""
        dlg = OrgStockInDialog(self.db, self.operator_name, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self.db.add_log('入库', '机关入库', None, '机关入库操作',
                            operator=self.operator_name, role='org')
            self.refresh_org_in_table()
            self.refresh_equipment_table()

    def delete_org_in(self):
        """删除选中的入库记录"""
        row = self.org_in_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "提示", "请先选中一条入库记录")
            return
        record_id = int(self.org_in_table.item(row, 0).text())  # 隐藏的ID列
        equip_name = self.org_in_table.item(row, 2).text()
        qty = self.org_in_table.item(row, 4).text()
        unit = self.org_in_table.item(row, 5).text()

        ret = QMessageBox.question(self, "确认删除",
                                   f"确定要删除 {equip_name} 的入库记录（数量 {qty}{unit}）吗？\n删除后库存会相应减少。",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ret == QMessageBox.Yes:
            try:
                # 删除 org_stock_in 表中对应记录
                self.db.cursor.execute("DELETE FROM org_stock_in WHERE id=?", (record_id,))
                self.db.conn.commit()
                self.db.add_log('删除', '机关入库', record_id,
                                f"删除入库记录: {equip_name} {qty}{unit}",
                                operator=self.operator_name, role='org')
                self.refresh_org_in_table()
                self.refresh_equipment_table()
            except Exception as e:
                QMessageBox.critical(self, "删除失败", str(e))

    def regen_today_org_in(self):
        """
        重新生成本日所有入库记录的入库单
        读取今天入库的数据，构造 all_items 列表，弹出分组选择后重新生成单据。
        """
        from datetime import date
        today_str = date.today().strftime('%Y-%m-%d')
        records = self.db.get_org_stock_in_by_date(today_str)
        if not records:
            QMessageBox.information(self, "提示", f"今天（{today_str}）没有入库记录")
            return

        # 构造生成单据所需的 items 列表
        all_items = []
        for r in records:
            equip = self.db.get_equipment_by_id(r['equipment_id'])
            all_items.append({
                'serial_no': r['serial_no'],
                'name': r['name'],
                'spec_model': r['spec_model'],
                'belong_equipment': equip['belong_equipment'] if equip else '',
                'quantity': int(r['quantity']),
                'unit': r['unit'],
                'unit_price': r['unit_price'],
                'total_price': r['total_price'],
                'remark': r['remark'],
                'manufacturer': r['manufacturer'],
                'contract_no': r['contract_no'],
            })

        dlg = ReceiptGroupDialog(self)
        if dlg.exec_() == QDialog.Accepted:
            output_dir = os.path.join(os.path.dirname(__file__), '入库单')
            os.makedirs(output_dir, exist_ok=True)
            try:
                if dlg.group_type == 'factory':
                    groups = {}
                    for item in all_items:
                        mf = item.get('manufacturer', '未知厂家') or '未知厂家'
                        contract_no = item.get('contract_no', '') or ''
                        if mf not in groups:
                            groups[mf] = {'items': [], 'contract': contract_no}
                        groups[mf]['items'].append(item)
                else:
                    contract_no = all_items[0].get('contract_no', '') if all_items else ''
                    groups = {today_str: {'items': all_items, 'contract': contract_no}}

                total_files = 0
                for group_name, data in groups.items():
                    files = ReceiptExporter.generate_in_receipt(
                        items=data['items'],
                        output_dir=output_dir,
                        group_name=group_name,
                        contract_no=data['contract'],
                        in_date=today_str,
                        stock_in_person='',
                        operator='',
                        group_type=dlg.group_type
                    )
                    total_files += len(files)

                QMessageBox.information(self, "成功",
                                        f"已重新生成 {total_files} 个入库单文件\n保存位置：{output_dir}")
            except Exception as e:
                QMessageBox.critical(self, "失败", str(e))

    def import_org_in_excel(self):
        """从机关统计表Excel导入入库记录，自动跳过空行和无效行"""
        from openpyxl import load_workbook
        file_path, _ = QFileDialog.getOpenFileName(self, "选择机关统计表", "",
                                                   "Excel文件(*.xlsx *.xls)")
        if not file_path:
            return
        try:
            wb = load_workbook(file_path)
            ws = wb.active
            # 从第2行开始读取
            rows = list(ws.iter_rows(min_row=2, values_only=True))
            if not rows:
                QMessageBox.information(self, "提示", "Excel中没有数据")
                return

            self.db.begin_transaction()
            batch_no = datetime.now().strftime('%Y%m%d%H%M%S') + '_import'
            count = 0
            detail_parts = []

            def safe_float(val, default=0.0):
                try:
                    return float(val)
                except:
                    return default

            def safe_str(val, default=''):
                try:
                    return str(val).strip()
                except:
                    return default

            for row in rows:
                # 跳过完全空白的行
                if not row or all(cell is None or (isinstance(cell, str) and cell.strip() == '') for cell in row):
                    continue

                # 提取关键字段
                name = safe_str(row[3])
                spec = safe_str(row[4])

                # 跳过无效行：名称为空、为"None"字符串、或包含标题关键词
                if not name or not spec or name.lower() == 'none' or spec.lower() == 'none':
                    continue
                if '器材名称' in name or '序号' in safe_str(row[0]):
                    continue

                serial_no = safe_str(row[5])
                qty = int(safe_float(row[6], 1.0))
                unit = safe_str(row[7])
                price = safe_float(row[8], 0.0)
                quality = safe_str(row[10], '新品')
                manufacturer = safe_str(row[11])
                category_code = safe_str(row[13])
                prod_date = safe_str(row[14])
                life = row[15] if row[15] else None
                source_type = safe_str(row[16])
                contract = safe_str(row[17])
                in_time = safe_str(row[18], datetime.now().strftime('%Y-%m-%d'))
                in_person = safe_str(row[19])
                operator = safe_str(row[20], self.operator_name)
                remark = safe_str(row[21]) if len(row) > 21 else ''
                belong_unit = safe_str(row[1])
                belong_equip = safe_str(row[2])

                # 查找或创建器材
                self.db.cursor.execute(
                    "SELECT id FROM equipment WHERE name=? AND spec_model=?", (name, spec))
                equip = self.db.cursor.fetchone()
                if not equip:
                    equip_id = self.db.add_equipment({
                        'name': name,
                        'spec_model': spec,
                        'serial_no': serial_no,
                        'category_code': category_code,
                        'unit': unit,
                        'quality_level': quality,
                        'manufacturer': manufacturer,
                        'unit_price': price,
                        'production_date': prod_date,
                        'storage_life': life,
                        'source_type': source_type,
                        'pricing_method': '',
                        'contract_no': contract,
                        'belong_unit': belong_unit,
                        'belong_equipment': belong_equip
                    })
                else:
                    equip_id = equip['id']

                total_price = qty * price
                self.db.cursor.execute(
                    "INSERT INTO org_stock_in (equipment_id, quantity, total_price, in_time, stock_in_person, operator, remark, batch_no) VALUES (?,?,?,?,?,?,?,?)",
                    (equip_id, qty, total_price, in_time, in_person, operator, remark, batch_no)
                )
                count += 1
                detail_parts.append(f"{name} {qty}{unit}")

            self.db.commit_transaction()
            if count == 0:
                QMessageBox.information(self, "提示", "没有导入任何有效记录，请检查Excel内容")
                return
            detail = "机关Excel导入: " + ", ".join(detail_parts)
            self.db.add_log('入库', '机关入库', None, detail,
                            operator=self.operator_name, role='org')
            QMessageBox.information(self, "导入成功", f"成功导入 {count} 条机关入库记录")
            self.refresh_org_in_table()
            self.refresh_equipment_table()

        except Exception as e:
            self.db.rollback_transaction()
            QMessageBox.critical(self, "导入失败", f"处理Excel时出错：{str(e)}")

    def import_unit_excel(self, record_type='in'):
        """从中队统计表Excel导入记录（入库/出库），自动跳过空行和无效行"""
        from openpyxl import load_workbook
        title = "中队入库统计表" if record_type == 'in' else "中队出库统计表"
        file_path, _ = QFileDialog.getOpenFileName(self, f"选择{title}", "",
                                                   "Excel文件(*.xlsx *.xls)")
        if not file_path:
            return
        try:
            wb = load_workbook(file_path)
            ws = wb.active
            rows = list(ws.iter_rows(min_row=2, values_only=True))
            if not rows:
                QMessageBox.information(self, "提示", "Excel中没有数据")
                return

            self.db.begin_transaction()
            batch_no = datetime.now().strftime('%Y%m%d%H%M%S') + '_import'
            count = 0
            detail_parts = []

            def safe_float(val, default=0.0):
                try:
                    return float(val)
                except:
                    return default

            def safe_str(val, default=''):
                try:
                    return str(val).strip()
                except:
                    return default

            for row in rows:
                # 跳过完全空白的行
                if not row or all(cell is None or (isinstance(cell, str) and cell.strip() == '') for cell in row):
                    continue

                name = safe_str(row[3])
                spec = safe_str(row[4])

                # 跳过无效行
                if not name or not spec or name.lower() == 'none' or spec.lower() == 'none':
                    continue
                if '器材名称' in name or '序号' in safe_str(row[0]):
                    continue

                serial_no = safe_str(row[5])
                qty = int(safe_float(row[6], 1.0))
                unit = safe_str(row[7])
                price = safe_float(row[8], 0.0)
                quality = safe_str(row[10], '新品')
                manufacturer = safe_str(row[11])
                category_code = safe_str(row[13])
                prod_date = safe_str(row[14])
                life = row[15] if row[15] else None
                source_type = safe_str(row[16])
                contract = safe_str(row[17])
                in_time = safe_str(row[18])
                out_time = safe_str(row[19])
                out_dir = safe_str(row[20])
                out_person = safe_str(row[21])
                operator = safe_str(row[22], self.operator_name)
                remark = safe_str(row[23]) if len(row) > 23 else ''
                belong_unit = safe_str(row[1], self.belong_unit)
                belong_equip = safe_str(row[2])

                if record_type == 'in' and not in_time:
                    continue
                if record_type == 'out' and not out_time:
                    continue

                # 查找或创建器材
                self.db.cursor.execute(
                    "SELECT id FROM equipment WHERE name=? AND spec_model=?", (name, spec))
                equip = self.db.cursor.fetchone()
                if not equip:
                    equip_id = self.db.add_equipment({
                        'name': name, 'spec_model': spec,
                        'serial_no': serial_no, 'category_code': category_code,
                        'unit': unit, 'quality_level': quality,
                        'manufacturer': manufacturer, 'unit_price': price,
                        'production_date': prod_date, 'storage_life': life,
                        'source_type': source_type, 'pricing_method': '',
                        'contract_no': contract, 'belong_unit': belong_unit,
                        'belong_equipment': belong_equip
                    })
                else:
                    equip_id = equip['id']

                total_price = qty * price
                if record_type == 'in':
                    self.db.add_unit_record({
                        'equipment_id': equip_id, 'belong_unit': belong_unit,
                        'record_type': 'in', 'quantity': qty,
                        'total_price': total_price, 'in_time': in_time,
                        'source': source_type, 'stock_in_person': '',
                        'operator': operator, 'remark': remark, 'batch_no': batch_no
                    })
                else:
                    self.db.add_unit_record({
                        'equipment_id': equip_id, 'belong_unit': belong_unit,
                        'record_type': 'out', 'quantity': qty,
                        'total_price': total_price, 'out_time': out_time,
                        'out_direction': out_dir, 'stock_out_person': out_person,
                        'operator': operator, 'remark': remark, 'batch_no': batch_no
                    })

                count += 1
                detail_parts.append(f"{name} {qty}{unit}")

            self.db.commit_transaction()
            if count == 0:
                QMessageBox.information(self, "提示", "没有导入任何有效记录，请检查Excel内容")
                return
            type_str = "入库" if record_type == 'in' else "出库"
            detail = f"中队Excel导入{type_str}: " + ", ".join(detail_parts)
            self.db.add_log(type_str, f'中队{type_str}', None, detail,
                            operator=self.operator_name, role='unit', belong_unit=self.belong_unit)
            QMessageBox.information(self, "导入成功", f"成功导入 {count} 条中队{type_str}记录")
            if record_type == 'in':
                self.refresh_unit_in_table()
            else:
                self.refresh_unit_out_table()
            self.refresh_equipment_table()

        except Exception as e:
            self.db.rollback_transaction()
            QMessageBox.critical(self, "导入失败", f"处理Excel时出错：{str(e)}")

    # ==================== 导出Excel（标准统计表格式） ====================
    def _write_stat_header(self, ws, num_cols, title_text):
        """写入第一行标题（合并、加粗、居中、行高约两行）"""
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
        cell = ws.cell(row=1, column=1, value=title_text)
        cell.font = Font(name='宋体', size=14, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 30  # 大约两行高度

    def export_org_in_excel(self):
        """导出机关入库记录为机关统计表（22列）"""
        records = self.db.get_all_org_stock_in()
        if not records:
            QMessageBox.information(self, "提示", "没有机关入库记录可导出")
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "导出机关入库统计表", "",
                                                   "Excel文件(*.xlsx)")
        if not file_path:
            return
        try:
            wb = Workbook()
            ws = wb.active
            # 第一行：标题
            self._write_stat_header(ws, 22, "装备维修器材入库情况统计表")
            # 第二行：表头
            headers = [
                '序号', '所属单位', '所属装备', '器材名称', '规格型号',
                '器材编号', '数量', '计量单位', '单价（元）', '总价（元）',
                '质量等级', '生产厂家', '计价方法', '品种标识码', '生产日期',
                '存储寿命', '来源类别', '合同编号', '入库时间', '入库人',
                '经办人', '备注'
            ]
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col, value=h)
                cell.font = Font(name='宋体', size=10, bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[2].height = 20

            # 数据行（从第3行开始）
            for i, r in enumerate(records, start=3):
                row_data = [
                    i - 2,  # 序号
                    r['belong_unit'] or '',  # 所属单位
                    r['belong_equipment'] or '',  # 所属装备
                    r['name'] or '',  # 器材名称
                    r['spec_model'] or '',  # 规格型号
                    r['serial_no'] or '',  # 器材编号
                    int(r['quantity']) or 0,  # 数量
                    r['unit'] or '',  # 计量单位
                    r['unit_price'] or 0,  # 单价
                    r['total_price'] or 0,  # 总价
                    r['quality_level'] or '',  # 质量等级
                    r['manufacturer'] or '',  # 生产厂家
                    r['pricing_method'] or '',  # 计价方法
                    r['category_code'] or '',  # 品种标识码
                    r['production_date'] or '',  # 生产日期
                    r['storage_life'] or '',  # 存储寿命
                    r['source_type'] or '',  # 来源类别
                    r['contract_no'] or '',  # 合同编号
                    r['in_time'] or '',  # 入库时间
                    r['stock_in_person'] or '',  # 入库人
                    r['operator'] or '',  # 经办人
                    r['remark'] or ''  # 备注
                ]
                for col, val in enumerate(row_data, 1):
                    ws.cell(row=i, column=col, value=val)
            wb.save(file_path)
            QMessageBox.information(self, "导出成功", f"已导出 {len(records)} 条机关入库记录")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))

    def export_org_out_excel(self):
        """导出机关出库记录为机关统计表（22列，出库）"""
        records = self.db.get_all_org_stock_out()
        if not records:
            QMessageBox.information(self, "提示", "没有机关出库记录可导出")
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "导出机关出库统计表", "",
                                                   "Excel文件(*.xlsx)")
        if not file_path:
            return
        try:
            wb = Workbook()
            ws = wb.active
            self._write_stat_header(ws, 22, "机关装备出入库统计表（出库）")
            headers = [
                '序号', '所属单位', '所属装备', '器材名称', '规格型号',
                '器材编号', '数量', '计量单位', '单价（元）', '总价（元）',
                '质量等级', '生产厂家', '计价方法', '品种标识码', '生产日期',
                '存储寿命', '来源类别', '合同编号', '出库时间', '出库人',
                '经办人', '备注'
            ]
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col, value=h)
                cell.font = Font(name='宋体', size=10, bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[2].height = 20

            for i, r in enumerate(records, start=3):
                row_data = [
                    i - 2,
                    r['to_unit'] or '',  # 所属单位（接收中队）
                    r['belong_equipment'] or '',
                    r['name'] or '',
                    r['spec_model'] or '',
                    r['serial_no'] or '',
                    int(r['quantity']) or 0,
                    r['unit'] or '',
                    r['unit_price'] or 0,
                    r['total_price'] or 0,
                    r['quality_level'] or '',
                    r['manufacturer'] or '',
                    r['pricing_method'] or '',
                    r['category_code'] or '',
                    r['production_date'] or '',
                    r['storage_life'] or '',
                    r['out_direction'] or '',  # 来源类别用出库去向替代
                    r['contract_no'] or '',
                    r['out_time'] or '',
                    r['stock_out_person'] or '',
                    r['operator'] or '',
                    r['remark'] or ''
                ]
                for col, val in enumerate(row_data, 1):
                    ws.cell(row=i, column=col, value=val)
            wb.save(file_path)
            QMessageBox.information(self, "导出成功", f"已导出 {len(records)} 条机关出库记录")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))

    def export_unit_in_excel(self):
        """导出中队入库记录为中队统计表（24列）"""
        records = self.db.get_unit_records(self.belong_unit, 'in')
        if not records:
            QMessageBox.information(self, "提示", "没有中队入库记录可导出")
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "导出中队入库统计表", "",
                                                   "Excel文件(*.xlsx)")
        if not file_path:
            return
        try:
            wb = Workbook()
            ws = wb.active
            self._write_stat_header(ws, 24, "装备维修器材入库情况统计表（中队）")
            headers = [
                '序号', '所属单位', '所属装备', '器材名称', '规格型号',
                '器材编号', '数量', '计量单位', '单价（元）', '总价（元）',
                '质量等级', '生产厂家', '计价方法', '品种标识码', '生产日期',
                '存储寿命', '来源类别', '合同编号', '入库时间', '出库时间',
                '出库去向', '出库人', '经办人', '备注'
            ]
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col, value=h)
                cell.font = Font(name='宋体', size=10, bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[2].height = 20

            for i, r in enumerate(records, start=3):
                row_data = [
                    i - 2,
                    self.belong_unit,  # 所属单位
                    r['belong_equipment'] or '',
                    r['name'] or '',
                    r['spec_model'] or '',
                    r['serial_no'] or '',
                    int(r['quantity']) or 0,
                    r['unit'] or '',
                    r['unit_price'] or 0,
                    r['total_price'] or 0,
                    r['quality_level'] or '',
                    r['manufacturer'] or '',
                    r['pricing_method'] or '',  # 来自 equipment 表
                    r['category_code'] or '',
                    r['production_date'] or '',
                    r['storage_life'] or '',
                    r['source'] or '',  # 来源
                    r['contract_no'] or '',
                    r['in_time'] or '',
                    '',  # 出库时间（入库为空）
                    '',  # 出库去向
                    '',  # 出库人
                    r['operator'] or '',
                    r['remark'] or ''
                ]
                for col, val in enumerate(row_data, 1):
                    ws.cell(row=i, column=col, value=val)
            wb.save(file_path)
            QMessageBox.information(self, "导出成功", f"已导出 {len(records)} 条中队入库记录")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))

    def export_unit_out_excel(self):
        """导出中队出库记录为中队统计表（24列）"""
        records = self.db.get_unit_records(self.belong_unit, 'out')
        if not records:
            QMessageBox.information(self, "提示", "没有中队出库记录可导出")
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "导出中队出库统计表", "",
                                                   "Excel文件(*.xlsx)")
        if not file_path:
            return
        try:
            wb = Workbook()
            ws = wb.active
            self._write_stat_header(ws, 24, "装备维修器材出库情况统计表（中队）")
            headers = [
                '序号', '所属单位', '所属装备', '器材名称', '规格型号',
                '器材编号', '数量', '计量单位', '单价（元）', '总价（元）',
                '质量等级', '生产厂家', '计价方法', '品种标识码', '生产日期',
                '存储寿命', '来源类别', '合同编号', '入库时间', '出库时间',
                '出库去向', '出库人', '经办人', '备注'
            ]
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col, value=h)
                cell.font = Font(name='宋体', size=10, bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[2].height = 20

            for i, r in enumerate(records, start=3):
                row_data = [
                    i - 2,
                    self.belong_unit,
                    r['belong_equipment'] or '',
                    r['name'] or '',
                    r['spec_model'] or '',
                    r['serial_no'] or '',
                    int(r['quantity']) or 0,
                    r['unit'] or '',
                    r['unit_price'] or 0,
                    r['total_price'] or 0,
                    r['quality_level'] or '',
                    r['manufacturer'] or '',
                    r['pricing_method'] or '',
                    r['category_code'] or '',
                    r['production_date'] or '',
                    r['storage_life'] or '',
                    '',  # 来源类别（出库无）
                    r['contract_no'] or '',
                    '',  # 入库时间
                    r['out_time'] or '',
                    r['out_direction'] or '',
                    r['stock_out_person'] or '',
                    r['operator'] or '',
                    r['remark'] or ''
                ]
                for col, val in enumerate(row_data, 1):
                    ws.cell(row=i, column=col, value=val)
            wb.save(file_path)
            QMessageBox.information(self, "导出成功", f"已导出 {len(records)} 条中队出库记录")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))

    def export_org_stock_excel(self):
        """导出机关当前库存统计表（22列，仅正库存器材）"""
        records = self.db.search_equipment()  # 获取所有器材
        items = []
        for r in records:
            stock = self.db.get_org_current_stock(r['id'])
            if stock > 0:
                r_dict = dict(r)  # sqlite3.Row 转字典
                r_dict['current_stock'] = stock
                items.append(r_dict)
        if not items:
            QMessageBox.information(self, "提示", "当前没有库存器材可导出")
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "导出机关库存统计表", "",
                                                   "Excel文件(*.xlsx)")
        if not file_path:
            return
        try:
            wb = Workbook()
            ws = wb.active
            self._write_stat_header(ws, 22, "机关库存统计表")
            headers = [
                '序号', '所属单位', '所属装备', '器材名称', '规格型号',
                '器材编号', '数量（库存）', '计量单位', '单价（元）', '总价（元）',
                '质量等级', '生产厂家', '计价方法', '品种标识码', '生产日期',
                '存储寿命', '来源类别', '合同编号', '', '', '', '备注'
            ]
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col, value=h)
                cell.font = Font(name='宋体', size=10, bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[2].height = 20

            for i, r in enumerate(items, start=3):
                row_data = [
                    i - 2,
                    r['belong_unit'] or '',
                    r['belong_equipment'] or '',
                    r['name'] or '',
                    r['spec_model'] or '',
                    r['serial_no'] or '',
                    r['current_stock'],  # 库存数量
                    r['unit'] or '',
                    r['unit_price'] or 0,
                    r['unit_price'] * r['current_stock'] if r['unit_price'] else 0,  # 总价=单价×库存
                    r['quality_level'] or '',
                    r['manufacturer'] or '',
                    r['pricing_method'] or '',
                    r['category_code'] or '',
                    r['production_date'] or '',
                    r['storage_life'] or '',
                    r['source_type'] or '',
                    r['contract_no'] or '',
                    '', '', '',  # 入库时间、入库人、经办人留空
                    ''  # 备注
                ]
                for col, val in enumerate(row_data, 1):
                    ws.cell(row=i, column=col, value=val)
            wb.save(file_path)
            QMessageBox.information(self, "导出成功", f"已导出 {len(items)} 条库存器材")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))

    def export_unit_stock_excel(self):
        """导出中队当前库存统计表（24列，仅正库存器材）"""
        records = self.db.search_equipment()
        items = []
        for r in records:
            stock = self.db.get_unit_current_stock(r['id'], self.belong_unit)
            if stock > 0:
                r_dict = dict(r)
                r_dict['current_stock'] = stock
                items.append(r_dict)
        if not items:
            QMessageBox.information(self, "提示", "当前没有库存器材可导出")
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "导出中队库存统计表", "",
                                                   "Excel文件(*.xlsx)")
        if not file_path:
            return
        try:
            wb = Workbook()
            ws = wb.active
            self._write_stat_header(ws, 24, f"{self.belong_unit} 库存统计表")
            headers = [
                '序号', '所属单位', '所属装备', '器材名称', '规格型号',
                '器材编号', '数量（库存）', '计量单位', '单价（元）', '总价（元）',
                '质量等级', '生产厂家', '计价方法', '品种标识码', '生产日期',
                '存储寿命', '来源类别', '合同编号', '', '', '', '', '经办人', '备注'
            ]
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col, value=h)
                cell.font = Font(name='宋体', size=10, bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[2].height = 20

            for i, r in enumerate(items, start=3):
                row_data = [
                    i - 2,
                    self.belong_unit,
                    r['belong_equipment'] or '',
                    r['name'] or '',
                    r['spec_model'] or '',
                    r['serial_no'] or '',
                    r['current_stock'],
                    r['unit'] or '',
                    r['unit_price'] or 0,
                    r['unit_price'] * r['current_stock'] if r['unit_price'] else 0,
                    r['quality_level'] or '',
                    r['manufacturer'] or '',
                    r['pricing_method'] or '',
                    r['category_code'] or '',
                    r['production_date'] or '',
                    r['storage_life'] or '',
                    r['source_type'] or '',
                    r['contract_no'] or '',
                    '', '', '',  # 入库时间、出库时间、出库去向留空
                    '',  # 经办人留空
                    ''
                ]
                for col, val in enumerate(row_data, 1):
                    ws.cell(row=i, column=col, value=val)
            wb.save(file_path)
            QMessageBox.information(self, "导出成功", f"已导出 {len(items)} 条库存器材")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))

    # ==================== 机关出库记录选项卡 ====================
    def create_org_out_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)

        bl = QHBoxLayout()
        btn_add = QPushButton("＋ 新增出库(发放中队)")
        btn_add.clicked.connect(self.add_org_out)
        bl.addWidget(btn_add)

        btn_export = QPushButton("📤 导出Excel")  # 新增
        btn_export.clicked.connect(self.export_org_out_excel)
        bl.addWidget(btn_export)

        btn_delete = QPushButton("✕ 删除选中记录")
        btn_delete.clicked.connect(self.delete_org_out)
        bl.addWidget(btn_delete)

        btn_regen = QPushButton("📄 重新生成本日出库单")
        btn_regen.clicked.connect(self.regen_today_org_out)
        bl.addWidget(btn_regen)

        bl.addStretch()
        btn_ref = QPushButton("↻ 刷新")
        btn_ref.clicked.connect(self.refresh_org_out_table)
        bl.addWidget(btn_ref)
        layout.addLayout(bl)

        self.org_out_table = QTableWidget()
        self.org_out_table.setColumnCount(13)
        self.org_out_table.setHorizontalHeaderLabels([
            'ID', '器材编号', '器材名称', '规格型号', '发放中队', '出库数量', '总价',
            '出库时间', '出库去向', '出库人', '经办人', '合同编号', '备注'
        ])
        self.org_out_table.setColumnHidden(0, True)
        self.org_out_table.horizontalHeader().setStretchLastSection(True)
        self.org_out_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.org_out_table.setAlternatingRowColors(True)
        self.org_out_table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.org_out_table, stretch=1)

        self.refresh_org_out_table()
        return w


    def refresh_org_out_table(self):
        """刷新机关出库记录表格"""
        records = self.db.get_all_org_stock_out()
        self.org_out_table.setRowCount(len(records))
        for i, r in enumerate(records):
            self.org_out_table.setItem(i, 0, QTableWidgetItem(str(r['id'])))
            self.org_out_table.setItem(i, 1, QTableWidgetItem(r['serial_no'] or ''))
            self.org_out_table.setItem(i, 2, QTableWidgetItem(r['name'] or ''))
            self.org_out_table.setItem(i, 3, QTableWidgetItem(r['spec_model'] or ''))
            self.org_out_table.setItem(i, 4, QTableWidgetItem(r['to_unit'] or ''))
            self.org_out_table.setItem(i, 5, QTableWidgetItem(str(int(r['quantity']))))
            self.org_out_table.setItem(i, 6, QTableWidgetItem(f"¥{r['total_price']:.2f}" if r['total_price'] else ''))
            self.org_out_table.setItem(i, 7, QTableWidgetItem(r['out_time'] or ''))
            self.org_out_table.setItem(i, 8, QTableWidgetItem(r['out_direction'] or ''))
            self.org_out_table.setItem(i, 9, QTableWidgetItem(r['stock_out_person'] or ''))
            self.org_out_table.setItem(i, 10, QTableWidgetItem(r['operator'] or ''))
            self.org_out_table.setItem(i, 11, QTableWidgetItem(r['contract_no'] or ''))
            self.org_out_table.setItem(i, 12, QTableWidgetItem(r['remark'] or ''))

    def add_org_out(self):
        """打开机关出库登记对话框"""
        dlg = OrgStockOutDialog(self.db, self.operator_name, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self.db.add_log('出库', '机关出库', None, '机关发放中队操作',
                            operator=self.operator_name, role='org')
            self.refresh_org_out_table()
            self.refresh_equipment_table()

    def delete_org_out(self):
        """删除选中的出库记录"""
        row = self.org_out_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "提示", "请先选中一条出库记录")
            return
        record_id = int(self.org_out_table.item(row, 0).text())
        equip_name = self.org_out_table.item(row, 2).text()
        qty = self.org_out_table.item(row, 5).text()
        to_unit = self.org_out_table.item(row, 4).text()

        ret = QMessageBox.question(self, "确认删除",
                                   f"确定要删除发放给 {to_unit} 的 {equip_name} 出库记录（数量 {qty}）吗？\n"
                                   "同时会删除对应中队的入库记录，库存会恢复。",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ret == QMessageBox.Yes:
            try:
                self.db.cursor.execute("DELETE FROM org_stock_out WHERE id=?", (record_id,))
                self.db.conn.commit()
                self.db.add_log('删除', '机关出库', record_id,
                                f"删除出库记录: {equip_name} {qty}，接收中队: {to_unit}",
                                operator=self.operator_name, role='org')
                self.refresh_org_out_table()
                self.refresh_equipment_table()
            except Exception as e:
                QMessageBox.critical(self, "删除失败", str(e))

    def regen_today_org_out(self):
        """重新生成本日出库单"""
        from datetime import date
        today_str = date.today().strftime('%Y-%m-%d')
        records = self.db.get_org_stock_out_by_date(today_str)
        if not records:
            QMessageBox.information(self, "提示", f"今天（{today_str}）没有出库记录")
            return

        all_items = []
        for r in records:
            equip = self.db.get_equipment_by_id(r['equipment_id'])
            all_items.append({
                'serial_no': r['serial_no'],
                'name': r['name'],
                'spec_model': r['spec_model'],
                'belong_equipment': equip['belong_equipment'] if equip else '',
                'quantity': int(r['quantity']),
                'unit': r['unit'],
                'unit_price': r['unit_price'],
                'total_price': r['total_price'],
                'remark': r['remark'],
                'manufacturer': r['manufacturer'],
                'contract_no': r['contract_no'],
            })

        dlg = OutReceiptGroupDialog(self)
        if dlg.exec_() == QDialog.Accepted:
            output_dir = os.path.join(os.path.dirname(__file__), '出库单')
            os.makedirs(output_dir, exist_ok=True)
            try:
                if dlg.group_type == 'equipment':
                    groups = {}
                    for item in all_items:
                        eq = item.get('belong_equipment', '未知装备') or '未知装备'
                        if eq not in groups:
                            groups[eq] = []
                        groups[eq].append(item)
                else:
                    groups = {today_str: all_items}

                total_files = 0
                for group_name, items in groups.items():
                    first = items[0]
                    files = ReceiptExporter.generate_out_receipt(
                        items=items,
                        output_dir=output_dir,
                        group_name=group_name,
                        contract_no=first.get('contract_no', ''),
                        out_date=today_str,
                        stock_out_person='',
                        operator='',
                        group_type=dlg.group_type
                    )
                    total_files += len(files)

                QMessageBox.information(self, "成功",
                                        f"已重新生成 {total_files} 个出库单文件\n保存位置：{output_dir}")
            except Exception as e:
                QMessageBox.critical(self, "失败", str(e))

    # ==================== 中队入库记录选项卡 ====================
    def create_unit_in_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)

        # ---------- 搜索筛选区域 ----------
        sg = QGroupBox("搜索筛选")
        sl = QHBoxLayout(sg)
        sl.addWidget(QLabel("关键词:"))
        self.unit_in_search = QLineEdit()
        self.unit_in_search.setPlaceholderText("器材名称/规格型号/编号")
        sl.addWidget(self.unit_in_search)
        sl.addWidget(QLabel("质量:"))
        self.unit_in_quality = QComboBox()
        self.unit_in_quality.addItems(['全部'] + QUALITY_LEVELS)
        sl.addWidget(self.unit_in_quality)
        btn_search = QPushButton("搜索")
        btn_search.clicked.connect(self.refresh_unit_in_table)
        sl.addWidget(btn_search)
        btn_reset = QPushButton("重置")
        btn_reset.clicked.connect(lambda: (
            self.unit_in_search.clear(),
            self.unit_in_quality.setCurrentIndex(0),
            self.refresh_unit_in_table()
        ))
        sl.addWidget(btn_reset)
        sl.addStretch()
        layout.addWidget(sg)

        # ---------- 操作按钮行 ----------
        bl = QHBoxLayout()
        btn_add = QPushButton("＋ 新增中队入库")
        btn_add.clicked.connect(self.add_unit_in)
        bl.addWidget(btn_add)
        btn_import = QPushButton("📥 从Excel导入")
        btn_import.clicked.connect(lambda: self.import_unit_excel('in'))
        bl.addWidget(btn_import)
        btn_export = QPushButton("📤 导出Excel")
        btn_export.clicked.connect(self.export_unit_in_excel)
        bl.addWidget(btn_export)
        btn_delete = QPushButton("✕ 删除选中记录")
        btn_delete.clicked.connect(self.delete_unit_record)
        bl.addWidget(btn_delete)
        bl.addStretch()
        btn_ref = QPushButton("↻ 刷新")
        btn_ref.clicked.connect(self.refresh_unit_in_table)
        bl.addWidget(btn_ref)
        layout.addLayout(bl)

        self.unit_in_table = QTableWidget()
        self.unit_in_table.setColumnCount(12)
        self.unit_in_table.setHorizontalHeaderLabels([
            'ID', '器材编号', '器材名称', '规格型号', '入库数量', '计量单位', '总价',
            '入库时间', '来源', '入库人', '经办人', '合同编号'
        ])
        self.unit_in_table.setColumnHidden(0, True)
        self.unit_in_table.horizontalHeader().setStretchLastSection(True)
        self.unit_in_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.unit_in_table.setAlternatingRowColors(True)
        self.unit_in_table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.unit_in_table, stretch=1)

        self.refresh_unit_in_table()
        return w

    def refresh_unit_in_table(self):
        """刷新中队入库记录表格，支持搜索筛选"""
        records = self.db.get_unit_records(self.belong_unit, 'in')
        keyword = self.unit_in_search.text().strip()
        q = self.unit_in_quality.currentText()
        if q == '全部':
            q = ''
        if keyword or q:
            filtered = []
            for r in records:
                if keyword:
                    kw = keyword.lower()
                    match = (kw in (r['name'] or '').lower() or
                             kw in (r['spec_model'] or '').lower() or
                             kw in (r['serial_no'] or '').lower())
                    if not match:
                        continue
                if q and (r['quality_level'] or '') != q:
                    continue
                filtered.append(r)
            records = filtered

        self.unit_in_table.setRowCount(len(records))
        for i, r in enumerate(records):
            self.unit_in_table.setItem(i, 0, QTableWidgetItem(str(r['id'])))
            self.unit_in_table.setItem(i, 1, QTableWidgetItem(r['serial_no'] or ''))
            self.unit_in_table.setItem(i, 2, QTableWidgetItem(r['name'] or ''))
            self.unit_in_table.setItem(i, 3, QTableWidgetItem(r['spec_model'] or ''))
            self.unit_in_table.setItem(i, 4, QTableWidgetItem(str(int(r['quantity']))))
            self.unit_in_table.setItem(i, 5, QTableWidgetItem(r['unit'] or ''))
            self.unit_in_table.setItem(i, 6, QTableWidgetItem(f"¥{r['total_price']:.2f}" if r['total_price'] else ''))
            self.unit_in_table.setItem(i, 7, QTableWidgetItem(r['in_time'] or ''))
            self.unit_in_table.setItem(i, 8, QTableWidgetItem(r['source'] or ''))
            self.unit_in_table.setItem(i, 9, QTableWidgetItem(r['stock_in_person'] or ''))
            self.unit_in_table.setItem(i, 10, QTableWidgetItem(r['operator'] or ''))
            self.unit_in_table.setItem(i, 11, QTableWidgetItem(r['contract_no'] or ''))

    def add_unit_in(self):
        dlg = UnitStockInDialog(self.db, self.belong_unit, self.operator_name, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self.refresh_unit_in_table()
            self.refresh_equipment_table()

    def delete_unit_record(self):
        """删除选中的中队入库/出库记录，同时被出库记录选项卡复用"""
        current_tab = self.tabs.currentWidget()
        table_attr = None
        if current_tab is self.tab_equip:
            return
        # 根据当前选项卡找到对应的表格和记录类型
        for attr_name in ['unit_in_table', 'unit_out_table']:
            if hasattr(self, attr_name):
                table = getattr(self, attr_name)
                if table.isVisible():
                    table_attr = attr_name
                    break
        if not table_attr:
            return
        table = getattr(self, table_attr)
        row = table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "提示", "请先选中一条记录")
            return
        record_id = int(table.item(row, 0).text())
        equip_name = table.item(row, 2).text()
        qty = table.item(row, 4).text()
        ret = QMessageBox.question(self, "确认删除",
                                   f"确定要删除 {equip_name} 的记录（数量 {qty}）吗？\n"
                                   "这将影响当前库存。",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ret == QMessageBox.Yes:
            self.db.cursor.execute("DELETE FROM unit_stock_record WHERE id=?", (record_id,))
            self.db.conn.commit()
            self.db.add_log('删除', '中队记录', record_id,
                            f"删除中队记录: {equip_name} {qty}",
                            operator=self.operator_name, role='unit', belong_unit=self.belong_unit)
            if table_attr == 'unit_in_table':
                self.refresh_unit_in_table()
            else:
                self.refresh_unit_out_table()
            self.refresh_equipment_table()

    # ==================== 中队出库记录选项卡 ====================
    def create_unit_out_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)

        bl = QHBoxLayout()
        btn_add = QPushButton("＋ 新增中队出库")
        btn_add.clicked.connect(self.add_unit_out)
        bl.addWidget(btn_add)

        btn_import = QPushButton("📥 从Excel导入")
        btn_import.clicked.connect(lambda: self.import_unit_excel('out'))
        bl.addWidget(btn_import)

        btn_export = QPushButton("📤 导出Excel")  # 新增
        btn_export.clicked.connect(self.export_unit_out_excel)
        bl.addWidget(btn_export)

        btn_delete = QPushButton("✕ 删除选中记录")
        btn_delete.clicked.connect(self.delete_unit_record)
        bl.addWidget(btn_delete)

        bl.addStretch()
        btn_ref = QPushButton("↻ 刷新")
        btn_ref.clicked.connect(self.refresh_unit_out_table)
        bl.addWidget(btn_ref)
        layout.addLayout(bl)

        self.unit_out_table = QTableWidget()
        self.unit_out_table.setColumnCount(12)
        self.unit_out_table.setHorizontalHeaderLabels([
            'ID', '器材编号', '器材名称', '规格型号', '出库数量', '总价',
            '出库时间', '出库去向', '出库人', '经办人', '合同编号', '备注'
        ])
        self.unit_out_table.setColumnHidden(0, True)
        self.unit_out_table.horizontalHeader().setStretchLastSection(True)
        self.unit_out_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.unit_out_table.setAlternatingRowColors(True)
        self.unit_out_table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.unit_out_table, stretch=1)

        self.refresh_unit_out_table()
        return w

    def refresh_unit_out_table(self):
        records = self.db.get_unit_records(self.belong_unit, 'out')
        self.unit_out_table.setRowCount(len(records))
        for i, r in enumerate(records):
            self.unit_out_table.setItem(i, 0, QTableWidgetItem(str(r['id'])))
            self.unit_out_table.setItem(i, 1, QTableWidgetItem(r['serial_no'] or ''))
            self.unit_out_table.setItem(i, 2, QTableWidgetItem(r['name'] or ''))
            self.unit_out_table.setItem(i, 3, QTableWidgetItem(r['spec_model'] or ''))
            self.unit_out_table.setItem(i, 4, QTableWidgetItem(str(int(r['quantity']))))
            self.unit_out_table.setItem(i, 5, QTableWidgetItem(f"¥{r['total_price']:.2f}" if r['total_price'] else ''))
            self.unit_out_table.setItem(i, 6, QTableWidgetItem(r['out_time'] or ''))
            self.unit_out_table.setItem(i, 7, QTableWidgetItem(r['out_direction'] or ''))
            self.unit_out_table.setItem(i, 8, QTableWidgetItem(r['stock_out_person'] or ''))
            self.unit_out_table.setItem(i, 9, QTableWidgetItem(r['operator'] or ''))
            self.unit_out_table.setItem(i, 10, QTableWidgetItem(r['contract_no'] or ''))
            self.unit_out_table.setItem(i, 11, QTableWidgetItem(r['remark'] or ''))

    def add_unit_out(self):
        dlg = UnitStockOutDialog(self.db, self.belong_unit, self.operator_name, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self.refresh_unit_out_table()
            self.refresh_equipment_table()

    # ==================== 机关库存看板 ====================
    def create_org_dashboard_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)

        sg = QGroupBox("机关总库库存总览")
        sl = QHBoxLayout(sg)
        self.org_c1 = self._card("器材种类", "0")
        self.org_c2 = self._card("入库总量", "0")
        self.org_c3 = self._card("出库总量", "0")
        self.org_c4 = self._card("当前库存", "0")
        sl.addWidget(self.org_c1)
        sl.addWidget(self.org_c2)
        sl.addWidget(self.org_c3)
        sl.addWidget(self.org_c4)
        layout.addWidget(sg)

        qg = QGroupBox("按质量等级")
        ql = QVBoxLayout(qg)
        self.org_qt = QTableWidget()
        self.org_qt.setColumnCount(3)
        self.org_qt.setHorizontalHeaderLabels(['质量等级', '器材数', '库存量'])
        self.org_qt.horizontalHeader().setStretchLastSection(True)
        self.org_qt.setEditTriggers(QTableWidget.NoEditTriggers)
        ql.addWidget(self.org_qt)
        layout.addWidget(qg)

        lg = QGroupBox("低库存器材(≤5)")
        ll = QVBoxLayout(lg)
        self.org_lt = QTableWidget()
        self.org_lt.setColumnCount(3)
        self.org_lt.setHorizontalHeaderLabels(['器材名称', '规格型号', '当前库存'])
        self.org_lt.horizontalHeader().setStretchLastSection(True)
        self.org_lt.setEditTriggers(QTableWidget.NoEditTriggers)
        ll.addWidget(self.org_lt)
        layout.addWidget(lg)

        # 按钮行
        bl = QHBoxLayout()
        btn_refresh = QPushButton("↻ 刷新看板")
        btn_refresh.clicked.connect(self.refresh_org_dash)
        bl.addWidget(btn_refresh)
        btn_export = QPushButton("📤 导出库存统计表")
        btn_export.clicked.connect(self.export_org_stock_excel)
        bl.addWidget(btn_export)
        bl.addStretch()
        layout.addLayout(bl)

        self.refresh_org_dash()
        return w

    def refresh_org_dash(self):
        """更新机关看板数据"""
        s = self.db.get_org_stock_summary()
        self._card_update(self.org_c1, str(s['total_types']))
        self._card_update(self.org_c2, str(int(s['total_in'])))
        self._card_update(self.org_c3, str(int(s['total_out'])))
        self._card_update(self.org_c4, str(int(s['current_stock'])))

        # 质量等级统计
        qs = s['quality_stats']
        self.org_qt.setRowCount(len(qs))
        for i, r in enumerate(qs):
            self.org_qt.setItem(i, 0, QTableWidgetItem(r['quality_level'] or '未设置'))
            self.org_qt.setItem(i, 1, QTableWidgetItem(str(r['count'])))
            st = r['total_stock'] or 0
            si = QTableWidgetItem(str(st))
            if st <= 0:
                si.setForeground(Qt.red)
            self.org_qt.setItem(i, 2, si)

        # 低库存器材
        low = self.db.get_org_low_stock(5)
        self.org_lt.setRowCount(len(low))
        for i, r in enumerate(low):
            self.org_lt.setItem(i, 0, QTableWidgetItem(r['name'] or ''))
            self.org_lt.setItem(i, 1, QTableWidgetItem(r['spec_model'] or ''))
            si = QTableWidgetItem(str(r['current_stock']))
            si.setForeground(Qt.red)
            self.org_lt.setItem(i, 2, si)

    # ==================== 中队库存看板 ====================
    def create_unit_dashboard_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)

        sg = QGroupBox(f"{self.belong_unit}库存总览")
        sl = QHBoxLayout(sg)
        self.unit_c1 = self._card("器材种类", "0")
        self.unit_c2 = self._card("入库总量", "0")
        self.unit_c3 = self._card("出库总量", "0")
        self.unit_c4 = self._card("当前库存", "0")
        sl.addWidget(self.unit_c1)
        sl.addWidget(self.unit_c2)
        sl.addWidget(self.unit_c3)
        sl.addWidget(self.unit_c4)
        layout.addWidget(sg)

        qg = QGroupBox("按质量等级")
        ql = QVBoxLayout(qg)
        self.unit_qt = QTableWidget()
        self.unit_qt.setColumnCount(3)
        self.unit_qt.setHorizontalHeaderLabels(['质量等级', '器材数', '库存量'])
        self.unit_qt.horizontalHeader().setStretchLastSection(True)
        self.unit_qt.setEditTriggers(QTableWidget.NoEditTriggers)
        ql.addWidget(self.unit_qt)
        layout.addWidget(qg)

        lg = QGroupBox("低库存器材(≤5)")
        ll = QVBoxLayout(lg)
        self.unit_lt = QTableWidget()
        self.unit_lt.setColumnCount(3)
        self.unit_lt.setHorizontalHeaderLabels(['器材名称', '规格型号', '当前库存'])
        self.unit_lt.horizontalHeader().setStretchLastSection(True)
        self.unit_lt.setEditTriggers(QTableWidget.NoEditTriggers)
        ll.addWidget(self.unit_lt)
        layout.addWidget(lg)

        bl = QHBoxLayout()
        btn_refresh = QPushButton("↻ 刷新看板")
        btn_refresh.clicked.connect(self.refresh_unit_dash)
        bl.addWidget(btn_refresh)
        btn_export = QPushButton("📤 导出库存统计表")
        btn_export.clicked.connect(self.export_unit_stock_excel)
        bl.addWidget(btn_export)
        bl.addStretch()
        layout.addLayout(bl)

        self.refresh_unit_dash()
        return w

    def refresh_unit_dash(self):
        """更新中队看板数据"""
        s = self.db.get_unit_stock_summary(self.belong_unit)
        self._card_update(self.unit_c1, str(s['total_types']))
        self._card_update(self.unit_c2, str(int(s['total_in'])))
        self._card_update(self.unit_c3, str(int(s['total_out'])))
        self._card_update(self.unit_c4, str(int(s['current_stock'])))

        qs = s['quality_stats']
        self.unit_qt.setRowCount(len(qs))
        for i, r in enumerate(qs):
            self.unit_qt.setItem(i, 0, QTableWidgetItem(r['quality_level'] or '未设置'))
            self.unit_qt.setItem(i, 1, QTableWidgetItem(str(r['count'])))
            st = r['total_stock'] or 0
            si = QTableWidgetItem(str(st))
            if st <= 0:
                si.setForeground(Qt.red)
            self.unit_qt.setItem(i, 2, si)

        low = self.db.get_unit_low_stock(self.belong_unit, 5)
        self.unit_lt.setRowCount(len(low))
        for i, r in enumerate(low):
            self.unit_lt.setItem(i, 0, QTableWidgetItem(r['name'] or ''))
            self.unit_lt.setItem(i, 1, QTableWidgetItem(r['spec_model'] or ''))
            si = QTableWidgetItem(str(int(r['current_stock'])))
            si.setForeground(Qt.red)
            self.unit_lt.setItem(i, 2, si)

    # ==================== 统计卡片辅助方法 ====================
    def _card(self, title, value):
        f = QFrame()
        f.setStyleSheet("QFrame{background:#f5f5f5;border-radius:8px;padding:10px;margin:5px;}")
        l = QVBoxLayout(f)
        t = QLabel(title)
        t.setAlignment(Qt.AlignCenter)
        t.setStyleSheet("font-size:14px;color:#666;")  # 从 11px 改为 14px
        l.addWidget(t)
        v = QLabel(value)
        v.setObjectName('v')
        v.setAlignment(Qt.AlignCenter)
        v.setStyleSheet("font-size:28px;font-weight:bold;color:#2196F3;")  # 从 22px 改为 28px
        l.addWidget(v)
        return f

    def _card_update(self, card, value):
        """更新卡片数值"""
        card.findChild(QLabel, 'v').setText(value)

    # ==================== 操作日志选项卡 ====================
    def create_log_tab(self):
        """操作日志：按类型和日期筛选，清理旧日志"""
        w = QWidget()
        layout = QVBoxLayout(w)

        fg = QGroupBox("筛选")
        fl = QHBoxLayout(fg)
        fl.addWidget(QLabel("操作类型:"))
        self.log_act = QComboBox()
        self.log_act.addItems(['全部', '新增', '编辑', '删除', '入库', '出库', '备份', '恢复', '登录', '退出'])
        fl.addWidget(self.log_act)
        fl.addWidget(QLabel("日期:"))
        self.log_d1 = QDateEdit()
        self.log_d1.setCalendarPopup(True)
        self.log_d1.setDate(QDate.currentDate().addMonths(-1))
        fl.addWidget(self.log_d1)
        fl.addWidget(QLabel("至"))
        self.log_d2 = QDateEdit()
        self.log_d2.setCalendarPopup(True)
        self.log_d2.setDate(QDate.currentDate())
        fl.addWidget(self.log_d2)
        btn = QPushButton("查询")
        btn.clicked.connect(self.refresh_log)
        fl.addWidget(btn)
        fl.addStretch()
        layout.addWidget(fg)

        self.log_table = QTableWidget()
        self.log_table.setColumnCount(7)
        self.log_table.setHorizontalHeaderLabels(['时间', '操作', '目标类型', '目标ID', '详情', '操作人', '角色/单位'])
        self.log_table.horizontalHeader().setStretchLastSection(True)
        self.log_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.log_table.setAlternatingRowColors(True)
        layout.addWidget(self.log_table, stretch=1)

        bl = QHBoxLayout()
        btn_r = QPushButton("↻ 刷新")
        btn_r.clicked.connect(self.refresh_log)
        bl.addWidget(btn_r)
        bl.addStretch()
        btn_c = QPushButton("清理旧日志")
        btn_c.clicked.connect(self.clean_logs)
        bl.addWidget(btn_c)
        layout.addLayout(bl)

        self.refresh_log()
        return w

    def refresh_log(self):
        """刷新日志表格"""
        act = self.log_act.currentText()
        if act == '全部':
            act = ''
        logs = self.db.get_logs(
            action=act,
            date_from=self.log_d1.date().toString('yyyy-MM-dd'),
            date_to=self.log_d2.date().toString('yyyy-MM-dd')
        )
        self.log_table.setRowCount(len(logs))
        for i, log in enumerate(logs):
            self.log_table.setItem(i, 0, QTableWidgetItem(log['created_at'] or ''))
            self.log_table.setItem(i, 1, QTableWidgetItem(log['action'] or ''))
            self.log_table.setItem(i, 2, QTableWidgetItem(log['target_type'] or ''))
            self.log_table.setItem(i, 3, QTableWidgetItem(str(log['target_id']) if log['target_id'] else ''))
            self.log_table.setItem(i, 4, QTableWidgetItem(log['detail'] or ''))
            self.log_table.setItem(i, 5, QTableWidgetItem(log['operator'] or ''))
            role_str = '机关' if log['role'] == 'org' else '中队'
            self.log_table.setItem(i, 6, QTableWidgetItem(f"{role_str}/{log['belong_unit'] or ''}"))

    def clean_logs(self):
        """清理几个月前的旧日志"""
        m, ok = QInputDialog.getInt(self, "清理日志", "删除几个月前的日志？", 3, 1, 24, 1)
        if ok:
            before = QDate.currentDate().addMonths(-m).toString('yyyy-MM-dd')
            deleted = self.db.clear_logs_before(before)
            QMessageBox.information(self, "完成", f"已清理 {deleted} 条日志")
            self.refresh_log()

    # ==================== 数据备份/恢复选项卡 ====================
    def create_backup_tab(self):
        """数据库备份与恢复"""
        w = QWidget()
        layout = QVBoxLayout(w)

        ig = QGroupBox("数据库信息")
        il = QVBoxLayout(ig)
        self.db_path_label = QLabel()
        self.db_size_label = QLabel()
        self.db_time_label = QLabel()
        il.addWidget(self.db_path_label)
        il.addWidget(self.db_size_label)
        il.addWidget(self.db_time_label)
        btn_i = QPushButton("刷新信息")
        btn_i.clicked.connect(self.refresh_db_info)
        il.addWidget(btn_i)
        layout.addWidget(ig)

        bg = QGroupBox("备份与恢复")
        bl = QVBoxLayout(bg)
        btn_b = QPushButton("💾 备份数据库")
        btn_b.setMinimumHeight(40)
        btn_b.clicked.connect(self.do_backup)
        bl.addWidget(btn_b)
        btn_r = QPushButton("📥 恢复数据库(⚠谨慎操作)")
        btn_r.setMinimumHeight(40)
        btn_r.setStyleSheet("color:red;font-weight:bold;")
        btn_r.clicked.connect(self.do_restore)
        bl.addWidget(btn_r)
        layout.addWidget(bg)
        layout.addStretch()

        self.refresh_db_info()
        return w

    def refresh_db_info(self):
        """显示数据库文件信息"""
        info = self.db.get_database_info()
        if info:
            self.db_path_label.setText(f"路径: {info['path']}")
            self.db_size_label.setText(f"大小: {info['size_mb']} MB")
            self.db_time_label.setText(f"修改时间: {info['modified']}")

    def do_backup(self):
        """备份数据库"""
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        fp, _ = QFileDialog.getSaveFileName(self, "备份", f"equipment_backup_{ts}.bak", "备份文件(*.bak)")
        if fp:
            try:
                self.db.backup_database(fp)
                self.db.add_log('备份', '数据库', None, f'备份到:{fp}',
                                operator=self.operator_name, role=self.role, belong_unit=self.belong_unit)
                QMessageBox.information(self, "成功", f"备份完成:\n{fp}")
            except Exception as e:
                QMessageBox.critical(self, "失败", str(e))

    def do_restore(self):
        """恢复数据库（需密码验证）"""
        # 密码验证
        password, ok = QInputDialog.getText(
            self, "安全验证",
            "请输入恢复操作密码：",
            QLineEdit.Password, ""
        )
        if not ok or password != "danger":
            if ok:
                QMessageBox.critical(self, "密码错误", "恢复密码错误，操作已取消！")
            return

        # 二次确认
        if QMessageBox.warning(self, "警告",
                               "恢复将覆盖当前所有数据！确定继续？",
                               QMessageBox.Yes | QMessageBox.No,
                               QMessageBox.No) != QMessageBox.Yes:
            return

        fp, _ = QFileDialog.getOpenFileName(self, "选择备份文件", "",
                                             "备份文件(*.bak *.db);;所有文件(*.*)")
        if fp:
            try:
                self.db.restore_database(fp)
                self.db.add_log('恢复', '数据库', None, f'从{fp}恢复',
                                operator=self.operator_name, role=self.role,
                                belong_unit=self.belong_unit)
                QMessageBox.information(self, "成功",
                                        "数据库已恢复，请重启程序。")
            except Exception as e:
                QMessageBox.critical(self, "恢复失败", str(e))

    # ==================== 到期提醒 ====================
    def check_expiry(self):
        """检查存储寿命即将到期的器材，并弹出提醒"""
        expiring = self.db.get_expiring_soon(30)
        if expiring:
            msg = "以下器材存储寿命即将在30天内到期：\n\n"
            for e in expiring[:10]:
                msg += f"• {e['name']} - {e['spec_model']} (到期日: {e['expire_date']})\n"
            if len(expiring) > 10:
                msg += f"\n...共 {len(expiring)} 条"
            QMessageBox.warning(self, "存储寿命到期提醒", msg)

    # ==================== 程序退出 ====================
    def closeEvent(self, event):
        """主窗口关闭时记录日志并关闭数据库"""
        self.db.add_log('退出', '系统', None, '退出系统',
                        operator=self.operator_name, role=self.role, belong_unit=self.belong_unit)
        self.db.close()
        event.accept()


# ==================== 程序入口 ====================
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    font = QFont("Noto Sans CJK SC", 10)
    if not font.exactMatch():
        font = QFont("WenQuanYi Micro Hei", 10)
    app.setFont(font)

    login = LoginDialog()
    if login.exec_() != QDialog.Accepted:
        sys.exit(0)

    window = MainWindow(login.role, login.belong_unit, login.operator)
    window.show()
    sys.exit(app.exec_())