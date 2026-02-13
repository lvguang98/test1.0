#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工伤案件管理系统 - 主程序
"""
import os
import sys
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.uic import loadUi
from PyQt5.QtCore import QSettings, Qt
from openpyxl import load_workbook
from config_manager import ConfigManager


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        # 1. 加载界面
        loadUi("main_window.ui", self)
        self.setWindowTitle("工伤案件管理系统")

        # 2. 初始化配置管理器
        self.config = ConfigManager()

        # 3. 加载Excel数据到ComboBox
        self.load_excel_to_combobox()

        # 3.1 设置ComboBox的自动完成和失去焦点保存功能
        self.setup_combobox_autosave()

        # 3.2 连接删除按钮
        self.setup_delete_buttons()

        # 4. 加载保存的配置
        self.load_config()

        # 5. 连接信号
        self.checkBox_remember.stateChanged.connect(self.on_remember_changed)
        self.btn_generate_record.clicked.connect(self.on_generate_record)
        # 身份证号框失去焦点
        self.lineEdit_id_card.editingFinished.connect(self.auto_calculate_id_info)

        # 6. 根据记住状态更新界面
        self.update_ui()

        # 连接人员类型切换信号
        self.radio_self.toggled.connect(self.on_person_type_changed)
        self.radio_witness.toggled.connect(self.on_person_type_changed)
        self.radio_legal_entity.toggled.connect(self.on_person_type_changed)

        # 本人姓名失去焦点时自动填入受伤职工
        self.lineEdit_name.editingFinished.connect(self.auto_fill_injured_worker)

    def auto_fill_injured_worker(self):
        """本人姓名输入完成时自动填入受伤职工"""
        if self.check_person_type() == "本人":
            name = self.lineEdit_name.text().strip()
            if name:
                self.lineEdit_injured_worker.setText(name)

    def on_person_type_changed(self):
        """人员类型切换时的处理"""
        if self.sender().isChecked():
            person_type = self.check_person_type()
            self.statusBar().showMessage(f"当前人员类型: {person_type}", 1500)

            # 切换时清空相关字段
            if person_type in ["本人", "证人", "法人"]:
                self.clear_person_fields()

    def clear_person_fields(self):
        """清空人员信息字段"""
        self.lineEdit_name.clear()
        self.lineEdit_age.clear()
        self.comboBox_gender.setCurrentIndex(-1)
        self.lineEdit_id_card.clear()
        self.lineEdit_id_address.clear()
        self.lineEdit_current_address.clear()
        self.lineEdit_phone.clear()
        self.lineEdit_position.clear()

    def setup_delete_buttons(self):
        """设置删除按钮功能"""
        self.btn_delete_employer.clicked.connect(
            lambda: self.delete_from_excel('comboBox_employer', self.employer_list, "用人单位名称汇总.xlsx", "用人单位")
        )
        self.btn_delete_work_unit.clicked.connect(
            lambda: self.delete_from_excel('comboBox_work_unit', self.work_unit_list, "用工单位名称汇总.xlsx", "用工单位")
        )
        self.btn_delete_workplace.clicked.connect(
            lambda: self.delete_from_excel('comboBox_workplace', self.workplace_list, "工作场所名称汇总.xlsx", "工作场所")
        )

    def delete_from_excel(self, combobox_name, data_list, filename, column_name):
        """从Excel删除当前选中的项目"""
        # 获取对应的ComboBox
        combobox = getattr(self, combobox_name)

        # 获取当前选中的文本
        selected_text = combobox.currentText().strip()

        if not selected_text:
            self.statusBar().showMessage("请先选择要删除的项目", 2000)
            return

        # 确认对话框
        from PyQt5.QtWidgets import QMessageBox
        reply = QMessageBox.question(
            self, '确认删除',
            f'确定要删除 "{selected_text}" 吗？',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )

        if reply == QMessageBox.No:
            return

        try:
            # 1. 从内存列表中删除
            if selected_text in data_list:
                data_list.remove(selected_text)

            # 2. 从ComboBox中删除
            index = combobox.findText(selected_text)
            if index >= 0:
                combobox.removeItem(index)

            # 3. 从Excel文件中删除
            current_dir = os.path.dirname(os.path.abspath(__file__))
            filepath = os.path.join(current_dir, filename)

            if os.path.exists(filepath):
                wb = load_workbook(filepath)
                ws = wb.active

                # 找到要删除的行
                row_to_delete = None
                for row in range(1, ws.max_row + 1):
                    cell_value = ws.cell(row=row, column=1).value
                    if cell_value and str(cell_value).strip() == selected_text:
                        row_to_delete = row
                        break

                # 删除行
                if row_to_delete:
                    ws.delete_rows(row_to_delete)
                    wb.save(filepath)
                    self.statusBar().showMessage(f'已删除: {selected_text}', 3000)
                    print(f"✅ 已从Excel删除: {selected_text}")
                else:
                    self.statusBar().showMessage("未在Excel中找到该项目", 3000)
            else:
                self.statusBar().showMessage("Excel文件不存在", 3000)

            # 4. 清空当前选择
            combobox.setCurrentIndex(-1)
            combobox.setCurrentText("")

        except Exception as e:
            self.statusBar().showMessage(f"删除失败: {str(e)}", 3000)
            print(f"❌ 删除失败: {e}")

    def setup_combobox_autosave(self):
        """设置ComboBox的自动完成和失去焦点保存功能"""
        # 为每个ComboBox设置相同的功能
        for combobox_name in ['comboBox_employer', 'comboBox_work_unit', 'comboBox_workplace']:
            combobox = getattr(self, combobox_name)

            # 设置可编辑
            combobox.setEditable(True)

            # 设置自动完成，显示最多3个相似项
            from PyQt5.QtCore import Qt
            from PyQt5.QtWidgets import QCompleter

            # 获取当前列表数据
            if combobox_name == 'comboBox_employer':
                data_list = self.employer_list
            elif combobox_name == 'comboBox_work_unit':
                data_list = self.work_unit_list
            else:  # comboBox_workplace
                data_list = self.workplace_list

            # 创建自动完成器
            completer = QCompleter(data_list)
            completer.setFilterMode(Qt.MatchContains)  # 包含匹配
            completer.setMaxVisibleItems(3)  # 最多显示3个
            combobox.setCompleter(completer)

            # 获取ComboBox内部的QLineEdit并连接失去焦点事件
            line_edit = combobox.lineEdit()
            line_edit.editingFinished.connect(
                lambda le=line_edit, cb=combobox, name=combobox_name, lst=data_list:
                self.on_combobox_editing_finished(le, cb, name, lst)
            )

    def on_combobox_editing_finished(self, line_edit, combobox, combobox_name, current_list):
        """ComboBox失去焦点时的处理"""
        # 获取用户输入的文本
        user_input = line_edit.text().strip()

        if not user_input:
            return  # 如果输入为空，不处理

        # 检查是否已经在列表中
        if user_input in current_list:
            return  # 如果已经在列表中，不重复添加

        # 如果不在列表中，保存到Excel
        self.save_to_excel(combobox_name, user_input, current_list)

        # 添加到内存列表和ComboBox
        current_list.append(user_input)
        combobox.addItem(user_input)

        # 保持用户输入的内容显示在界面上
        combobox.setCurrentText(user_input)

    def save_to_excel(self, combobox_name, new_item, current_list):
        """保存新项目到对应的Excel文件"""
        # 确定文件名和列名
        if combobox_name == 'comboBox_employer':
            filename = "用人单位名称汇总.xlsx"
            column_name = "用人单位"
        elif combobox_name == 'comboBox_work_unit':
            filename = "用工单位名称汇总.xlsx"
            column_name = "用工单位"
        else:  # comboBox_workplace
            filename = "工作场所名称汇总.xlsx"
            column_name = "工作场所"

        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            filepath = os.path.join(current_dir, filename)

            # 如果文件存在，追加数据
            if os.path.exists(filepath):
                wb = load_workbook(filepath)
                ws = wb.active

                # 找到第一个空行
                row = 1
                while ws.cell(row=row, column=1).value is not None:
                    row += 1

                # 写入新数据
                ws.cell(row=row, column=1, value=new_item)
                wb.save(filepath)
            else:
                # 文件不存在，创建新文件
                wb = load_workbook()
                ws = wb.active
                ws.title = "汇总表"
                ws.cell(row=1, column=1, value=column_name)
                ws.cell(row=2, column=1, value=new_item)
                wb.save(filepath)

        except Exception as e:
            print(f"保存到Excel失败: {e}")

    def auto_calculate_id_info(self):
        """自动计算身份证信息"""
        id_card = self.lineEdit_id_card.text().strip()
        if id_card:
            _, age, gender = self.calculate_id_info(id_card)
            if age:
                self.lineEdit_age.setText(str(age))
            if gender:
                self.comboBox_gender.setCurrentText(gender)

    def load_excel_to_combobox(self):
        """从Excel文件加载数据到ComboBox"""
        current_dir = os.path.dirname(os.path.abspath(__file__))

        # 加载用人单位 - 确保属性存在
        self.employer_list = self.load_excel_data(os.path.join(current_dir, "用人单位名称汇总.xlsx"))
        self.comboBox_employer.addItems(self.employer_list)

        # 加载用工单位 - 确保属性存在
        self.work_unit_list = self.load_excel_data(os.path.join(current_dir, "用工单位名称汇总.xlsx"))
        self.comboBox_work_unit.addItems(self.work_unit_list)

        # 加载工作场所 - 确保属性存在
        self.workplace_list = self.load_excel_data(os.path.join(current_dir, "工作场所名称汇总.xlsx"))
        self.comboBox_workplace.addItems(self.workplace_list)

    def load_excel_data(self, filepath):
        """从Excel文件加载数据到列表"""
        data_list = []

        try:
            if os.path.exists(filepath):
                wb = load_workbook(filepath)
                ws = wb.active

                # 读取第一列所有非空数据
                for row in ws.iter_rows(min_row=1, max_col=1, values_only=True):
                    if row[0] and str(row[0]).strip():
                        data_list.append(str(row[0]).strip())
        except Exception as e:
            print(f"读取Excel失败 {filepath}: {e}")

        return data_list

    def load_config(self):
        """加载配置到界面"""
        config = self.config.load_config()

        # 设置控件内容
        self.lineEdit_operator.setText(config['operator'])
        self.lineEdit_api_url.setText(config['api_url'])
        self.lineEdit_api_key.setText(config['api_key'])
        self.checkBox_remember.setChecked(config['remember'])

    def update_ui(self):
        """更新界面状态"""
        remember = self.checkBox_remember.isChecked()

        # 设置输入框是否可编辑
        self.lineEdit_operator.setEnabled(not remember)
        self.lineEdit_api_url.setEnabled(not remember)
        self.lineEdit_api_key.setEnabled(not remember)

        # 设置样式
        if remember:
            style = "background-color: #f0f0f0; color: #666;"
            self.statusBar().showMessage("配置已记住，取消勾选可修改", 2000)
        else:
            style = ""

        self.lineEdit_operator.setStyleSheet(f"QLineEdit {{ {style} }}")
        self.lineEdit_api_url.setStyleSheet(f"QLineEdit {{ {style} }}")
        self.lineEdit_api_key.setStyleSheet(f"QLineEdit {{ {style} }}")

    def on_remember_changed(self):
        """记住我复选框状态变化"""
        remember = self.checkBox_remember.isChecked()

        if remember:
            # 保存当前配置
            operator = self.lineEdit_operator.text().strip()
            api_url = self.lineEdit_api_url.text().strip()
            api_key = self.lineEdit_api_key.text().strip()

            self.config.save_config(operator, api_url, api_key, True)
            self.statusBar().showMessage("配置已保存", 1500)
        else:
            # 清除配置
            self.config.clear_config()
            self.statusBar().showMessage("配置已清除", 1500)

        # 更新界面
        self.update_ui()

    def check_case_type(self):
        """检查案件类型"""
        is_personal = self.checkBox_personal.isChecked()
        is_death = self.checkBox_death.isChecked()

        if is_personal and is_death:
            return "个人申请死亡案件"
        elif is_personal:
            return "个人案件"
        elif is_death:
            return "死亡案件"
        else:
            return "普通案件"

    def check_person_type(self):
        """检查人员类型"""
        if self.radio_self.isChecked():
            return "本人"
        elif self.radio_witness.isChecked():
            return "证人"
        elif self.radio_legal_entity.isChecked():
            return "法人"

    def calculate_id_info(self, id_card):
        """根据身份证号计算年龄和性别"""
        if len(id_card) != 18:
            return id_card, None, None

        # 提取出生年月日
        birth_year = int(id_card[6:10])
        birth_month = int(id_card[10:12])
        birth_day = int(id_card[12:14])

        # 计算年龄
        current_year = datetime.now().year
        current_month = datetime.now().month
        current_day = datetime.now().day

        age = current_year - birth_year
        if current_month < birth_month or (current_month == birth_month and current_day < birth_day):
            age -= 1

        # 计算性别
        gender_num = int(id_card[16])
        gender = "男" if gender_num % 2 == 1 else "女"

        return id_card, age, gender

    def on_generate_record(self):
        """生成笔录按钮点击"""
        # 1. 收集数据
        data = self.collect_form_data()

        # 2. 根据人员类型分流
        if data['人员类型'] == "本人":
            self.handle_person_case(data)
        elif data['人员类型'] == "证人":
            self.handle_witness_case(data)
        elif data['人员类型'] == "法人":
            self.handle_legal_case(data)

    def handle_person_case(self, data):
        """处理本人案件"""
        # ========== 检查是否已有案本 ==========
        import json
        index_file = os.path.join(os.path.dirname(__file__), "cases_index.json")

        if os.path.exists(index_file):
            with open(index_file, 'r', encoding='utf-8') as f:
                index_data = json.load(f)

            # 搜索同名+同身份证的案件
            same_person_cases = []
            for case in index_data.get('cases', []):
                if (case['person_name'] == data['受伤职工'] and
                        case['id_card'] == data['本人身份证号']):
                    same_person_cases.append(case)

            # 如果有多个，让用户选择
            if same_person_cases:
                selected_case = self.show_case_selection_dialog(
                    data['受伤职工'],
                    same_person_cases,
                    data['本人身份证号']
                )

                if selected_case == "new":
                    # 用户选“新建” → 继续往下走新建逻辑
                    pass
                elif selected_case:
                    # 用户选了具体案本 → 关联并打开
                    # 1. 把本人信息填到界面
                    person_info = selected_case.get('person_info', {})
                    self.lineEdit_name.setText(person_info.get('name', ''))
                    self.lineEdit_id_card.setText(selected_case.get('id_card', ''))
                    self.lineEdit_phone.setText(person_info.get('phone', ''))
                    if selected_case.get('id_card'):
                        self.auto_calculate_id_info()

                    # 2. 检查是否有本人笔录文件
                    case_folder = os.path.join(os.path.dirname(__file__), selected_case['folder_path'])
                    transcript_file = os.path.join(case_folder, f"{selected_case['case_number']}_笔录.docx")

                    if os.path.exists(transcript_file):
                        os.startfile(transcript_file)
                        self.statusBar().showMessage(f"已打开案本: {selected_case['case_number']}", 3000)
                    else:
                        self.statusBar().showMessage(f"该案本无本人笔录文件", 3000)

                    return  # 结束，不新建
                else:
                    return  # 用户取消对话框

        # ========== 原有新建逻辑 ==========
        case_number = self.generate_case_number(data['受伤职工'])
        data['案本号'] = case_number
        year_folder = self.get_current_year_folder()
        case_folder = os.path.join(year_folder, case_number)
        os.makedirs(case_folder, exist_ok=True)
        self.update_case_index(case_number, data['受伤职工'], data)
        self.generate_transcript(case_folder, "本人普通案件模板.docx", data)

    def handle_witness_case(self, data):
        """处理证人案件"""
        injured_name = data['受伤职工']  # 受伤职工姓名
        witness_name = data['证人姓名']  # 证人姓名（从界面获取）
        witness_id = data['证人身份证号']  # 证人身份证号

        # ========== 1. 查找本人文件夹 ==========
        year_folder = self.get_current_year_folder()

        # 搜索所有案本文件夹，找出包含受伤职工姓名的
        matching_folders = []
        if os.path.exists(year_folder):
            for folder in os.listdir(year_folder):
                # 案本号格式：前缀-姓名-序号，例如 "GS-张三-001"
                parts = folder.split('-')
                if len(parts) >= 2 and parts[1] == injured_name:
                    matching_folders.append(folder)

        if not matching_folders:
            # ========== 情况1：没有本人文件夹 ==========
            reply = QMessageBox.question(
                self, '未找到本人案本',
                f'未找到伤者 "{injured_name}" 的案本文件夹\n是否新建案本？',
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )

            if reply == QMessageBox.No:
                return

            # 用受伤职工姓名创建新案本
            case_number = self.generate_case_number(injured_name)
            data['案本号'] = case_number
            case_folder = os.path.join(year_folder, case_number)
            os.makedirs(case_folder, exist_ok=True)

            # 生成第一个证人笔录
            self.create_witness_transcript(case_folder, data, witness_number=1)
            return

        # ========== 情况2：有本人文件夹，取最新的一个 ==========
        case_folder_name = sorted(matching_folders)[-1]  # 取最新的
        case_number = case_folder_name
        data['案本号'] = case_number
        case_folder = os.path.join(year_folder, case_folder_name)

        # ========== 3. 查找该文件夹下所有证人笔录 ==========
        witness_files = []
        if os.path.exists(case_folder):
            for file in os.listdir(case_folder):
                if file.endswith('.docx') and '证人' in file:
                    witness_files.append(file)

        if not witness_files:
            # ========== 情况2.1：没有证人笔录，直接生成第一个 ==========
            self.create_witness_transcript(case_folder, data, witness_number=1)
            return

        # ========== 情况2.2：已有证人笔录，检查是否同一证人 ==========
        # 这里简化处理：遍历所有证人文件，看是否已有当前证人
        # 实际项目中可能需要读取文件内容比对身份证

        import re
        witness_exists = False
        max_number = 0

        for file in witness_files:
            # 文件名格式：受伤职工姓名_证人XX_证人姓名.docx
            # 例如：张三_证人01_李四.docx
            match = re.search(r'证人(\d+)_(.+?)\.docx', file)
            if match:
                num = int(match.group(1))
                existing_witness_name = match.group(2)
                max_number = max(max_number, num)

                # 如果证人姓名相同，视为同一证人
                if existing_witness_name == witness_name:
                    witness_exists = True
                    existing_file = os.path.join(case_folder, file)

        if witness_exists:
            # ========== 情况2.2.1：同一证人，询问关联或新建 ==========
            reply = QMessageBox.question(
                self, '证人已存在',
                f'证人 "{witness_name}" 已有笔录\n是否打开？\n\n选“是”=打开\n选“否”=新建另一份',
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )

            if reply == QMessageBox.Yes:
                os.startfile(existing_file)
                self.statusBar().showMessage(f"已打开证人笔录", 3000)
            else:
                # 新建另一份（编号+1）
                self.create_witness_transcript(case_folder, data, witness_number=max_number + 1)
        else:
            # ========== 情况2.2.2：新证人，直接生成 ==========
            self.create_witness_transcript(case_folder, data, witness_number=max_number + 1)

    # 和证人方法同一个范围的的方法，可能有补充和调整
    def create_witness_transcript(self, case_folder, data, witness_number):
        """生成证人笔录"""
        injured_name = data['受伤职工']
        witness_name = data['证人姓名']

        # 生成文件名：受伤职工姓名_证人XX_证人姓名.docx
        filename = f"{injured_name}_证人{witness_number:02d}_{witness_name}.docx"
        filepath = os.path.join(case_folder, filename)

        # 使用证人模板
        template_path = os.path.join(os.path.dirname(__file__), "templates", "证人模板.docx")

        if not os.path.exists(template_path):
            self.statusBar().showMessage("证人模板不存在", 3000)
            return False

        from docx import Document
        doc = Document(template_path)

        # 替换占位符（需要你根据模板实际占位符调整）
        placeholders = {
            '受伤职工': injured_name,
            '证人姓名': witness_name,
            '证人身份证': data.get('证人身份证号', ''),
            '证人电话': data.get('证人电话', ''),
            '当前日期': datetime.now().strftime('%Y年%m月%d日'),
            '当前时间': datetime.now().strftime('%H时%M分'),
        }

        for paragraph in doc.paragraphs:
            text = paragraph.text
            for key, value in placeholders.items():
                if f"{{{key}}}" in text:
                    text = text.replace(f"{{{key}}}", value)
            paragraph.text = text

        doc.save(filepath)
        os.startfile(filepath)

        self.statusBar().showMessage(f"证人笔录已生成: {filename}", 3000)
        return True

    def handle_legal_case(self, data):
        """处理法人案件"""
        case_number = self.generate_case_number(data['受伤职工'])
        data['案本号'] = case_number

        year_folder = self.get_current_year_folder()
        case_folder = os.path.join(year_folder, case_number)
        os.makedirs(case_folder, exist_ok=True)

        self.generate_transcript(case_folder, "法人模板.docx", data)

    def search_same_name_cases(self, name, id_card):
        """搜索同名案件"""
        cases = []

        # 读取索引文件
        index_file = os.path.join(os.path.dirname(__file__), "cases_index.json")

        if os.path.exists(index_file):
            try:
                import json
                with open(index_file, 'r', encoding='utf-8') as f:
                    index_data = json.load(f)

                for case in index_data.get('cases', []):
                    if case['person_name'] == name:
                        # 检查身份证号（如果有）
                        case_id = case.get('id_card', '')
                        if id_card and case_id:
                            # 有身份证输入，进行比对
                            if id_card == case_id:
                                case['match_type'] = '身份证完全匹配'
                            else:
                                case['match_type'] = '姓名匹配(身份证不同)'
                        else:
                            case['match_type'] = '姓名匹配'

                        cases.append(case)

            except Exception as e:
                print(f"读取索引文件失败: {e}")

        return cases

    def show_case_selection_dialog(self, name, cases, id_card):
        """显示案件选择对话框"""
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QRadioButton, QButtonGroup, QPushButton, QHBoxLayout

        dialog = QDialog(self)
        dialog.setWindowTitle("发现同名案件")
        dialog.resize(400, 300)

        layout = QVBoxLayout()

        # 标题
        if id_card:
            title = f'发现与"{name}"(身份证:{id_card[-4:]})同名的案件:'
        else:
            title = f'发现与"{name}"同名的案件:'

        layout.addWidget(QLabel(title))

        # 创建单选按钮组
        button_group = QButtonGroup()

        # 添加"新建案件"选项
        new_case_radio = QRadioButton("新建案件（不关联已有）")
        new_case_radio.setChecked(True)
        button_group.addButton(new_case_radio, 0)
        layout.addWidget(new_case_radio)

        layout.addWidget(QLabel("已有案本:"))

        # 添加已有案件选项
        for i, case in enumerate(cases, 1):
            case_num = case['case_number']
            case_id = case.get('id_card', '')

            if case_id:
                # 显示身份证后4位
                id_display = case_id[-4:] if len(case_id) >= 4 else case_id
                text = f"{case_num} (身份证:{id_display})"
            else:
                text = f"{case_num} (身份证:无)"

            radio = QRadioButton(text)
            button_group.addButton(radio, i)
            layout.addWidget(radio)

        # 按钮区域
        btn_layout = QHBoxLayout()

        btn_ok = QPushButton("确定")
        btn_cancel = QPushButton("取消")

        btn_ok.clicked.connect(dialog.accept)
        btn_cancel.clicked.connect(dialog.reject)

        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

        dialog.setLayout(layout)

        if dialog.exec_() == QDialog.Accepted:
            selected_id = button_group.checkedId()
            if selected_id == 0:
                return "new"
            elif selected_id > 0:
                return cases[selected_id - 1]

        return None

    def link_to_existing_case(self, selected_case, injured_name):
        """关联到已有案件"""
        case_number = selected_case['case_number']
        year = selected_case['year']

        # 构建案件文件夹路径
        case_folder = os.path.join(os.path.dirname(__file__), str(year), case_number)

        # 检查是否已有本人笔录
        transcript_file = os.path.join(case_folder, f"{case_number}_笔录.docx")

        if os.path.exists(transcript_file):
            # 询问打开还是补充
            choice = self.show_transcript_exists_dialog(case_number)

            if choice == "open":
                os.startfile(transcript_file)
                return
            elif choice == "supplement":
                # 使用补充模板
                template_name = "本人补充笔录.docx"
            else:
                return  # 用户取消
        else:
            # 没有笔录，用普通模板
            template_name = "本人普通案件模板.docx"

        # 继续生成笔录
        self.generate_transcript(case_number, case_folder, template_name, injured_name)

    def show_transcript_exists_dialog(self, case_number):
        """显示已有笔录对话框"""
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QHBoxLayout

        dialog = QDialog(self)
        dialog.setWindowTitle("已有本人笔录")
        dialog.resize(300, 150)

        layout = QVBoxLayout()
        layout.addWidget(QLabel(f"案本 {case_number} 已有本人笔录文件"))

        btn_layout = QHBoxLayout()

        btn_open = QPushButton("打开现有笔录")
        btn_supplement = QPushButton("生成补充笔录")
        btn_cancel = QPushButton("取消")

        btn_open.clicked.connect(lambda: dialog.done(1))
        btn_supplement.clicked.connect(lambda: dialog.done(2))
        btn_cancel.clicked.connect(dialog.reject)

        btn_layout.addWidget(btn_open)
        btn_layout.addWidget(btn_supplement)
        btn_layout.addWidget(btn_cancel)

        layout.addLayout(btn_layout)
        dialog.setLayout(layout)

        result = dialog.exec_()

        if result == 1:
            return "open"
        elif result == 2:
            return "supplement"
        else:
            return "cancel"

    def create_new_case(self, injured_name):
        """创建新案件"""
        case_number = self.generate_case_number(injured_name)
        year_folder = self.get_current_year_folder()
        case_folder = os.path.join(year_folder, case_number)
        os.makedirs(case_folder, exist_ok=True)

        # 更新索引文件
        self.update_case_index(case_number, injured_name)

        # 生成笔录
        self.generate_transcript(case_number, case_folder, "本人普通案件模板.docx", injured_name)

    def collect_form_data(self):
        """收集当前表单所有数据"""
        id_card = self.lineEdit_id_card.text().strip()
        _, age, gender = self.calculate_id_info(id_card) if id_card else (None, None, None)

        # 获取人员类型作为前缀
        prefix = self.check_person_type()  # "本人" / "证人" / "法人"
        data = {
            '案本号': '',
            '受伤职工': self.lineEdit_injured_worker.text().strip(),
            '用人单位': self.comboBox_employer.currentText().strip(),
            '用工单位': self.comboBox_work_unit.currentText().strip(),
            '工作场所': self.comboBox_workplace.currentText().strip(),
            '人员类型': prefix,
            '案件类型': self.check_case_type(),
            '条例': self.comboBox_regulations.currentText(),
            '操作员': self.lineEdit_operator.text().strip(),
            '当前日期': datetime.now().strftime('%Y年%m月%d日'),
            '当前时间': datetime.now().strftime('%H时%M分'),
        }

        # 用变量作为前缀
        data[f'{prefix}姓名'] = self.lineEdit_name.text().strip()
        data[f'{prefix}性别'] = gender if gender else self.comboBox_gender.currentText()
        data[f'{prefix}年龄'] = str(age) if age else self.lineEdit_age.text().strip()
        data[f'{prefix}身份证号'] = id_card
        data[f'{prefix}身份证地址'] = self.lineEdit_id_address.text().strip()
        data[f'{prefix}现住址'] = self.lineEdit_current_address.text().strip()
        data[f'{prefix}电话'] = self.lineEdit_phone.text().strip()
        data[f'{prefix}岗位'] = self.lineEdit_position.text().strip()

        return data

    def generate_transcript(self, case_folder, template_name, data):
        """生成Word文档"""
        template_path = os.path.join(os.path.dirname(__file__), "templates", template_name)

        if not os.path.exists(template_path):
            self.statusBar().showMessage(f"模板不存在: {template_name}", 3000)
            return False

        from docx import Document
        doc = Document(template_path)

        # 替换占位符
        for paragraph in doc.paragraphs:
            text = paragraph.text
            for key, value in data.items():
                if f"{{{key}}}" in text:
                    text = text.replace(f"{{{key}}}", value)
            paragraph.text = text

        # 保存文件
        doc_file = os.path.join(case_folder, f"{data['案本号']}_笔录.docx")
        doc.save(doc_file)
        os.startfile(doc_file)

        self.statusBar().showMessage(f"笔录生成完成: {data['案本号']}", 3000)
        return True

    def update_case_index(self, case_number, person_name, data):
        """更新案件索引"""
        index_file = os.path.join(os.path.dirname(__file__), "cases_index.json")

        case_data = {
            'case_number': case_number,
            'person_name': person_name,
            'id_card': data['本人身份证号'],
            'case_type': data['案件类型'],
            'year': datetime.now().year,
            'folder_path': f"{datetime.now().year}/{case_number}",
            'created_date': datetime.now().strftime('%Y-%m-%d'),
            'person_info': {
                'name': data['本人姓名'],
                'gender': data['本人性别'],
                'age': data['本人年龄'],
                'phone': data['本人电话']
            }
        }

        try:
            import json

            if os.path.exists(index_file):
                with open(index_file, 'r', encoding='utf-8') as f:
                    index_data = json.load(f)
            else:
                index_data = {'cases': [], 'total_cases': 0, 'last_update': ''}

            index_data['cases'].append(case_data)
            index_data['total_cases'] = len(index_data['cases'])
            index_data['last_update'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            with open(index_file, 'w', encoding='utf-8') as f:
                json.dump(index_data, f, ensure_ascii=False, indent=2)

        except Exception as e:
            print(f"更新索引失败: {e}")

    def get_current_year_folder(self):
        """获取当前年份的cases文件夹"""
        current_year = datetime.now().year
        year_folder = os.path.join(os.path.dirname(__file__), str(current_year))
        os.makedirs(year_folder, exist_ok=True)
        return year_folder

    def generate_case_number(self, injured_name):
        """生成案本号：类型-姓名-序号（按年份）"""
        # 确定类型前缀
        case_type = self.check_case_type()

        if case_type == "普通案件":
            prefix = "GS"  # 普通工伤
        elif case_type == "个人案件":
            prefix = "GR"  # 个人申请
        elif case_type == "死亡案件":
            prefix = "GSW"  # 工亡案件（单位申请）
        elif case_type == "个人申请死亡案件":
            prefix = "GRW"  # 个人申请工亡
        else:
            prefix = "GS"

        # 使用年份文件夹
        year_folder = self.get_current_year_folder()

        # 计算下一个序号
        existing_numbers = []
        if os.path.exists(year_folder):
            for folder in os.listdir(year_folder):
                # 匹配格式：前缀-姓名-数字
                if folder.startswith(f"{prefix}-{injured_name}-"):
                    try:
                        num = int(folder.split('-')[-1])
                        existing_numbers.append(num)
                    except:
                        continue

        # 生成新序号
        if existing_numbers:
            next_num = max(existing_numbers) + 1
        else:
            next_num = 1

        case_number = f"{prefix}-{injured_name}-{next_num:03d}"
        return case_number

    def closeEvent(self, event):
        """窗口关闭时最后保存一次"""
        if self.checkBox_remember.isChecked():
            operator = self.lineEdit_operator.text().strip()
            api_url = self.lineEdit_api_url.text().strip()
            api_key = self.lineEdit_api_key.text().strip()
            self.config.save_config(operator, api_url, api_key, True)

        # 保存窗口大小位置
        settings = QSettings("WorkInjuryApp", "Window")
        settings.setValue("geometry", self.saveGeometry())

        event.accept()

    # 以下是测试程序，编程完成以后需要删除
    def keyPressEvent(self, event):
        """键盘按下事件"""
        if event.key() == Qt.Key_F2:  # 按 F2 键
            self.fill_test_data()
        elif event.key() == Qt.Key_F3:  # 按 F3 键填下一组
            self.fill_next_test_data()

    def fill_test_data(self):
        """填入测试数据（第一组）"""
        self.test_index = getattr(self, 'test_index', 0)
        self.fill_next_test_data()

    def fill_next_test_data(self):
        """填入下一组测试数据"""
        # 测试数据（5组）
        test_data = [
            {
                "name": "张三",
                "id_card": "410101199001011234",
                "id_address": "河南省郑州市中原区建设路1号",
                "current_address": "河南省郑州市金水区花园路2号院3号楼",
                "phone": "13800138000",
                "position": "车间主任"
            },
            {
                "name": "李四",
                "id_card": "410101199105022345",
                "id_address": "河南省洛阳市西工区中州路5号",
                "current_address": "河南省洛阳市涧西区南昌路8号院",
                "phone": "13900139001",
                "position": "技术员"
            },
            {
                "name": "王五",
                "id_card": "410101198206033456",
                "id_address": "河南省开封市龙亭区中山路10号",
                "current_address": "河南省开封市禹王台区五一路3号",
                "phone": "13700137002",
                "position": "安全员"
            }
        ]

        # 获取当前索引
        if not hasattr(self, 'test_index'):
            self.test_index = 0

        # 取当前组数据
        data = test_data[self.test_index]

        # 填入数据
        self.lineEdit_name.setText(data['name'])
        self.lineEdit_id_card.setText(data['id_card'])
        self.lineEdit_id_address.setText(data['id_address'])
        self.lineEdit_current_address.setText(data['current_address'])
        self.lineEdit_phone.setText(data['phone'])
        self.lineEdit_position.setText(data['position'])

        # 触发自动计算
        self.auto_calculate_id_info()

        # 如果是本人，触发自动填入受伤职工
        if self.check_person_type() == "本人":
            self.auto_fill_injured_worker()

        # 更新索引（循环）
        self.test_index = (self.test_index + 1) % len(test_data)

        self.statusBar().showMessage(f"已填入测试数据: {data['name']}", 2000)


def main():
    app = QApplication(sys.argv)

    app.setApplicationName("工伤案件管理系统")
    app.setOrganizationName("WorkInjuryApp")

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()