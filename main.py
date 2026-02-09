#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工伤案件管理系统 - 主程序
"""
import os
import sys
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.uic import loadUi
from PyQt5.QtCore import QSettings
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

        # 3.1设置ComboBox的自动完成和失去焦点保存功能
        self.setup_combobox_autosave()

        # 4. 加载保存的配置
        self.load_config()

        # 5. 连接信号
        self.checkBox_remember.stateChanged.connect(self.on_remember_changed)
        self.btn_generate_record.clicked.connect(self.on_generate_record)
        # 身份证号框失去焦点
        self.lineEdit_id_card.editingFinished.connect(self.auto_calculate_id_info)

        # 6. 根据记住状态更新界面
        self.update_ui()

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
            from openpyxl import load_workbook

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

        # 加载用人单位
        self.employer_list = self.load_excel_data(os.path.join(current_dir, "用人单位名称汇总.xlsx"))
        self.comboBox_employer.addItems(self.employer_list)

        # 加载用工单位
        self.work_unit_list = self.load_excel_data(os.path.join(current_dir, "用工单位名称汇总.xlsx"))
        self.comboBox_work_unit.addItems(self.work_unit_list)

        # 加载工作场所
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
        """生成笔录按钮点击事件"""
        # 1. 获取人员类型
        person_type = self.check_person_type()

        # 2. 如果是本人，检查姓名
        if person_type == "本人":
            name = self.lineEdit_name.text().strip()
            if name:
                self.lineEdit_injured_worker.setText(name)
            else:
                self.statusBar().showMessage("本人信息未填写", 3000)
                return

        # 3. 处理身份证信息
        id_card = self.lineEdit_id_card.text().strip()
        age = None
        gender = None
        if id_card:
            id_card, age, gender = self.calculate_id_info(id_card)
            if age and gender:
                self.lineEdit_age.setText(str(age))
                self.comboBox_gender.setCurrentText(gender)

        # 4. 获取拟用条例
        regulation_index = self.comboBox_regulations.currentIndex()
        regulation_mapping = {
            0: "第十四条第一款第一项",
            1: "第十四条第一款第二项",
            2: "第十四条第一款第三项",
            3: "第十四条第一款第四项",
            4: "第十四条第一款第五项",
            5: "第十四条第一款第六项",
            6: "第十五条第一款第一项"
        }
        regulation_key = regulation_mapping.get(regulation_index, "未知条例")

        # 5. 收集数据
        current_date = datetime.now().strftime('%Y年%m月%d日')
        current_time = datetime.now().strftime('%H时%M分')
        operator = self.lineEdit_operator.text().strip()

        data_dict = {
            '案件类型': self.check_case_type(),
            '人员类型': person_type,
            '条例': regulation_key,
            '当前日期': current_date,
            '当前时间': current_time,
            '姓名': self.lineEdit_name.text().strip(),
            '性别': gender if gender else '',
            '年龄': str(age) if age else '',
            '身份证号': id_card if id_card else '',
            '身份证地址': self.lineEdit_id_address.text().strip(),
            '现住址': self.lineEdit_current_address.text().strip(),
            '电话': self.lineEdit_phone.text().strip(),
            '岗位': self.lineEdit_position.text().strip(),
            '受伤职工': self.lineEdit_injured_worker.text().strip(),
            '用人单位': self.comboBox_employer.currentText().strip(),
            '用工单位': self.comboBox_work_unit.currentText().strip(),
            '工作场所': self.comboBox_workplace.currentText().strip(),
            '操作员': operator if operator else "未填写",
            '生成时间': f"{current_date} {current_time}"
        }

        # 6. 打开模板并填充
        template_path = os.path.join(os.path.dirname(__file__), "templates", "本人普通案件模板.docx")

        if os.path.exists(template_path):
            from docx import Document
            doc = Document(template_path)

            for paragraph in doc.paragraphs:
                text = paragraph.text
                for key, value in data_dict.items():
                    if f"{{{key}}}" in text:
                        text = text.replace(f"{{{key}}}", value)
                paragraph.text = text

            temp_file = f"temp_笔录.docx"
            doc.save(temp_file)
            os.startfile(temp_file)

            self.statusBar().showMessage("笔录生成完成", 3000)
        else:
            self.statusBar().showMessage("模板文件不存在", 3000)

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


def main():
    app = QApplication(sys.argv)

    app.setApplicationName("工伤案件管理系统")
    app.setOrganizationName("WorkInjuryApp")

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()