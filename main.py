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
from PyQt5.QtCore import Qt, QSettings
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

        # 4. 加载保存的配置
        self.load_config()

        # 5. 连接信号
        self.checkBox_remember.stateChanged.connect(self.on_remember_changed)
        self.btn_generate_record.clicked.connect(self.on_generate_record)

        # 6. 根据记住状态更新界面
        self.update_ui()

    def determine_word_template(self, person_type, case_type, regulation_key):
        """根据条件确定Word模板文件路径"""

        # 1. 检查是否是本人 + 普通案件（最简单的情况）
        if person_type == "本人" and case_type == "普通案件":
            # 先检查是否有对应的模板文件
            template_path = "templates/本人普通案件模板.docx"
            if os.path.exists(template_path):
                return template_path
            else:
                # 如果文件不存在，创建一个简单的提示
                self.statusBar().showMessage("模板文件不存在: " + template_path, 3000)
                return None

        # 2. 其他情况暂时返回默认模板
        else:
            # 可以在这里添加更多的模板判断逻辑
            default_template = "templates/通用模板.docx"
            if os.path.exists(default_template):
                self.statusBar().showMessage(f"使用通用模板: {person_type}+{case_type}", 3000)
                return default_template
            else:
                self.statusBar().showMessage("模板文件不存在，请检查templates目录", 3000)
                return None

    def load_excel_to_combobox(self):
        """从Excel文件加载数据到ComboBox"""
        # 获取当前目录
        current_dir = os.path.dirname(os.path.abspath(__file__))

        # 读取用人单位Excel
        try:
            employer_file = os.path.join(current_dir, "用人单位名称汇总.xlsx")
            if os.path.exists(employer_file):
                wb = load_workbook(employer_file)
                ws = wb.active
                # 读取第一列所有有数据的单元格
                for row in ws.iter_rows(min_row=1, max_col=1, values_only=True):
                    if row[0]:  # 检查单元格是否为空
                        self.comboBox_employer.addItem(str(row[0]))
        except Exception as e:
            print(f"读取用人单位Excel失败: {e}")

        # 读取用工单位Excel
        try:
            work_unit_file = os.path.join(current_dir, "用工单位名称汇总.xlsx")
            if os.path.exists(work_unit_file):
                wb = load_workbook(work_unit_file)
                ws = wb.active
                for row in ws.iter_rows(min_row=1, max_col=1, values_only=True):
                    if row[0]:
                        self.comboBox_work_unit.addItem(str(row[0]))
        except Exception as e:
            print(f"读取用工单位Excel失败: {e}")

        # 读取工作场所Excel
        try:
            workplace_file = os.path.join(current_dir, "工作场所名称汇总.xlsx")
            if os.path.exists(workplace_file):
                wb = load_workbook(workplace_file)
                ws = wb.active
                for row in ws.iter_rows(min_row=1, max_col=1, values_only=True):
                    if row[0]:
                        self.comboBox_workplace.addItem(str(row[0]))
        except Exception as e:
            print(f"读取工作场所Excel失败: {e}")

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

        # 设置样式（灰色背景表示不可编辑）
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

        # 计算性别（第17位，奇数为男，偶数为女）
        gender_num = int(id_card[16])
        gender = "男" if gender_num % 2 == 1 else "女"

        return id_card, age, gender

    def on_generate_record(self):
        """生成笔录按钮点击事件"""
        print("📝 生成笔录按钮被点击")

        # 检查案件类型
        case_type = self.check_case_type()
        print(f"案件类型: {case_type}")

        # 检查人员类型
        person_type = self.check_person_type()
        print(f"人员类型: {person_type}")

        # 如果是本人类型，检查姓名并复制到受伤职工
        if person_type == "本人":
            name = self.lineEdit_name.text().strip()
            print(f"本人姓名: '{name}'")

            if name:
                self.lineEdit_injured_worker.setText(name)
                print("✅ 姓名已复制到受伤职工")
            else:
                self.statusBar().showMessage("本人信息未填写", 3000)
                print("错误：本人信息未填写")
                return

        # 处理身份证信息
        id_card = self.lineEdit_id_card.text().strip()
        if id_card:
            id_card, age, gender = self.calculate_id_info(id_card)
            if age and gender:
                self.lineEdit_age.setText(str(age))
                self.comboBox_gender.setCurrentText(gender)

        # 获取其他基本信息
        id_address = self.lineEdit_id_address.text().strip()
        current_address = self.lineEdit_current_address.text().strip()
        phone = self.lineEdit_phone.text().strip()
        position = self.lineEdit_position.text().strip()

        # 获取拟用条例
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

        # 获取单位信息
        employer = self.comboBox_employer.currentText().strip()
        work_unit = self.comboBox_work_unit.currentText().strip()
        workplace = self.comboBox_workplace.currentText().strip()

        # ====== 打开Word模板 ======
        print(f"当前目录: {os.path.dirname(__file__)}")

        # 先测试直接打开
        template_path = os.path.join(os.path.dirname(__file__), "templates", "本人普通案件模板.docx")
        print(f"模板路径: {template_path}")
        print(f"模板存在: {os.path.exists(template_path)}")

        if os.path.exists(template_path):
            print(f"✅ 找到模板，准备打开Word文件")

            # 简单测试：直接打开
            os.startfile(template_path)  # Windows直接打开

            # 或者使用你的完整方法
            # self.open_word_template(template_path, {
            #     '案件类型': case_type,
            #     '人员类型': person_type,
            #     '条例': regulation_key,
            #     '姓名': self.lineEdit_name.text().strip(),
            #     '年龄': age if 'age' in locals() and age else '',
            #     '性别': gender if 'gender' in locals() and gender else '',
            #     '身份证号': id_card if id_card else '',
            #     '身份证地址': id_address,
            #     '现住址': current_address,
            #     '电话': phone,
            #     '岗位': position,
            #     '用人单位': employer,
            #     '用工单位': work_unit,
            #     '工作场所': workplace
            # })

            self.statusBar().showMessage("已打开Word文件", 3000)
        else:
            print(f"❌ 模板不存在")
            # 列出templates目录内容
            templates_dir = os.path.join(os.path.dirname(__file__), "templates")
            if os.path.exists(templates_dir):
                files = os.listdir(templates_dir)
                print(f"templates目录中的文件: {files}")
            else:
                print(f"templates目录不存在")

            self.statusBar().showMessage("模板文件不存在", 3000)

        # 显示结果
        result = f"案件类型: {case_type}, 人员类型: {person_type}, 条例: {regulation_key}"
        print(result)

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

    # 设置应用信息
    app.setApplicationName("工伤案件管理系统")
    app.setOrganizationName("WorkInjuryApp")

    # 创建并显示窗口
    window = MainWindow()
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()