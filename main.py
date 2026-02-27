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

        self.BASE_DIR = os.path.dirname(__file__)
        self.TEMPLATE_DIR = os.path.join(self.BASE_DIR, "templates")

        # 3. 加载Excel数据到ComboBox
        self.load_excel_to_combobox()

        # 3.1 设置ComboBox的自动完成和失去焦点保存功能
        self.setup_combobox_autosave()

        # 3.2 连接删除按钮
        self.setup_delete_buttons()

        # 3.3 连接新增按钮（放在这里）
        self.setup_document_buttons()

        # 4. 加载保存的配置
        self.load_config()

        # 5. 连接信号
        self.checkBox_remember.stateChanged.connect(self.on_remember_changed)
        self.btn_generate_record.clicked.connect(self.on_generate_record)
        # 身份证号框失去焦点
        self.lineEdit_id_card.editingFinished.connect(self.auto_calculate_id_info)
        # 用人单位失去焦点时触发
        self.comboBox_employer.lineEdit().editingFinished.connect(self.auto_fill_applicant)
        # 本人姓名失去焦点时触发
        self.lineEdit_name.editingFinished.connect(self.auto_fill_applicant)

        # 6. 根据记住状态更新界面
        self.update_ui()

        # 连接人员类型切换信号
        self.radio_self.toggled.connect(self.on_person_type_changed)
        self.radio_witness.toggled.connect(self.on_person_type_changed)
        self.radio_legal_entity.toggled.connect(self.on_person_type_changed)

        # 连接案件类型复选框的信号
        self.checkBox_personal.stateChanged.connect(self.on_case_type_changed)
        self.checkBox_death.stateChanged.connect(self.on_case_type_changed)

        # 用人单位改变时更新申请人
        self.comboBox_employer.currentTextChanged.connect(self.on_case_type_changed)
        # 本人姓名改变时更新申请人
        self.lineEdit_name.textChanged.connect(self.on_case_type_changed)

        self.current_case_number = None  # 当前使用的案本号
        self.current_folder_path = None  # 当前使用的文件夹路径

        self.label_current_case.setText("当前案本：无")

    def on_case_type_changed(self):
        """案件类型改变时，重新计算并更新申请人"""
        # 只处理本人类型
        if not self.radio_self.isChecked():
            return

        case_type = self.check_case_type()

        # 根据案件类型重新计算申请人
        if case_type in ["普通案件", "死亡案件"]:
            # 普通案件或死亡案件：申请人 = 用人单位
            employer = self.comboBox_employer.currentText().strip()
            if employer and employer != "用人单位名称汇总":
                self.lineEdit_applicant.setText(employer)
            else:
                self.lineEdit_applicant.clear()  # 如果没有用人单位，清空申请人

        elif case_type == "个人案件":
            # 个人案件：申请人 = 本人姓名
            name = self.lineEdit_name.text().strip()
            if name:
                self.lineEdit_applicant.setText(name)
            else:
                self.lineEdit_applicant.clear()

        elif case_type == "个人申请死亡案件":
            # 个人死亡案件：不清空，但可以给个提示
            if not self.lineEdit_applicant.text().strip():
                self.lineEdit_applicant.setPlaceholderText("请输入家属姓名")

    def auto_fill_applicant(self):
        """根据案件类型自动填充申请人"""
        # 只处理本人类型
        if not self.radio_self.isChecked():
            return

        case_type = self.check_case_type()

        # 普通案件或普通死亡案件：申请人 = 用人单位
        if case_type in ["普通案件", "死亡案件"]:
            employer = self.comboBox_employer.currentText().strip()
            if employer and employer != "用人单位名称汇总":
                self.lineEdit_applicant.setText(employer)

        # 个人案件：申请人 = 本人姓名
        elif case_type == "个人案件":
            name = self.lineEdit_name.text().strip()
            if name:
                self.lineEdit_applicant.setText(name)

        # 个人死亡案件：不清空，让用户手动输入
        # elif case_type == "个人申请死亡案件":
        #     pass  # 不做自动填充

    def setup_document_buttons(self):
        """连接各类文书生成按钮"""
        # 案件审批表
        self.btn_case_approval.clicked.connect(self.generate_case_approval)

        # 工伤告知书
        self.btn_injury_notice.clicked.connect(self.generate_injury_notice)

        # 谈话通知书
        self.btn_interview_notice.clicked.connect(self.generate_interview_notice)

    def generate_case_approval(self):
        """生成案件审批表"""
        if not self.current_case_number:
            QMessageBox.warning(self, "错误", "请先生成本人案本或关联已有案本")
            return

        try:
            import json
            from docx import Document
            from datetime import datetime

            # 读取索引文件
            index_file = os.path.join(self.BASE_DIR, "cases_index.json")
            with open(index_file, 'r', encoding='utf-8') as f:
                index_data = json.load(f)

            # 查找当前案本
            case_data = None
            for case in index_data.get('cases', []):
                if case['case_number'] == self.current_case_number:
                    case_data = case
                    break

            if not case_data:
                QMessageBox.warning(self, "错误", f"未找到案本 {self.current_case_number} 的数据")
                return

            # 获取模板
            template_path = os.path.join(self.BASE_DIR, "templates", "工伤案件审批表（模板）.docx")
            if not os.path.exists(template_path):
                QMessageBox.warning(self, "错误", "模板不存在")
                return

            # 创建文档对象
            doc = Document(template_path)

            # 获取受伤职工姓名
            injured_name = case_data.get('person_name', '')
            person_info = case_data.get('person_info', {})

            # 定义文本处理函数
            def process_self_intro(text):
                if not text:
                    return text
                return text[2:] if text.startswith("我是") else text

            def process_text(text):
                if not text:
                    return text
                text = text.replace("我们", "他们")
                text = text.replace("我", injured_name)
                return text

            def process_conclusion(text):
                if not text:
                    return text
                colon_index = text.find("：")
                if colon_index != -1:
                    return text[colon_index + 1:].strip()
                return text

            # 处理文本
            processed_self_intro = process_self_intro(person_info.get('自我介绍', ''))
            processed_injury = process_text(person_info.get('受伤经过', ''))
            processed_medical = process_text(person_info.get('就医情况', ''))
            processed_conclusion = process_conclusion(person_info.get('医疗结论', ''))

            # 准备替换数据
            replace_data = {
                '{案本号}': case_data.get('case_number', ''),
                '{受伤职工}': injured_name,
                '{申请人}': case_data.get('applicant', ''),
                '{性别}': person_info.get('gender', ''),
                '{年龄}': person_info.get('age', ''),
                '{身份证号}': person_info.get('id_card', ''),
                '{身份证地址}': person_info.get('address', ''),
                '{现住址}': person_info.get('current_address', ''),
                '{联系电话}': person_info.get('phone', ''),
                '{岗位}': person_info.get('position', ''),
                '{自我介绍}': processed_self_intro,
                '{受伤经过}': processed_injury,
                '{就医情况}': processed_medical,
                '{医疗结论}': processed_conclusion,
                '{用人单位}': case_data.get('employer', ''),
                '{用工单位}': case_data.get('work_unit', ''),
                '{工作场所}': case_data.get('workplace', ''),
                '{条例}': case_data.get('regulation', ''),
                '{案件类型}': case_data.get('case_type', ''),
                '{操作员}': case_data.get('operator', ''),
                '{当前日期}': datetime.now().strftime('%Y年%m月%d日'),
            }

            # 替换段落中的占位符
            for paragraph in doc.paragraphs:
                for key, value in replace_data.items():
                    if key in paragraph.text:
                        paragraph.text = paragraph.text.replace(key, value)

            # 替换表格中的占位符
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for key, value in replace_data.items():
                                if key in paragraph.text:
                                    paragraph.text = paragraph.text.replace(key, value)

            # 保存文件
            case_folder = os.path.join(self.BASE_DIR, case_data.get('folder_path', ''))
            filename = f"{self.current_case_number}_案件审批表.docx"
            filepath = os.path.join(case_folder, filename)

            doc.save(filepath)
            os.startfile(filepath)

            # ===== 新增：从审批表中提取申请时间和受理时间 =====
            self.extract_approval_times(doc, case_data, index_data, index_file)

            self.statusBar().showMessage(f"已生成案件审批表: {filename}", 3000)

        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成失败: {str(e)}")
            import traceback
            traceback.print_exc()

    def extract_approval_times(self, doc, case_data, index_data, index_file):
        """从审批表文档中提取申请时间和受理时间并保存到JSON"""
        try:
            import re
            from datetime import datetime

            application_time = ""
            acceptance_time = ""

            # 遍历表格查找申请时间和受理时间
            for table in doc.tables:
                for row in table.rows:
                    row_text = ""
                    for cell in row.cells:
                        row_text += cell.text.strip()

                    # 查找包含"申请时间"的行
                    if "申请时间" in row_text:
                        # 获取申请时间的值（通常在下一个单元格）
                        cells = row.cells
                        for i, cell in enumerate(cells):
                            if "申请时间" in cell.text:
                                if i + 1 < len(cells):
                                    time_value = cells[i + 1].text.strip()
                                    if time_value:
                                        application_time = self.format_date(time_value)
                                break

                    # 查找包含"受理时间"的行
                    if "受理时间" in row_text:
                        # 获取受理时间的值（通常在下一个单元格）
                        cells = row.cells
                        for i, cell in enumerate(cells):
                            if "受理时间" in cell.text:
                                if i + 1 < len(cells):
                                    time_value = cells[i + 1].text.strip()
                                    if time_value:
                                        acceptance_time = self.format_date(time_value)
                                break

            # 如果找到了时间，更新索引文件
            if application_time or acceptance_time:
                # 查找当前案件并更新时间
                for case in index_data.get('cases', []):
                    if case['case_number'] == case_data['case_number']:
                        if 'approval_info' not in case:
                            case['approval_info'] = {}

                        if application_time:
                            case['approval_info']['申请时间'] = application_time
                        if acceptance_time:
                            case['approval_info']['受理时间'] = acceptance_time

                        print(f"已保存申请时间: {application_time}, 受理时间: {acceptance_time}")
                        break

                # 保存更新后的索引文件
                import json
                import shutil
                temp_file = index_file + ".tmp"
                with open(temp_file, 'w', encoding='utf-8') as f:
                    json.dump(index_data, f, ensure_ascii=False, indent=2)
                shutil.move(temp_file, index_file)

        except Exception as e:
            print(f"提取申请/受理时间失败: {e}")
            import traceback
            traceback.print_exc()

    def format_date(self, date_str):
        """将日期字符串格式化为 xxxx年xx月xx日"""
        if not date_str:
            return ""

        # 移除所有非数字字符
        import re
        digits = re.sub(r'\D', '', date_str)

        # 如果是8位数字（如20260101）
        if len(digits) == 8:
            year = digits[0:4]
            month = digits[4:6].lstrip('0')  # 去掉前导零
            day = digits[6:8].lstrip('0')  # 去掉前导零
            return f"{year}年{month}月{day}日"

        # 如果是其他格式，尝试解析
        try:
            from datetime import datetime
            # 尝试常见格式
            for fmt in ["%Y%m%d", "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"]:
                try:
                    dt = datetime.strptime(date_str, fmt)
                    return dt.strftime("%Y年%m月%d日")
                except:
                    continue
        except:
            pass

        # 如果都无法解析，返回原字符串
        return date_str

    def generate_injury_notice(self):
        """生成工伤认定告知书"""
        if not self.current_case_number:
            QMessageBox.warning(self, "错误", "请先生成本人案本或关联已有案本")
            return

        try:
            import json
            from docx import Document
            from datetime import datetime
            import os

            # 读取索引文件
            index_file = os.path.join(self.BASE_DIR, "cases_index.json")
            with open(index_file, 'r', encoding='utf-8') as f:
                index_data = json.load(f)

            # 查找当前案本
            case_data = None
            for case in index_data.get('cases', []):
                if case['case_number'] == self.current_case_number:
                    case_data = case
                    break

            if not case_data:
                QMessageBox.warning(self, "错误", f"未找到案本 {self.current_case_number} 的数据")
                return

            # 查找审批表文件
            case_folder = os.path.join(self.BASE_DIR, case_data.get('folder_path', ''))
            approval_file = None
            for file in os.listdir(case_folder):
                if file.endswith('_案件审批表.docx'):
                    approval_file = os.path.join(case_folder, file)
                    break

            if not approval_file:
                QMessageBox.warning(self, "提示", "请先生成案件审批表")
                return

            # 从审批表读取数据
            approval_doc = Document(approval_file)

            申请时间 = ""
            受理时间 = ""
            综合情况 = ""
            医疗结论 = ""

            # 遍历表格查找数据
            for table in approval_doc.tables:
                for row in table.rows:
                    for i, cell in enumerate(row.cells):
                        text = cell.text
                        if "申请时间" in text and i + 1 < len(row.cells):
                            申请时间 = row.cells[i + 1].text.strip()
                        elif "受理时间" in text and i + 1 < len(row.cells):
                            受理时间 = row.cells[i + 1].text.strip()
                        elif "受伤经过" in text and i + 1 < len(row.cells):
                            综合情况 = row.cells[i + 1].text.strip()
                        elif "医疗诊断" in text and i + 1 < len(row.cells):
                            医疗结论 = row.cells[i + 1].text.strip()

            # 格式化时间（20260101 -> 2026年1月1日）
            if len(申请时间) == 8 and 申请时间.isdigit():
                申请时间 = f"{申请时间[:4]}年{int(申请时间[4:6])}月{int(申请时间[6:])}日"
            if len(受理时间) == 8 and 受理时间.isdigit():
                受理时间 = f"{受理时间[:4]}年{int(受理时间[4:6])}月{int(受理时间[6:])}日"

            if not 申请时间 or not 受理时间:
                QMessageBox.warning(self, "提示", "审批表中未找到申请时间或受理时间")
                return

            # 获取模板
            template_path = os.path.join(self.TEMPLATE_DIR, "工伤认定告知书（模板）.docx")
            if not os.path.exists(template_path):
                QMessageBox.warning(self, "错误", "工伤认定告知书模板不存在")
                return

            doc = Document(template_path)

            # 准备替换数据
            replace_data = {
                '{用人单位}': case_data.get('employer', ''),
                '{申请人}': case_data.get('applicant', ''),
                '{申请时间}': 申请时间,
                '{受理时间}': 受理时间,
                '{受伤职工}': case_data.get('person_name', ''),
                '{综合情况}': 综合情况,
                '{医疗结论}': 医疗结论,
                '{条例}': case_data.get('regulation', ''),
                '{当前时期}': datetime.now().strftime('%Y年%m月%d日'),
            }

            # 替换占位符
            for para in doc.paragraphs:
                for key, value in replace_data.items():
                    if key in para.text:
                        para.text = para.text.replace(key, value)

            # 保存文件
            filename = f"{self.current_case_number}_工伤认定告知书.docx"
            filepath = os.path.join(case_folder, filename)
            doc.save(filepath)
            os.startfile(filepath)

            self.statusBar().showMessage(f"已生成工伤认定告知书: {filename}", 3000)

        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成工伤认定告知书失败: {str(e)}")
            import traceback
            traceback.print_exc()

    def generate_interview_notice(self):
        """生成接受谈话通知书"""
        if not self.current_case_number:
            QMessageBox.warning(self, "错误", "请先生成本人案本或关联已有案本")
            return

        try:
            import json
            from docx import Document
            from datetime import datetime
            import os

            # 读取索引文件
            index_file = os.path.join(self.BASE_DIR, "cases_index.json")
            with open(index_file, 'r', encoding='utf-8') as f:
                index_data = json.load(f)

            # 查找当前案本
            case_data = None
            for case in index_data.get('cases', []):
                if case['case_number'] == self.current_case_number:
                    case_data = case
                    break

            if not case_data:
                QMessageBox.warning(self, "错误", f"未找到案本 {self.current_case_number} 的数据")
                return

            # 查找审批表文件
            case_folder = os.path.join(self.BASE_DIR, case_data.get('folder_path', ''))
            approval_file = None
            for file in os.listdir(case_folder):
                if file.endswith('_案件审批表.docx'):
                    approval_file = os.path.join(case_folder, file)
                    break

            if not approval_file:
                QMessageBox.warning(self, "提示", "请先生成案件审批表")
                return

            # 从审批表读取数据
            approval_doc = Document(approval_file)

            申请时间 = ""
            受理时间 = ""
            综合情况 = ""
            医疗结论 = ""

            # 遍历表格查找数据
            for table in approval_doc.tables:
                for row in table.rows:
                    for i, cell in enumerate(row.cells):
                        text = cell.text
                        if "申请时间" in text and i + 1 < len(row.cells):
                            申请时间 = row.cells[i + 1].text.strip()
                        elif "受理时间" in text and i + 1 < len(row.cells):
                            受理时间 = row.cells[i + 1].text.strip()
                        elif "受伤经过" in text and i + 1 < len(row.cells):
                            综合情况 = row.cells[i + 1].text.strip()
                        elif "医疗诊断" in text and i + 1 < len(row.cells):
                            医疗结论 = row.cells[i + 1].text.strip()

            # 格式化时间（20260101 -> 2026年1月1日）
            if len(申请时间) == 8 and 申请时间.isdigit():
                申请时间 = f"{申请时间[:4]}年{int(申请时间[4:6])}月{int(申请时间[6:])}日"
            if len(受理时间) == 8 and 受理时间.isdigit():
                受理时间 = f"{受理时间[:4]}年{int(受理时间[4:6])}月{int(受理时间[6:])}日"

            if not 申请时间 or not 受理时间:
                QMessageBox.warning(self, "提示", "审批表中未找到申请时间或受理时间")
                return

            # 生成通知书
            template_path = os.path.join(self.TEMPLATE_DIR, "接受谈话通知书（模板）.docx")
            doc = Document(template_path)

            replace_data = {
                '{用人单位}': case_data.get('employer', ''),
                '{申请人}': case_data.get('applicant', ''),
                '{申请时间}': 申请时间,
                '{受理时间}': 受理时间,
                '{本人姓名}': case_data.get('person_name', ''),
                '{本人身份证}': case_data.get('person_info', {}).get('id_card', ''),
                '{综合情况}': 综合情况,
                '{医疗结论}': 医疗结论,
                '{当前时期}': datetime.now().strftime('%Y年%m月%d日'),
            }

            # 替换占位符
            for para in doc.paragraphs:
                for key, value in replace_data.items():
                    if key in para.text:
                        para.text = para.text.replace(key, value)

            # 保存文件
            filename = f"{self.current_case_number}_接受谈话通知书.docx"
            filepath = os.path.join(case_folder, filename)
            doc.save(filepath)
            os.startfile(filepath)

            self.statusBar().showMessage(f"已生成接受谈话通知书: {filename}", 3000)

        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成失败: {str(e)}")

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
                else:
                    self.statusBar().showMessage("未在Excel中找到该项目", 3000)
            else:
                self.statusBar().showMessage("Excel文件不存在", 3000)

            # 4. 清空当前选择
            combobox.setCurrentIndex(-1)
            combobox.setCurrentText("")

        except Exception as e:
            self.statusBar().showMessage(f"删除失败: {str(e)}", 3000)

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
        # 检查个人死亡案件的申请人
        if self.check_case_type() == "个人申请死亡案件" and self.radio_self.isChecked():
            if not self.lineEdit_applicant.text().strip():
                QMessageBox.warning(self, "提示", "请填写申请人信息（家属姓名）")
                return

        # 1. 收集数据
        data = self.collect_form_data()
        if not data:  # 如果收集数据失败（比如申请人未填写）
            return

        # 2. 根据人员类型分流
        if data['人员类型'] == "本人":
            self.handle_person_case(data)
        elif data['人员类型'] == "证人":
            self.handle_witness_case(data)
        elif data['人员类型'] == "法人":
            self.handle_legal_case(data)

    def handle_person_case(self, data):
        # 1. 生成自我介绍
        description = self.generate_description(data)
        print(description)

        # 2. 根据案件类型生成问答句（仅用于显示）
        case_type = data['案件类型']
        if case_type != "普通案件":
            questions = self.generate_case_questions(case_type, data)
            if questions:
                print(f"\n=== {case_type}问答句 ===")
                for q in questions:
                    print(q)

        # ========== 检查是否已有案本 ==========
        import json
        index_file = os.path.join(self.BASE_DIR, "cases_index.json")

        if os.path.exists(index_file):
            with open(index_file, 'r', encoding='utf-8') as f:
                index_data = json.load(f)

            # 搜索同名案件
            same_person_cases = []
            for case in index_data.get('cases', []):
                if case['person_name'] == data['受伤职工']:
                    same_person_cases.append(case)

            if same_person_cases:
                selected_case = self.show_case_selection_dialog(
                    data['受伤职工'],
                    same_person_cases,
                    data['本人身份证号']
                )

                if selected_case == "new":
                    pass
                elif selected_case:
                    self.current_case_number = selected_case['case_number']
                    self.current_folder_path = selected_case['folder_path']

                    # 更新案本号显示（只改文本）
                    self.label_current_case.setText(f"当前案本：{self.current_case_number}")

                    person_info = selected_case.get('person_info', {})
                    self.lineEdit_name.setText(person_info.get('name', ''))
                    self.lineEdit_id_card.setText(selected_case.get('id_card', ''))
                    self.lineEdit_phone.setText(person_info.get('phone', ''))
                    if selected_case.get('id_card'):
                        self.auto_calculate_id_info()

                    self.statusBar().showMessage(f"已关联案本: {selected_case['case_number']}", 3000)

                    if "死亡" in case_type:
                        QMessageBox.information(self, "提示",
                                                f"已关联死亡职工案本：{selected_case['case_number']}\n\n请继续输入证人笔录信息",
                                                QMessageBox.Ok)
                    return
                else:
                    return

        # ========== 新建案件 ==========
        case_number = self.generate_case_number(data['受伤职工'])
        data['案本号'] = case_number
        year_folder = self.get_current_year_folder()
        case_folder = os.path.join(year_folder, case_number)
        os.makedirs(case_folder, exist_ok=True)

        # 保存当前使用的案本信息
        self.current_case_number = case_number
        self.current_folder_path = f"{datetime.now().year}/{case_number}"

        # 更新案本号显示（只改文本）
        self.label_current_case.setText(f"当前案本：{self.current_case_number}")

        # 在数据中添加自我介绍
        data['自我介绍'] = description

        # 预留三个字段，等待后续提取
        data['受伤经过'] = ''
        data['就医情况'] = ''
        data['医疗结论'] = ''

        # 更新索引
        self.update_case_index(case_number, data['受伤职工'], data)

        # 判断是否为死亡案件
        if "死亡" in case_type:
            QMessageBox.information(self, "提示",
                                    f"死亡职工信息已保存\n案本号：{case_number}\n\n请继续输入证人笔录信息",
                                    QMessageBox.Ok)
        else:
            template_name = self.get_template_name(data)
            self.generate_transcript(case_folder, template_name, data)

    def handle_witness_case(self, data):
        # 1. 生成自我介绍并保存到data
        description = self.generate_description(data)
        print(description)
        data['自我介绍'] = description  # ← 添加这行，保存自我介绍

        # 2. 根据案件类型生成问答句
        case_type = data['案件类型']
        if case_type != "普通案件":
            questions = self.generate_case_questions(case_type, data)
            if questions:
                print(f"\n=== {case_type}问答句（证人版）===")
                for q in questions:
                    print(q)

        # ===== 强制检查是否有当前案本号 =====
        if not self.current_case_number:
            QMessageBox.warning(self, "错误", "请先生成本人案本或关联已有案本")
            return

        witness_name = data.get('证人姓名', '')
        # 直接使用保存的案本信息
        case_number = self.current_case_number
        data['案本号'] = case_number

        year_folder = self.get_current_year_folder()
        case_folder = os.path.join(year_folder, case_number)

        # 确保文件夹存在
        if not os.path.exists(case_folder):
            os.makedirs(case_folder, exist_ok=True)

        # ===== 查找该文件夹下所有证人笔录 =====
        witness_files = []
        if os.path.exists(case_folder):
            for file in os.listdir(case_folder):
                if file.endswith('.docx') and '证人' in file:
                    witness_files.append(file)

        if not witness_files:
            # ========== 情况2.1：没有证人笔录，直接生成第一个 ==========
            template_name = self.get_template_name(data)
            self.create_witness_transcript(case_folder, data, witness_number=1, template_name=template_name)
            return

        # ========== 情况2.2：已有证人笔录，检查是否同一证人 ==========
        import re
        witness_exists = False
        max_number = 0
        existing_file = None

        for file in witness_files:
            # 文件名格式：受伤职工姓名_证人XX_证人姓名.docx
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
                if existing_file and os.path.exists(existing_file):
                    os.startfile(existing_file)
                    self.statusBar().showMessage(f"已打开证人笔录", 3000)
                else:
                    QMessageBox.warning(self, "错误", "找不到证人笔录文件")
            else:
                # 新建另一份（编号+1）
                template_name = self.get_template_name(data)
                self.create_witness_transcript(case_folder, data, witness_number=max_number + 1,
                                               template_name=template_name)
        else:
            # ========== 情况2.2.2：新证人，直接生成 ==========
            template_name = self.get_template_name(data)
            self.create_witness_transcript(case_folder, data, witness_number=max_number + 1,
                                           template_name=template_name)

    def create_witness_transcript(self, case_folder, data, witness_number, template_name):
        """生成证人笔录"""
        return self.generate_transcript_unified(
            case_folder=case_folder,
            data=data,
            template_name=template_name,
            file_prefix="证人",
            person_type="证人",
            person_name=data.get('证人姓名', '')
        )

    def create_legal_transcript(self, case_folder, data, legal_number, template_name):
        """生成法人笔录"""
        return self.generate_transcript_unified(
            case_folder=case_folder,
            data=data,
            template_name=template_name,
            file_prefix="法人",
            person_type="法人",
            person_name=data.get('法人姓名', '')
        )

    def handle_legal_case(self, data):
        # 1. 生成自我介绍并保存到data
        description = self.generate_description(data)
        print(description)
        data['自我介绍'] = description  # ← 添加这行，保存自我介绍

        # 2. 根据案件类型生成问答句
        case_type = data['案件类型']
        if case_type != "普通案件":
            questions = self.generate_case_questions(case_type, data)
            if questions:
                print(f"\n=== {case_type}问答句（法人版）===")
                for q in questions:
                    print(q)

        # ===== 强制检查是否有当前案本号 =====
        if not self.current_case_number:
            QMessageBox.warning(self, "错误", "请先生成本人案本或关联已有案本")
            return

        try:
            injured_name = data['受伤职工']
            legal_name = data['法人姓名']

            # 直接使用保存的案本信息
            case_number = self.current_case_number
            data['案本号'] = case_number

            year_folder = self.get_current_year_folder()
            case_folder = os.path.join(year_folder, case_number)

            # 确保文件夹存在
            if not os.path.exists(case_folder):
                os.makedirs(case_folder, exist_ok=True)

            # ===== 查找该文件夹下所有法人笔录 =====
            legal_files = []
            if os.path.exists(case_folder):
                for file in os.listdir(case_folder):
                    if file.endswith('.docx') and '法人' in file:
                        legal_files.append(file)

            if not legal_files:
                # ===== 没有法人笔录，直接生成第一个 =====
                template_name = self.get_template_name(data)
                self.create_legal_transcript(case_folder, data, legal_number=1, template_name=template_name)
                return

            # ===== 已有法人笔录，检查是否同一法人 =====
            import re
            legal_exists = False
            max_number = 0
            existing_file = None

            for file in legal_files:
                # 文件名格式：受伤职工姓名_法人XX_法人姓名.docx
                match = re.search(r'法人(\d+)_(.+?)\.docx', file)
                if match:
                    num = int(match.group(1))
                    existing_legal_name = match.group(2)
                    max_number = max(max_number, num)

                    # 如果法人姓名相同，视为同一法人
                    if existing_legal_name == legal_name:
                        legal_exists = True
                        existing_file = os.path.join(case_folder, file)

            if legal_exists:
                # ===== 同一法人，询问关联或新建 =====
                reply = QMessageBox.question(
                    self, '法人已存在',
                    f'法人 "{legal_name}" 已有笔录\n是否打开？\n\n选“是”=打开\n选“否”=新建另一份',
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )

                if reply == QMessageBox.Yes:
                    if existing_file and os.path.exists(existing_file):
                        os.startfile(existing_file)
                        self.statusBar().showMessage(f"已打开法人笔录", 3000)
                    else:
                        QMessageBox.warning(self, "错误", "找不到法人笔录文件")
                else:
                    # 新建另一份（编号+1）
                    template_name = self.get_template_name(data)
                    self.create_legal_transcript(case_folder, data, legal_number=max_number + 1,
                                                 template_name=template_name)
            else:
                # ===== 新法人，直接生成 =====
                template_name = self.get_template_name(data)
                self.create_legal_transcript(case_folder, data, legal_number=max_number + 1,
                                             template_name=template_name)

        except Exception as e:
            import traceback
            print("=" * 50)
            print("法人案件处理出错:")
            traceback.print_exc()
            print("=" * 50)
            self.statusBar().showMessage(f"错误: {str(e)}", 3000)

    def generate_transcript_unified(self, case_folder, data, template_name, file_prefix, person_type, person_name):
        """
        统一的笔录生成方法（使用占位符替换）
        """
        try:
            from docx import Document
            import threading
            import time

            template_path = os.path.join(self.TEMPLATE_DIR, template_name)
            if not os.path.exists(template_path):
                self.statusBar().showMessage(f"模板不存在: {template_name}", 3000)
                return None

            doc = Document(template_path)

            # 1. 准备所有替换数据
            placeholders = {
                # 基本信息
                '受伤职工': data.get('受伤职工', ''),
                '用人单位': data.get('用人单位', ''),
                '用工单位': data.get('用工单位', ''),
                '工作场所': data.get('工作场所', ''),
                '操作员': data.get('操作员', ''),
                '当前日期': datetime.now().strftime('%Y年%m月%d日'),
                '当前时间': datetime.now().strftime('%H时%M分'),

                # 自我介绍
                '自我介绍': self.generate_description(data),

                # 人员特定信息
                f'{person_type}姓名': person_name,
                f'{person_type}性别': data.get(f'{person_type}性别', ''),
                f'{person_type}年龄': data.get(f'{person_type}年龄', ''),
                f'{person_type}身份证': data.get(f'{person_type}身份证号', ''),
                f'{person_type}身份证地址': data.get(f'{person_type}身份证地址', ''),
                f'{person_type}电话': data.get(f'{person_type}电话', ''),
                f'{person_type}岗位': data.get(f'{person_type}岗位', ''),
            }

            # 2. 替换所有占位符
            for paragraph in doc.paragraphs:
                text = paragraph.text
                for key, value in placeholders.items():
                    placeholder = f"{{{key}}}"
                    if placeholder in text:
                        text = text.replace(placeholder, str(value))
                paragraph.text = text

            # 3. 替换表格中的占位符
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            text = paragraph.text
                            for key, value in placeholders.items():
                                placeholder = f"{{{key}}}"
                                if placeholder in text:
                                    text = text.replace(placeholder, str(value))
                            paragraph.text = text

            # 4. 插入案件问答句
            doc = self.add_questions_to_doc(doc, data)

            # 5. 生成文件名并保存
            injured_name = data.get('受伤职工', '')
            import re
            max_num = 0
            if os.path.exists(case_folder):
                for file in os.listdir(case_folder):
                    pattern = rf'{injured_name}_{file_prefix}(\d+)_'
                    match = re.search(pattern, file)
                    if match:
                        num = int(match.group(1))
                        max_num = max(max_num, num)

            next_num = max_num + 1
            filename = f"{injured_name}_{file_prefix}{next_num:02d}_{person_name}.docx"
            filepath = os.path.join(case_folder, filename)

            doc.save(filepath)

            # 打开文件
            os.startfile(filepath)

            self.statusBar().showMessage(f"{person_type}笔录已生成: {filename}", 3000)

            # 更新索引
            self.update_case_index(data['案本号'], data['受伤职工'], data)

            # ===== 新增：如果是本人笔录，启动后台线程监控文件关闭 =====
            if person_type == "本人":
                def wait_for_file_close():
                    """等待文件关闭后自动提取信息"""
                    time.sleep(2)  # 给Word一点启动时间

                    # 等待文件关闭（尝试以写入模式打开，如果失败说明文件还在使用中）
                    file_closed = False
                    while not file_closed:
                        try:
                            # 尝试以追加模式打开，如果成功说明文件已关闭
                            with open(filepath, 'a'):
                                pass
                            file_closed = True
                        except:
                            # 文件还在使用中，等待1秒后重试
                            time.sleep(1)

                    # 文件已关闭，提取信息
                    self.extract_person_info_from_doc(filepath, data['案本号'])

                # 启动后台线程
                threading.Thread(target=wait_for_file_close, daemon=True).start()

            return filepath

        except Exception as e:
            self.statusBar().showMessage(f"生成失败: {str(e)}", 3000)
            import traceback
            traceback.print_exc()
            return None

    def generate_case_questions(self, case_type, data):
        """根据案件类型生成对应的问答句"""

        if case_type == "个人案件":
            return [
                "问：你是个人申请工伤认定吗？",
                "答：是的，我是个人申请。",
                "问：单位为什么没有为你申请？",
                "答：单位说让我自己申请。",
                # ... 更多个人案件专用问题
            ]

        elif case_type == "死亡案件":
            return [
                "问：你是死亡职工的家属吗？",
                "答：是的，我是他的家属。",
                "问：死亡时间和原因是什么？",
                "答：...",
                # ... 更多死亡案件专用问题
            ]

        elif case_type == "个人申请死亡案件":
            return [
                "问：你是以家属身份个人申请工亡吗？",
                "答：是的。",
                "问：单位没有为死者申报吗？",
                "答：没有。",
                # ... 综合问题
            ]

        else:  # 普通案件
            return []  # 返回空列表

    def search_same_name_cases(self, name, id_card):
        """搜索同名案件"""
        cases = []

        # 读取索引文件
        index_file = os.path.join(self.BASE_DIR, "cases_index.json")

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

    def generate_description(self, data):
        """根据人员类型和单位情况生成描述语句"""
        person_type = data['人员类型']  # 本人/证人/法人
        has_employer = bool(data.get('用人单位', ''))
        has_work_unit = bool(data.get('用工单位', ''))
        has_workplace = bool(data.get('工作场所', ''))

        # 获取姓名（根据人员类型不同，键名不同）
        if person_type == "本人":
            name = data.get('本人姓名', '')
        elif person_type == "证人":
            name = data.get('证人姓名', '')
        else:  # 法人
            name = data.get('法人姓名', '')

        employer = data.get('用人单位', '')
        work_unit = data.get('用工单位', '')
        workplace = data.get('工作场所', '')
        position = data.get(f'{person_type}岗位', '')

        # 生成描述语句
        if has_employer and has_work_unit and has_workplace:
            description = f"我是{name}，系{employer}的职工，被指派到{work_unit}的{workplace}工作。从事{position}工作。"
        elif has_employer and has_work_unit:
            description = f"我是{name}，系{employer}的职工，被指派到{work_unit}工作。从事{position}工作。"
        elif has_employer and has_workplace:
            description = f"我是{name}，系{employer}的职工，被指派到{workplace}工作。从事{position}工作。"
        elif has_employer:
            description = f"我是{name}，系{employer}的职工。从事{position}工作。"
        else:
            description = f"我是{name}。从事{position}工作。"

        return description

    def show_case_selection_dialog(self, name, cases, id_card):
        """显示案件选择对话框（支持红色显示身份证不同的案件）"""
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QRadioButton, QButtonGroup, QPushButton, QHBoxLayout
        from PyQt5.QtGui import QColor

        dialog = QDialog(self)
        dialog.setWindowTitle("发现同名案件")
        dialog.resize(450, 350)

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
            # 从 person_info 里取身份证
            case_id = case.get('person_info', {}).get('id_card', '')

            # 创建水平布局放单选按钮和文本
            h_layout = QHBoxLayout()
            radio = QRadioButton()
            button_group.addButton(radio, i)
            h_layout.addWidget(radio)

            # 创建显示文本的标签
            if case_id:
                id_display = case_id[-4:] if len(case_id) >= 4 else case_id
                text = f"{case_num} (身份证:{id_display})"
            else:
                text = f"{case_num} (身份证:无)"

            label = QLabel(text)

            # 如果输入的身份证和案件的身份证不同，设为红色
            if id_card and case_id and id_card != case_id:
                label.setStyleSheet("color: red;")

            h_layout.addWidget(label)
            h_layout.addStretch()
            layout.addLayout(h_layout)

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

    def collect_form_data(self):
        """收集当前表单所有数据"""
        id_card = self.lineEdit_id_card.text().strip()
        _, age, gender = self.calculate_id_info(id_card) if id_card else (None, None, None)

        # 获取人员类型作为前缀
        prefix = self.check_person_type()

        # 获取条例选择
        regulation_text = self.comboBox_regulations.currentText().strip()

        # 简单的判断：如果ComboBox显示的是文件名，就视为空
        employer = self.comboBox_employer.currentText().strip()
        if employer == "用人单位名称汇总":
            employer = ""

        work_unit = self.comboBox_work_unit.currentText().strip()
        if work_unit == "用工单位名称汇总":
            work_unit = ""

        workplace = self.comboBox_workplace.currentText().strip()
        if workplace == "工作场所名称汇总":
            workplace = ""

        applicant = self.lineEdit_applicant.text().strip()

        # 获取受伤职工（从本人姓名获取）
        injured_worker = self.lineEdit_name.text().strip()  # ✅ 改为从本人姓名获取

        # 个人死亡案件检查
        if self.check_case_type() == "个人申请死亡案件" and self.radio_self.isChecked():
            if not applicant:
                return None  # 返回None表示数据不完整

        data = {
            '案本号': '',
            '受伤职工': injured_worker,  # 从本人姓名获取
            '申请人': applicant,
            '用人单位': employer,
            '用工单位': work_unit,
            '工作场所': workplace,
            '人员类型': prefix,
            '案件类型': self.check_case_type(),
            '条例': regulation_text,
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
        """生成本人笔录"""
        return self.generate_transcript_unified(
            case_folder=case_folder,
            data=data,
            template_name=template_name,
            file_prefix="本人",
            person_type="本人",
            person_name=data.get('本人姓名', '')
        )

    def extract_person_info_from_doc(self, doc_file, case_number):
        """从Word文档中提取本人关键信息"""
        try:
            from docx import Document
            import json

            if not os.path.exists(doc_file):
                return

            doc = Document(doc_file)

            # 要搜索的关键词
            question_keywords = {
                '受伤经过': ['什么工作原因', '事故发生', '具体经过'],
                '就医情况': ['受伤后', '哪个医院', '是谁送你'],
                '医疗结论': ['此次受伤', '医院对你', '医疗结论']
            }

            extracted_info = {}

            # 遍历所有段落，查找问题
            for i, paragraph in enumerate(doc.paragraphs):
                text = paragraph.text.strip()
                if not text:
                    continue

                for info_key, keywords in question_keywords.items():
                    if info_key in extracted_info:
                        continue

                    match_count = 0
                    for keyword in keywords:
                        if keyword in text:
                            match_count += 1

                    if match_count >= 2:
                        if i + 1 < len(doc.paragraphs):
                            answer = doc.paragraphs[i + 1].text.strip()
                            if answer.startswith('答：'):
                                answer = answer[2:].strip()
                            elif answer.startswith('答:'):
                                answer = answer[1:].strip()

                            extracted_info[info_key] = answer
                            break

            if extracted_info:
                self.update_extracted_info_in_index(case_number, extracted_info)

        except Exception as e:
            print(f"提取信息失败: {e}")

    def update_extracted_info_in_index(self, case_number, extracted_info):
        """只更新受伤经过、就医情况、医疗结论三个字段"""
        index_file = os.path.join(self.BASE_DIR, "cases_index.json")

        try:
            import json
            with open(index_file, 'r', encoding='utf-8') as f:
                index_data = json.load(f)

            for case in index_data.get('cases', []):
                if case['case_number'] == case_number:
                    if 'person_info' in case:
                        case['person_info']['受伤经过'] = extracted_info.get('受伤经过', '')
                        case['person_info']['就医情况'] = extracted_info.get('就医情况', '')
                        case['person_info']['医疗结论'] = extracted_info.get('医疗结论', '')
                    break

            with open(index_file, 'w', encoding='utf-8') as f:
                json.dump(index_data, f, ensure_ascii=False, indent=2)

        except Exception as e:
            print(f"更新提取信息失败: {e}")

    def update_person_info_in_index(self, case_number, extracted_info):
        """在索引文件中更新本人的额外信息"""
        index_file = os.path.join(self.BASE_DIR, "cases_index.json")

        try:
            import json

            if not os.path.exists(index_file):
                self.statusBar().showMessage("索引文件不存在", 3000)
                return

            with open(index_file, 'r', encoding='utf-8') as f:
                index_data = json.load(f)

            # 查找对应的案本
            updated = False
            for case in index_data.get('cases', []):
                if case['case_number'] == case_number:
                    # 确保person_info存在
                    if 'person_info' not in case:
                        case['person_info'] = {}

                    # 添加提取的信息
                    case['person_info']['受伤经过'] = extracted_info.get('受伤经过', '')
                    case['person_info']['就医情况'] = extracted_info.get('就医情况', '')
                    case['person_info']['医疗结论'] = extracted_info.get('医疗结论', '')

                    updated = True
                    break

            if updated:
                # 保存更新后的索引
                with open(index_file, 'w', encoding='utf-8') as f:
                    json.dump(index_data, f, ensure_ascii=False, indent=2)
                print(f"已更新案本 {case_number} 的本人信息")
            else:
                print(f"未找到案本 {case_number}")

        except Exception as e:
            print(f"更新索引失败: {e}")

    def add_questions_to_doc(self, doc, data):
        """将案件类型问答句添加到文档中"""
        case_type = data['案件类型']
        if case_type != "普通案件":
            questions = self.generate_case_questions(case_type, data)
            if questions:
                doc.add_paragraph()  # 空行
                for q in questions:
                    doc.add_paragraph(q)
        return doc

    def update_case_index(self, case_number, person_name, data):
        """更新案件索引文件（合并数据，避免覆盖）"""
        index_file = os.path.join(self.BASE_DIR, "cases_index.json")

        try:
            import json
            import shutil
            from datetime import datetime

            # 读取现有索引
            if os.path.exists(index_file):
                with open(index_file, 'r', encoding='utf-8') as f:
                    index_data = json.load(f)
            else:
                index_data = {'cases': [], 'total_cases': 0, 'last_update': ''}

            # 查找现有案件
            found = False
            for i, existing_case in enumerate(index_data['cases']):
                if existing_case['case_number'] == case_number:
                    # 获取现有的 person_info
                    existing_person_info = existing_case.get('person_info', {})

                    # 构建新的 person_info（保留旧数据，用新数据覆盖）
                    new_person_info = {
                        'name': data.get('本人姓名', existing_person_info.get('name', '')),
                        'gender': data.get('本人性别', existing_person_info.get('gender', '')),
                        'age': data.get('本人年龄', existing_person_info.get('age', '')),
                        'phone': data.get('本人电话', existing_person_info.get('phone', '')),
                        'id_card': data.get('本人身份证号', existing_person_info.get('id_card', '')),
                        'address': data.get('本人身份证地址', existing_person_info.get('address', '')),
                        'current_address': data.get('本人现住址', existing_person_info.get('current_address', '')),
                        'position': data.get('本人岗位', existing_person_info.get('position', '')),
                        '自我介绍': data.get('自我介绍', existing_person_info.get('自我介绍', '')),
                        '受伤经过': data.get('受伤经过', existing_person_info.get('受伤经过', '')),
                        '就医情况': data.get('就医情况', existing_person_info.get('就医情况', '')),
                        '医疗结论': data.get('医疗结论', existing_person_info.get('医疗结论', ''))
                    }

                    # 构建完整的案件数据（保留所有现有字段）
                    index_data['cases'][i] = {
                        'case_number': case_number,
                        'person_name': person_name,
                        'applicant': data.get('申请人', ''),  # 新增
                        'case_type': data.get('案件类型', existing_case.get('case_type', '')),
                        'year': datetime.now().year,
                        'folder_path': existing_case.get('folder_path', f"{datetime.now().year}/{case_number}"),
                        'created_date': existing_case.get('created_date', datetime.now().strftime('%Y-%m-%d')),
                        'employer': data.get('用人单位', existing_case.get('employer', '')),
                        'work_unit': data.get('用工单位', existing_case.get('work_unit', '')),
                        'workplace': data.get('工作场所', existing_case.get('workplace', '')),
                        'regulation': data.get('条例', existing_case.get('regulation', '')),
                        'operator': data.get('操作员', existing_case.get('operator', '')),
                        'person_info': new_person_info,
                        'witnesses': existing_case.get('witnesses', []),  # 保留现有证人
                        'legal_persons': existing_case.get('legal_persons', [])  # 保留现有法人
                    }
                    found = True
                    break

            if not found:
                # 新建案件（没有现有数据）
                case_data = {
                    'case_number': case_number,
                    'person_name': person_name,
                    'applicant': data.get('申请人', ''),  # 新增
                    'case_type': data.get('案件类型', ''),
                    'year': datetime.now().year,
                    'folder_path': f"{datetime.now().year}/{case_number}",
                    'created_date': datetime.now().strftime('%Y-%m-%d'),
                    'employer': data.get('用人单位', ''),
                    'work_unit': data.get('用工单位', ''),
                    'workplace': data.get('工作场所', ''),
                    'regulation': data.get('条例', ''),
                    'operator': data.get('操作员', ''),
                    'person_info': {
                        'name': data.get('本人姓名', ''),
                        'gender': data.get('本人性别', ''),
                        'age': data.get('本人年龄', ''),
                        'phone': data.get('本人电话', ''),
                        'id_card': data.get('本人身份证号', ''),
                        'address': data.get('本人身份证地址', ''),
                        'current_address': data.get('本人现住址', ''),
                        'position': data.get('本人岗位', ''),
                        '自我介绍': data.get('自我介绍', ''),
                        '受伤经过': data.get('受伤经过', ''),
                        '就医情况': data.get('就医情况', ''),
                        '医疗结论': data.get('医疗结论', '')
                    },
                    'witnesses': [],
                    'legal_persons': []
                }
                index_data['cases'].append(case_data)

            # 更新统计信息
            index_data['total_cases'] = len(index_data['cases'])
            index_data['last_update'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # 先写临时文件，再替换（防止文件损坏）
            temp_file = index_file + ".tmp"
            with open(temp_file, 'w', encoding='utf-8') as f:
                json.dump(index_data, f, ensure_ascii=False, indent=2)

            # 替换原文件
            shutil.move(temp_file, index_file)

        except Exception as e:
            print(f"更新索引失败: {e}")
            import traceback
            traceback.print_exc()

    def get_current_year_folder(self):
        """获取当前年份的cases文件夹"""
        current_year = datetime.now().year
        year_folder = os.path.join(self.BASE_DIR, str(current_year))
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

    def insert_description_into_doc(self, doc, data):
        """
        将自我介绍插入到文档中（使用占位符替换）
        不再依赖特定问题文本，而是直接替换 {自我介绍} 占位符
        """
        description = self.generate_description(data)

        # 遍历所有段落，替换 {自我介绍} 占位符
        for paragraph in doc.paragraphs:
            if "{自我介绍}" in paragraph.text:
                paragraph.text = paragraph.text.replace("{自我介绍}", description)

        # 也检查表格中的单元格
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if "{自我介绍}" in paragraph.text:
                            paragraph.text = paragraph.text.replace("{自我介绍}", description)

        return doc

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
        # 测试数据（10组）
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
            },
            {
                "name": "赵六",
                "id_card": "410101198503044567",
                "id_address": "河南省新乡市红旗区平原路15号",
                "current_address": "河南省新乡市卫滨区解放路20号院",
                "phone": "13600136003",
                "position": "操作工"
            },
            {
                "name": "孙七",
                "id_card": "410101197808055678",
                "id_address": "河南省焦作市山阳区塔南路25号",
                "current_address": "河南省焦作市中站区跃进路8号",
                "phone": "13500135004",
                "position": "维修工"
            },
            {
                "name": "周八",
                "id_card": "410101199212066789",
                "id_address": "河南省濮阳市华龙区中原路30号",
                "current_address": "河南省濮阳市濮阳县红旗路12号",
                "phone": "13400134005",
                "position": "电工"
            },
            {
                "name": "吴九",
                "id_card": "410101198909077890",
                "id_address": "河南省许昌市魏都区七一路18号",
                "current_address": "河南省许昌市建安区新许路6号",
                "phone": "13300133006",
                "position": "焊工"
            },
            {
                "name": "郑十",
                "id_card": "410101199311088901",
                "id_address": "河南省漯河市郾城区黄河路22号",
                "current_address": "河南省漯河市源汇区人民路15号",
                "phone": "13200132007",
                "position": "司机"
            },
            {
                "name": "钱多多",
                "id_card": "410101198012099012",
                "id_address": "河南省三门峡市湖滨区崤山路35号",
                "current_address": "河南省三门峡市陕州区神泉路9号",
                "phone": "13100131008",
                "position": "仓库管理员"
            },
            {
                "name": "刘能",
                "id_card": "410101199510101123",
                "id_address": "河南省南阳市卧龙区中州路45号",
                "current_address": "河南省南阳市宛城区建设路28号",
                "phone": "13000130009",
                "position": "质检员"
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

        # 更新索引（循环）
        self.test_index = (self.test_index + 1) % len(test_data)

        self.statusBar().showMessage(f"已填入测试数据 ({self.test_index}/{len(test_data)}): {data['name']}", 2000)

    def get_template_name(self, data):
        """根据人员类型和条例返回对应的模板文件名"""
        person_type = data['人员类型']  # 本人/证人/法人
        regulation = data.get('条例', '')  # 获取条例

        # 条例到文件名的映射
        regulation_map = {
            "第十四条第一款第一项（普通工伤案件）": "普通工伤案件",
            "第十四条第一款第二项（预备收尾案件）": "预备收尾案件",
            "第十四条第一款第三项（暴力伤害案件）": "暴力伤害案件",
            "第十四条第一款第四项（患职业病案件）": "患职业病案件",
            "第十四条第一款第五项（因工外出案件）": "因工外出案件",
            "第十四条第一款第六项（上下班时案件）": "上下班时案件",
            "第十五条第一款第一项（工作时因病亡故案件）": "工作时因病亡故案件",
            # 可以继续添加其他映射
        }

        # 获取条例对应的案件类型，如果没有匹配则用"普通工伤案件"
        case_type = regulation_map.get(regulation, "普通工伤案件")

        # 生成模板名：人员类型 + 谈话笔录（ + 案件类型 + ）.docx
        template_name = f"{person_type}谈话笔录（{case_type}）.docx"
        return template_name


def main():
    app = QApplication(sys.argv)

    app.setApplicationName("工伤案件管理系统")
    app.setOrganizationName("WorkInjuryApp")

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()