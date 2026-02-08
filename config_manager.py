#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
安全配置管理器 - 使用QSettings
"""

import base64
from PyQt5.QtCore import QSettings


class ConfigManager:
    """配置管理器 - 最简单实现"""

    def __init__(self):
        # QSettings会自动处理不同系统的存储
        self.settings = QSettings("WorkInjuryApp", "Config")

    def save_config(self, operator="", api_url="", api_key="", remember=False):
        """
        保存配置
        """
        # 保存基础配置
        self.settings.setValue("operator", operator)
        self.settings.setValue("api_url", api_url)
        self.settings.setValue("remember", remember)

        # API密钥简单编码后保存
        if api_key and remember:
            encoded = base64.b64encode(api_key.encode()).decode()
            self.settings.setValue("api_key_encoded", encoded)
        else:
            self.settings.remove("api_key_encoded")

    def load_config(self):
        """
        加载配置
        返回: dict {operator, api_url, api_key, remember}
        """
        # 检查是否记住
        remember = self.settings.value("remember", False, type=bool)

        # 加载基础配置
        operator = self.settings.value("operator", "", type=str)
        api_url = self.settings.value("api_url", "", type=str)

        # 加载API密钥
        api_key = ""
        if remember:
            encoded = self.settings.value("api_key_encoded", "", type=str)
            if encoded:
                try:
                    api_key = base64.b64decode(encoded.encode()).decode()
                except:
                    api_key = ""

        return {
            "operator": operator,
            "api_url": api_url,
            "api_key": api_key,
            "remember": remember
        }

    def clear_config(self):
        """清除所有配置"""
        self.settings.remove("operator")
        self.settings.remove("api_url")
        self.settings.remove("api_key_encoded")
        self.settings.setValue("remember", False)