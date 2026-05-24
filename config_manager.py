#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
安全配置管理器 - 使用 QSettings + Windows DPAPI 加密
"""
import ctypes
from ctypes import wintypes
from PyQt5.QtCore import QSettings

SERVICE_NAME = "WorkInjuryApp"


class ConfigManager:
    """配置管理器 — API Key 使用 Windows DPAPI 加密存储"""

    def __init__(self):
        self.settings = QSettings("WorkInjuryApp", "Config")

    @staticmethod
    def _encrypt(plain_text):
        """使用 Windows DPAPI 加密字符串"""
        data_in = plain_text.encode('utf-16le')
        blob_in = ctypes.create_string_buffer(data_in, len(data_in))

        data_in_struct = ctypes.c_buffer(data_in)
        blob_in_struct = ctypes.c_buffer(blob_in.raw)

        class DATA_BLOB(ctypes.Structure):
            _fields_ = [
                ("cbData", wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_char)),
            ]

        data_in_blob = DATA_BLOB(len(data_in), ctypes.cast(data_in_struct, ctypes.POINTER(ctypes.c_char)))
        data_out_blob = DATA_BLOB()

        if not ctypes.windll.crypt32.CryptProtectData(
            ctypes.byref(data_in_blob),
            SERVICE_NAME,
            None, None, None,
            0,  # CRYPTPROTECT_UI_FORBIDDEN
            ctypes.byref(data_out_blob),
        ):
            return ""

        encrypted = ctypes.string_at(data_out_blob.pbData, data_out_blob.cbData)
        ctypes.windll.kernel32.LocalFree(data_out_blob.pbData)
        return encrypted.hex()

    @staticmethod
    def _decrypt(hex_data):
        """使用 Windows DPAPI 解密字符串"""
        if not hex_data:
            return ""
        encrypted = bytes.fromhex(hex_data)

        data_in_struct = ctypes.c_buffer(encrypted)

        class DATA_BLOB(ctypes.Structure):
            _fields_ = [
                ("cbData", wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_char)),
            ]

        data_in_blob = DATA_BLOB(len(encrypted), ctypes.cast(data_in_struct, ctypes.POINTER(ctypes.c_char)))
        data_out_blob = DATA_BLOB()

        if not ctypes.windll.crypt32.CryptUnprotectData(
            ctypes.byref(data_in_blob),
            None, None, None, None,
            0,
            ctypes.byref(data_out_blob),
        ):
            return ""

        decrypted = ctypes.string_at(data_out_blob.pbData, data_out_blob.cbData)
        ctypes.windll.kernel32.LocalFree(data_out_blob.pbData)
        return decrypted.decode('utf-16le', errors='ignore')

    def save_config(self, operator="", api_url="", api_key="", remember=False):
        self.settings.setValue("operator", operator)
        self.settings.setValue("api_url", api_url)
        self.settings.setValue("remember", remember)

        if api_key and remember:
            encrypted = self._encrypt(api_key)
            self.settings.setValue("api_key_dpapi", encrypted)
        else:
            self.settings.remove("api_key_dpapi")

    def load_config(self):
        remember = self.settings.value("remember", False, type=bool)
        operator = self.settings.value("operator", "", type=str)
        api_url = self.settings.value("api_url", "", type=str)

        api_key = ""
        if remember:
            encrypted = self.settings.value("api_key_dpapi", "", type=str)
            if encrypted:
                api_key = self._decrypt(encrypted)

        return {
            "operator": operator,
            "api_url": api_url,
            "api_key": api_key,
            "remember": remember
        }

    def clear_config(self):
        self.settings.remove("operator")
        self.settings.remove("api_url")
        self.settings.remove("api_key_dpapi")
        self.settings.setValue("remember", False)
