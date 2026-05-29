# -*- coding: utf-8 -*-
"""
ChromeDriver 管理：自动版本匹配 + 国内镜像下载

方案说明：
- Selenium 4.6+ 内置 Selenium Manager，可自动管理 ChromeDriver
- 通过配置 SE_DRIVER_MIRROR_URL 环境变量，使用 npmmirror 国内镜像
- 无需科学上网即可自动下载匹配的 ChromeDriver
"""

import os
import sys
import ctypes
import winreg

# 国内镜像配置（npmmirror 提供 Chrome for Testing 镜像）
# Selenium Manager 会使用此 URL 作为驱动下载源
DRIVER_MIRROR_URL = "https://registry.npmmirror.com/-/binary/chrome-for-testing"

# 浏览器镜像配置（可选，用于自动下载 Chrome）
BROWSER_MIRROR_URL = "https://registry.npmmirror.com/-/binary/chrome-for-testing"


def setup_selenium_manager_mirror():
    """配置 Selenium Manager 使用国内镜像
    
    Selenium Manager 支持三种配置方式（优先级从高到低）：
    1. 命令行参数
    2. 配置文件 (~/.cache/selenium/se-config.toml)
    3. 环境变量（SE_ 前缀）
    
    这里使用环境变量方式，最简单且无需修改配置文件。
    """
    # 设置驱动镜像 URL
    if "SE_DRIVER_MIRROR_URL" not in os.environ:
        os.environ["SE_DRIVER_MIRROR_URL"] = DRIVER_MIRROR_URL
    
    # 设置浏览器镜像 URL（可选，用于自动下载 Chrome）
    if "SE_BROWSER_MIRROR_URL" not in os.environ:
        os.environ["SE_BROWSER_MIRROR_URL"] = BROWSER_MIRROR_URL
    
    # 关闭匿名统计数据收集
    if "SE_AVOID_STATS" not in os.environ:
        os.environ["SE_AVOID_STATS"] = "true"


def find_chrome_path():
    """查找本地 Chrome 可执行文件路径"""
    candidates = [
        os.path.expandvars(
            r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(
            r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(
            r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]
    for p in candidates:
        if os.path.isfile(p):
            return p
    return None


def get_chrome_version(chrome_path):
    """获取 Chrome 版本号"""
    if sys.platform != "win32":
        return None

    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Google\Chrome\BLBeacon")
        version, _ = winreg.QueryValueEx(key, "version")
        return version
    except Exception:
        pass

    try:
        size = ctypes.windll.version.GetFileVersionInfoSizeW(
            chrome_path, None)
        if size:
            buf = ctypes.create_string_buffer(size)
            ctypes.windll.version.GetFileVersionInfoW(
                chrome_path, None, size, buf)
            res = ctypes.create_string_buffer(size)
            ctypes.windll.version.VerQueryValueW(
                buf, r"\VarFileInfo\Translation", ctypes.byref(res),
                ctypes.byref(ctypes.c_ulong()))
            trans = res.raw[:8]
            lang_code = trans.hex()
            query = f"\\StringFileInfo\\{lang_code}\\FileVersion"
            ctypes.windll.version.VerQueryValueW(
                buf, query, ctypes.byref(res),
                ctypes.byref(ctypes.c_ulong()))
            return res.value.decode('utf-16le').rstrip('\x00')
    except Exception:
        pass

    return None