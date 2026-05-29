# -*- coding: utf-8 -*-
"""
WebDriver 工厂

使用 Selenium Manager 自动管理 ChromeDriver：
- Selenium 4.6+ 内置 Selenium Manager
- 通过环境变量配置国内镜像，无需科学上网
- 自动版本匹配，无需手动下载
"""

import os
import sys

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from .chromedriver_manager import (
    find_chrome_path,
    setup_selenium_manager_mirror,
)


def create_driver():
    """创建 Chrome WebDriver
    
    优先级：
    1. 程序同目录下的 chromedriver.exe（手动放置）
    2. Selenium Manager 自动管理（使用国内镜像）
    """
    chrome_path = find_chrome_path()
    if not chrome_path:
        raise Exception(
            "未找到 Chrome 浏览器！请确认已安装 Google Chrome。")

    # 优先检查程序同目录下是否有 chromedriver.exe（手动放置的情况）
    if getattr(sys, 'frozen', False):
        app_dir = os.path.dirname(sys.executable)
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))
    local_driver = os.path.join(app_dir, "chromedriver.exe")
    if os.path.isfile(local_driver):
        from selenium.webdriver.chrome.service import Service
        opts = _create_chrome_options()
        service = Service(executable_path=local_driver)
        return webdriver.Chrome(service=service, options=opts)

    # 配置 Selenium Manager 使用国内镜像
    setup_selenium_manager_mirror()

    try:
        # Selenium Manager 会自动：
        # 1. 检测本地 Chrome 版本
        # 2. 从镜像下载匹配的 ChromeDriver
        # 3. 缓存到 ~/.cache/selenium
        driver = webdriver.Chrome(options=_create_chrome_options())
        return driver

    except Exception as e:
        raise Exception(
            f"启动 Chrome 失败: {e}\n\n"
            "解决方案：\n"
            "1. 确认已安装 Google Chrome 浏览器\n"
            "2. 检查网络连接是否正常\n"
            "3. 如仍有问题，可手动下载 ChromeDriver：\n"
            "   - 打开 Chrome 设置 → 关于 Chrome，记下版本号\n"
            "   - 访问 https://registry.npmmirror.com/binary.html?path=chrome-for-testing/\n"
            "   - 下载对应版本的 win64/chromedriver-win64.zip 并解压\n"
            "   - 将 chromedriver.exe 放到程序同目录下")


def _create_chrome_options():
    """创建 Chrome 选项"""
    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    return opts