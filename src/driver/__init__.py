# -*- coding: utf-8 -*-
"""
浏览器驱动管理模块

使用 Selenium Manager 自动管理 ChromeDriver：
- Selenium 4.6+ 内置 Selenium Manager
- 通过环境变量配置国内镜像，无需科学上网
- 自动版本匹配，无需手动下载
"""

from .driver_factory import create_driver
from .chromedriver_manager import (
    find_chrome_path,
    get_chrome_version,
    setup_selenium_manager_mirror,
)