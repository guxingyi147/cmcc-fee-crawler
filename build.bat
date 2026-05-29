@echo off
chcp 65001 >nul
echo ============================================
echo   中国移动资费公示爬取工具 - 打包脚本
echo ============================================
echo.

:: 检查Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到Python，请先安装Python 3.8+
    pause
    exit /b 1
)

:: 检查依赖
echo [1/3] 检查并安装依赖...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [错误] 依赖安装失败
    pause
    exit /b 1
)

:: 检查pyinstaller
echo [2/3] 检查PyInstaller...
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在安装PyInstaller...
    pip install pyinstaller
)

:: 打包（添加隐藏导入，解决打包后模块缺失问题）
echo [3/3] 开始打包...
pyinstaller --noconfirm --onedir --windowed --name "cmcc_crawler" ^
    --hidden-import=src ^
    --hidden-import=src.gui ^
    --hidden-import=src.gui.main_window ^
    --hidden-import=src.crawler ^
    --hidden-import=src.crawler.crawler_thread ^
    --hidden-import=src.driver ^
    --hidden-import=src.driver.chromedriver_manager ^
    --hidden-import=src.driver.driver_factory ^
    --hidden-import=src.parser ^
    --hidden-import=src.parser.text_parser ^
    --hidden-import=src.export ^
    --hidden-import=src.export.excel_writer ^
    --hidden-import=src.constants ^
    --hidden-import=selenium ^
    --hidden-import=selenium.webdriver ^
    --hidden-import=selenium.webdriver.chrome ^
    --hidden-import=selenium.webdriver.chrome.webdriver ^
    --hidden-import=selenium.webdriver.chrome.service ^
    --hidden-import=selenium.webdriver.chrome.options ^
    --hidden-import=selenium.webdriver.support ^
    --hidden-import=selenium.webdriver.support.ui ^
    --hidden-import=selenium.webdriver.support.expected_conditions ^
    --hidden-import=selenium.webdriver.common.by ^
    --hidden-import=selenium.webdriver.common.keys ^
    --hidden-import=selenium.webdriver.common.action_chains ^
    --hidden-import=urllib.request ^
    --hidden-import=json ^
    --hidden-import=ssl ^
    main.py

if %errorlevel% equ 0 (
    echo.
    echo ============================================
    echo   打包成功！
    echo   输出目录: dist\cmcc_crawler\
    echo   可执行文件: dist\cmcc_crawler\cmcc_crawler.exe
    echo ============================================
) else (
    echo.
    echo [错误] 打包失败，请检查错误信息
)

pause
