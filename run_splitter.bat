@echo off
chcp 65001 >nul
echo ========================================
echo    Word文档拆分工具
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误: 未找到Python，请先安装Python 3.11+
    pause
    exit /b 1
)

echo ✅ Python已安装

REM 检查依赖包
echo 📦 检查依赖包...
python -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo ⚠️  未找到python-docx，正在安装...
    pip install python-docx
    if errorlevel 1 (
        echo ❌ 依赖包安装失败
        pause
        exit /b 1
    )
    echo ✅ 依赖包安装成功
) else (
    echo ✅ 依赖包已安装
)

REM 检查输入目录
if not exist "input" (
    echo 📁 创建输入目录...
    mkdir input
    echo ✅ 输入目录已创建: %CD%\input
    echo.
    echo 💡 请将需要拆分的Word文档放入input目录，然后重新运行此脚本
    pause
    exit /b 0
)

REM 检查输入文件
dir /b "input\*.docx" "input\*.doc" >nul 2>&1
if errorlevel 1 (
    echo ⚠️  在input目录中未找到Word文档
    echo 💡 请将.docx或.doc文件放入input目录
    echo 📂 输入目录位置: %CD%\input
    pause
    exit /b 0
)

echo 📄 找到Word文档，开始处理...
echo.

REM 运行主程序
python main.py

if errorlevel 1 (
    echo.
    echo ❌ 程序执行过程中发生错误
    echo 📋 请查看word_splitter.log文件获取详细信息
) else (
    echo.
    echo 🎉 处理完成！
    echo 📂 结果保存在: %CD%\output
    echo.
    echo 是否打开输出目录？ (Y/N)
    set /p choice=
    if /i "%choice%"=="Y" (
        explorer "output"
    )
)

echo.
echo 按任意键退出...
pause >nul