@echo off
chcp 65001 > nul
echo ============================================
echo   FALCON2 インストーラー作成
echo ============================================
echo.

cd /d %~dp0..

REM dist/installer フォルダを作成
if not exist dist\installer mkdir dist\installer

REM Inno Setup のパスを探す
set ISCC=
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"               set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if exist "C:\Program Files\Inno Setup 6\ISCC.exe"                     set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"
if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"              set ISCC="%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"

if "%ISCC%"=="" (
    echo [ERROR] Inno Setup 6 が見つかりません。
    echo         https://jrsoftware.org/isinfo.php からインストールしてください。
    pause
    exit /b 1
)

echo [1/2] EXE がビルド済みか確認中...
if not exist dist\FALCON2\FALCON2.exe (
    echo [ERROR] dist\FALCON2\FALCON2.exe が見つかりません。
    echo         先に tools\build_exe.bat を実行してください。
    pause
    exit /b 1
)

echo [2/2] インストーラーをビルド中...
%ISCC% tools\falcon2_setup.iss
if errorlevel 1 (
    echo [ERROR] インストーラーの作成に失敗しました。
    pause
    exit /b 1
)

echo.
echo ============================================
echo   完了！
echo   出力先: dist\installer\FALCON2_Setup_v1.0.0.exe
echo ============================================
echo.
pause
