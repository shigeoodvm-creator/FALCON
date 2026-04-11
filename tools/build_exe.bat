@echo off
chcp 65001 > nul
echo ============================================
echo   FALCON2 EXE ビルド
echo ============================================
echo.

cd /d %~dp0..

REM ── PyInstaller が入っているか確認 ──
python -m PyInstaller --version > nul 2>&1
if errorlevel 1 (
    echo [ERROR] PyInstaller が見つかりません。
    echo         pip install pyinstaller を実行してください。
    pause
    exit /b 1
)

REM ── 前回のビルド成果物をクリア ──
echo [1/3] 前回のビルドをクリア中...
if exist build\FALCON2  rmdir /s /q build\FALCON2
if exist dist\FALCON2   rmdir /s /q dist\FALCON2
echo       完了

REM ── ビルド実行 ──
echo [2/3] PyInstaller ビルド中...（数分かかります）
python -m PyInstaller falcon2.spec --noconfirm
if errorlevel 1 (
    echo.
    echo [ERROR] ビルドに失敗しました。上記のエラーを確認してください。
    pause
    exit /b 1
)

REM ── config フォルダを dist に配置（初期設定ファイル用） ──
echo [3/3] config フォルダを配置中...
if not exist dist\FALCON2\config mkdir dist\FALCON2\config
if exist config\activity_log.json (
    copy /y config\activity_log.json dist\FALCON2\config\ > nul
)
if exist config\app_settings.json (
    copy /y config\app_settings.json dist\FALCON2\config\ > nul
)

echo.
echo ============================================
echo   ビルド完了！
echo   出力先: dist\FALCON2\FALCON2.exe
echo ============================================
echo.
echo 配布前に以下を確認してください：
echo   1. dist\FALCON2\FALCON2.exe を起動して動作確認
echo   2. 試用期間カウントが正常に開始されるか確認
echo   3. ライセンスキーの認証が通るか確認
echo.
pause
