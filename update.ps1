# ========================================================
# X_AutoTweetBot 最新バージョン自動ビルド＆上書きアップデートスクリプト
# ========================================================

$ErrorActionPreference = "Continue"

try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host " X_AutoTweetBot 最新版アップデートを開始します..." -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan

# 作業ディレクトリの決定（スクリプトのある場所を優先）
$scriptDir = $PSScriptRoot
if (-not $scriptDir) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if ($scriptDir -and (Test-Path (Join-Path $scriptDir "auto_tweet_playwright.py"))) {
    Set-Location $scriptDir
} elseif (Test-Path "C:\Users\PC_User\Desktop\X_BOT") {
    Set-Location "C:\Users\PC_User\Desktop\X_BOT"
}
Write-Host "作業フォルダ: $(Get-Location)" -ForegroundColor Gray

# 配布フォルダの自動検出（デスクトップ / OneDrive / 親フォルダ等）
$desktopDir = [Environment]::GetFolderPath("Desktop")
$portableCandidates = @(
    (Join-Path $desktopDir "X_BOT_配布用(Python不要)"),
    (Join-Path $env:USERPROFILE "Desktop\X_BOT_配布用(Python不要)"),
    (Join-Path $env:USERPROFILE "OneDrive\Desktop\X_BOT_配布用(Python不要)"),
    (Join-Path $env:USERPROFILE "OneDrive\デスクトップ\X_BOT_配布用(Python不要)"),
    "C:\Users\PC_User\Desktop\X_BOT_配布用(Python不要)",
    (Join-Path (Split-Path -Parent (Get-Location)) "X_BOT_配布用(Python不要)")
)

$portableDir = $null
foreach ($cand in $portableCandidates) {
    if ($cand -and (Test-Path $cand)) {
        $portableDir = $cand
        break
    }
}
if ($portableDir) {
    Write-Host "配布フォルダ検出: $portableDir" -ForegroundColor Green
} else {
    Write-Host "※ 既存の配布フォルダが見つからなかったため、distフォルダに出力します。" -ForegroundColor Yellow
}

# 1. 必要フォルダ（images / log / screenshots）の確認・作成
Write-Host "`n[1/4] images / log / screenshots フォルダを確認・作成中..." -ForegroundColor Yellow
foreach ($subDir in @("images", "log", "screenshots")) {
    if (-not (Test-Path $subDir)) {
        New-Item -ItemType Directory -Path $subDir -Force | Out-Null
        Write-Host "  + フォルダ作成: $subDir" -ForegroundColor Green
    }
    if ($portableDir) {
        $portSub = Join-Path $portableDir $subDir
        if (-not (Test-Path $portSub)) {
            New-Item -ItemType Directory -Path $portSub -Force | Out-Null
            Write-Host "  + フォルダ作成: $portSub" -ForegroundColor Green
        }
    }
}

# 2. Pythonパスの検出
$pyExe = $null
$pythonCandidates = @(
    "C:\Users\PC_User\AppData\Local\Programs\Python\Python312\python.exe",
    (Join-Path $env:LOCALAPPDATA "Programs\Python\Python312\python.exe"),
    (Join-Path $env:LOCALAPPDATA "Programs\Python\Python311\python.exe"),
    (Join-Path $env:LOCALAPPDATA "Programs\Python\Python313\python.exe"),
    (Join-Path $env:LOCALAPPDATA "Programs\Python\Python310\python.exe")
)
foreach ($p in $pythonCandidates) {
    if ($p -and (Test-Path $p)) {
        $pyExe = $p
        break
    }
}
if (-not $pyExe) {
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd -and ($cmd.Source -notmatch "WindowsApps")) { $pyExe = $cmd.Source }
}
if (-not $pyExe) {
    $cmdPy = Get-Command py -ErrorAction SilentlyContinue
    if ($cmdPy) { $pyExe = $cmdPy.Source }
}

if (-not $pyExe) {
    Write-Host "`n❌ Python が見つかりませんでした。Python 3.10〜3.12 をインストールしてから再実行してください。" -ForegroundColor Red
    Read-Host "Enterキーを押すと終了します"
    exit 1
}
Write-Host "使用Python: $pyExe" -ForegroundColor Gray

# openpyxl / playwright / pyinstaller の確認・不足時は自動インストール
& $pyExe -m PyInstaller --version *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Host "必要ライブラリを自動インストールします..." -ForegroundColor Yellow
    & $pyExe -m pip install pyinstaller playwright openpyxl pillow
}

# 3. Excelファイルのテーマ別レイアウト更新＆カラー書式保護
Write-Host "`n[2/4] Excelファイルをテーマ別レイアウト・削除耐性カラー書式に更新中..." -ForegroundColor Yellow
if (Test-Path "update_themes_excel.py") {
    & $pyExe update_themes_excel.py
}

# 4. PyInstallerで X_AutoTweetBot.exe を再ビルド（必要な本体のみ）
Write-Host "`n[3/4] 最新 X_AutoTweetBot.exe をビルド中 (PyInstaller)..." -ForegroundColor Yellow
if (Test-Path "auto_tweet_playwright.py") {
    & $pyExe -m PyInstaller `
        --noconfirm `
        --onedir `
        --console `
        --name "X_AutoTweetBot" `
        --collect-all playwright `
        --collect-all openpyxl `
        auto_tweet_playwright.py
}

# 5. 配布フォルダへの上書き反映（不要なランチャー類は生成しない＆旧不要ファイルを整理）
Write-Host "`n[4/4] 配布フォルダへ必要ファイルのみをコピー中..." -ForegroundColor Yellow
$distDir = "dist\X_AutoTweetBot"

$obsoleteFiles = @(
    "START_ROTATION.vbs",
    "START_ROTATION.bat",
    "CHROME_LAUNCHER.vbs",
    "CHROME_LAUNCHER.bat",
    "chrome_launcher.py",
    "chrome_launcher.exe",
    "START.bat",
    "START_DM.bat",
    "TEST_ONE_TWEET.bat",
    "BUILD_EXE.bat",
    "SETUP.bat",
    "auto_dm_playwright.py",
    "auto_tweet_selenium.py",
    "auto_tweet_api.py",
    "x_auto_tweet_colab.ipynb",
    "dm_targets.xlsx",
    "アプデ実行.bat",
    "browser_auto_tweet.log"
)

if ($portableDir) {
    if (Test-Path $distDir) {
        if (Test-Path "$distDir\X_AutoTweetBot.exe") {
            Copy-Item "$distDir\X_AutoTweetBot.exe" -Destination "$portableDir\X_AutoTweetBot.exe" -Force
        }
        if (Test-Path "$distDir\_internal") {
            New-Item -ItemType Directory -Path "$portableDir\_internal" -Force | Out-Null
            Copy-Item "$distDir\_internal\*" -Destination "$portableDir\_internal" -Recurse -Force -ErrorAction SilentlyContinue
        }

        $msBrowsers = "$env:LOCALAPPDATA\ms-playwright"
        $targetBrowsers = "$portableDir\_internal\playwright\driver\package\.local-browsers"
        if (Test-Path $msBrowsers) {
            New-Item -ItemType Directory -Path $targetBrowsers -Force | Out-Null
            Copy-Item "$msBrowsers\*" -Destination $targetBrowsers -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    # 必要な最小限の実行・GUIファイルのみ配置（START_ROTATION.vbs等の重複・不要ファイルは生成しない）
    foreach ($helperFile in @("auto_tweet_playwright.py", "update_themes_excel.py", "x_bot_gui.py", "X_BOT_GUI.vbs", "STOP_BOT.vbs")) {
        if (Test-Path $helperFile) {
            Copy-Item $helperFile -Destination (Join-Path $portableDir $helperFile) -Force
        }
    }

    # 旧バージョンの不要ランチャーを整理（Excel・画像・セッション・ログ・スクショ等のユーザーデータは一切触らない）
    foreach ($obs in $obsoleteFiles) {
        $obsPath = Join-Path $portableDir $obs
        if (Test-Path $obsPath) {
            Remove-Item $obsPath -Force -ErrorAction SilentlyContinue
            Write-Host "  - 不要ファイル整理: $obs" -ForegroundColor DarkGray
        }
    }

    try {
        Get-ChildItem -Path $portableDir -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue
    } catch {}

    Write-Host "✅ 配布フォルダ ($portableDir) への反映が完了しました！" -ForegroundColor Green
} else {
    Write-Host "✅ distフォルダ内にビルドが完了しました:" -ForegroundColor Green
    if (Test-Path "$distDir\X_AutoTweetBot.exe") {
        Write-Host "   - $(Join-Path (Get-Location) "$distDir\X_AutoTweetBot.exe")" -ForegroundColor Green
    }
}

# 作業フォルダ内の不要な一時ビルドファイルと旧不要ファイルを整理
foreach ($obsLocal in @("START_ROTATION.vbs", "START_ROTATION.bat", "CHROME_LAUNCHER.vbs", "CHROME_LAUNCHER.bat", "chrome_launcher.py", "START.bat", "START_DM.bat", "TEST_ONE_TWEET.bat", "BUILD_EXE.bat", "SETUP.bat", "X_AutoTweetBot.spec", "chrome_launcher.spec", "browser_auto_tweet.log")) {
    if (Test-Path $obsLocal) {
        Remove-Item $obsLocal -Force -ErrorAction SilentlyContinue
    }
}
if (Test-Path "build") {
    Remove-Item "build" -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "========================================================" -ForegroundColor Green
Write-Host "【アップデート完了】最新機能・スッキリ構成へ更新しました！" -ForegroundColor Green
Write-Host "1. 【アカウント別ログ】log/<アカウント名>.log にアカウント単位で自動保存" -ForegroundColor Yellow
Write-Host "2. 【アカウント別スクショ】screenshots/<アカウント名>/ に投稿・エラー画面を保存" -ForegroundColor Yellow
Write-Host "3. 【不要ファイル削除】X_BOT_GUI.vbs と STOP_BOT.vbs に集約（START_ROTATION.vbs等なし）" -ForegroundColor Yellow
Write-Host "4. 【テーマ管理＆時間フィルタ】各シート1〜2行目に設定、時間外ポストは即スキップ" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Green
Write-Host ""
Read-Host "Enterキーを押すとウィンドウを閉じます"
