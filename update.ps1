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

# 1. 必要な画像フォルダ（朝枠・昼枠・夕方枠・夜枠）の自動作成
Write-Host "`n[1/5] 画像フォルダを確認・作成中..." -ForegroundColor Yellow
$slotNames = @("朝枠", "昼枠", "夕方枠", "夜枠")
foreach ($s in $slotNames) {
    $localImg = Join-Path "images" $s
    if (-not (Test-Path $localImg)) {
        New-Item -ItemType Directory -Path $localImg -Force | Out-Null
        Write-Host "  + フォルダ作成: $localImg" -ForegroundColor Green
    }
    if ($portableDir) {
        $portImg = Join-Path $portableDir (Join-Path "images" $s)
        if (-not (Test-Path $portImg)) {
            New-Item -ItemType Directory -Path $portImg -Force | Out-Null
            Write-Host "  + フォルダ作成: $portImg" -ForegroundColor Green
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
    if ($cmd) { $pyExe = $cmd.Source }
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

# PyInstaller / openpyxl / playwright の確認・不足時は自動インストール
& $pyExe -m PyInstaller --version *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Host "PyInstaller が未導入のため自動インストールします..." -ForegroundColor Yellow
    & $pyExe -m pip install pyinstaller playwright openpyxl pillow
}

# 3. Excelに日常ツイートシートを追加更新
Write-Host "`n[2/5] Excelファイルの日常ツイートシートを更新中..." -ForegroundColor Yellow
if (Test-Path "add_daily_tweets_excel.py") {
    & $pyExe add_daily_tweets_excel.py
}

# 4. PyInstallerで X_AutoTweetBot.exe を再ビルド
Write-Host "`n[3/5] 最新 X_AutoTweetBot.exe をビルド中 (PyInstaller)..." -ForegroundColor Yellow
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

# 4.5. PyInstallerで chrome_launcher.exe をビルド
Write-Host "`n[4/5] 初回手動ログイン用 chrome_launcher.exe をビルド中 (PyInstaller)..." -ForegroundColor Yellow
if (Test-Path "chrome_launcher.py") {
    & $pyExe -m PyInstaller `
        --noconfirm `
        --onefile `
        --console `
        --name "chrome_launcher" `
        chrome_launcher.py
}

# 5. 配布フォルダへの上書き反映
Write-Host "`n[5/5] 配布フォルダへ最新ファイルをコピー中..." -ForegroundColor Yellow
$distDir = "dist\X_AutoTweetBot"

if ($portableDir) {
    if (Test-Path $distDir) {
        if (Test-Path "$distDir\X_AutoTweetBot.exe") {
            Copy-Item "$distDir\X_AutoTweetBot.exe" -Destination "$portableDir\X_AutoTweetBot.exe" -Force
        }
        if (Test-Path "$distDir\_internal") {
            New-Item -ItemType Directory -Path "$portableDir\_internal" -Force | Out-Null
            Copy-Item "$distDir\_internal\*" -Destination "$portableDir\_internal" -Recurse -Force -ErrorAction SilentlyContinue
        }

        # Playwrightブラウザの実体をポータブルフォルダへバックアップ配置
        $msBrowsers = "$env:LOCALAPPDATA\ms-playwright"
        $targetBrowsers = "$portableDir\_internal\playwright\driver\package\.local-browsers"
        if (Test-Path $msBrowsers) {
            New-Item -ItemType Directory -Path $targetBrowsers -Force | Out-Null
            Copy-Item "$msBrowsers\*" -Destination $targetBrowsers -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    if (Test-Path "dist\chrome_launcher.exe") {
        Copy-Item "dist\chrome_launcher.exe" -Destination "$portableDir\chrome_launcher.exe" -Force
        Write-Host "✅ chrome_launcher.exe を配布フォルダにコピーしました！" -ForegroundColor Green
    }

    $currentUpdatePs1 = Join-Path (Get-Location) "update.ps1"
    $updateBatContent = "@echo off`r`nchcp 65001 > nul`r`necho 最新版へアップデート中...`r`npowershell -NoProfile -ExecutionPolicy Bypass -File `"$currentUpdatePs1`"`r`npause`r`n"
    Set-Content -Path "$portableDir\アプデ実行.bat" -Value $updateBatContent -Encoding UTF8 -ErrorAction SilentlyContinue

    Write-Host "✅ 配布フォルダ ($portableDir) への反映が完了しました！" -ForegroundColor Green
} else {
    Write-Host "✅ distフォルダ内にビルドが完了しました:" -ForegroundColor Green
    if (Test-Path "dist\chrome_launcher.exe") {
        Write-Host "   - $(Join-Path (Get-Location) 'dist\chrome_launcher.exe')" -ForegroundColor Green
    }
    if (Test-Path "$distDir\X_AutoTweetBot.exe") {
        Write-Host "   - $(Join-Path (Get-Location) "$distDir\X_AutoTweetBot.exe")" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "========================================================" -ForegroundColor Green
Write-Host "【アップデート完了】全機能の統合更新が完了しました！" -ForegroundColor Green
Write-Host "1. 【初回手動ログイン】chrome_launcher.exe を追加（サイト勝手に開かず保持）" -ForegroundColor Yellow
Write-Host "2. 【括弧付きフォルダ対応】X_BOT_配布用(Python不要) での起動エラー対策済み" -ForegroundColor Yellow
Write-Host "3. 【案3: 画像クールダウン】同一垢30日間禁止 ＆ 他垢7日間共有禁止" -ForegroundColor Yellow
Write-Host "4. 【日常ツイート100個＆保護】アップロード通信保護 ＋ 4枠フォルダ自動選出" -ForegroundColor Yellow
Write-Host "========================================================" -ForegroundColor Green
Write-Host ""
Read-Host "Enterキーを押すとウィンドウを閉じます"
