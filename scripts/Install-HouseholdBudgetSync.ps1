# OneDriveの家計.xlsmを5分ごとにGitHubへ同期するタスクを登録します。
# このスクリプトは、リポジトリをWindows PCへcloneした後に1回だけ実行します。

$ErrorActionPreference = "Stop"
$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$SyncScript = Join-Path $RepoRoot "scripts\Sync-HouseholdBudget.ps1"
$TaskName = "HouseholdBudget-OneDrive-Sync"

if (-not (Test-Path -LiteralPath $SyncScript)) {
    throw "同期スクリプトが見つかりません: $SyncScript"
}

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    throw "Git が見つかりません。Git for Windowsをインストールしてください。"
}

if (-not (Get-Command py -ErrorAction SilentlyContinue) -and -not (Get-Command python -ErrorAction SilentlyContinue)) {
    throw "Python 3 が見つかりません。Pythonをインストールしてください。"
}

# 必要なPythonライブラリを現在のPython環境へ入れます。
if (Get-Command py -ErrorAction SilentlyContinue) {
    & py -3 -m pip install -r (Join-Path $RepoRoot "requirements-local.txt")
} else {
    & python -m pip install -r (Join-Path $RepoRoot "requirements-local.txt")
}
if ($LASTEXITCODE -ne 0) {
    throw "Python依存関係のインストールに失敗しました。"
}

# ログオン中のユーザーとして5分間隔で実行します。
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$SyncScript`""
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 5)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Description "OneDriveの家計.xlsmを家計簿Webアプリへ自動反映" -Force | Out-Null

Write-Host "自動同期を登録しました: $TaskName" -ForegroundColor Green
Write-Host "5分ごとに C:\Users\syrup\OneDrive\デスクトップ\家計.xlsm を確認します。"
Write-Host "停止する場合: Unregister-ScheduledTask -TaskName '$TaskName' -Confirm:`$false"
