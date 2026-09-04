# OneDrive上の家計.xlsmを読み込み、budget-data.jsonを更新してGitHubへ反映します。
# 初回だけGitHub認証（GitHub CLI）を済ませてください。

$ErrorActionPreference = "Stop"
$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot ".." )).Path
$SourceFile = "C:\Users\syrup\OneDrive\デスクトップ\家計.xlsm"
$OutputFile = Join-Path $RepoRoot "data\budget-data.json"
$ImportScript = Join-Path $RepoRoot "scripts\import_budget.py"

Set-Location $RepoRoot

if (-not (Test-Path -LiteralPath $SourceFile)) {
    Write-Host "家計.xlsm が見つかりません: $SourceFile" -ForegroundColor Yellow
    exit 2
}

# ExcelがOneDrive同期中でも読み取り専用で処理できるよう、まずファイルを開けるか確認します。
try {
    $stream = [System.IO.File]::Open($SourceFile, 'Open', 'Read', 'ReadWrite')
    $stream.Close()
    $stream.Dispose()
} catch {
    Write-Host "家計.xlsm を読み取れません。Excel/OneDriveの処理が終わってから再実行します。" -ForegroundColor Yellow
    exit 3
}

$beforeHash = if (Test-Path -LiteralPath $OutputFile) { (Get-FileHash -Algorithm SHA256 -LiteralPath $OutputFile).Hash } else { "" }

$env:HOUSEHOLD_BUDGET_SOURCE = $SourceFile

# WindowsのPython Launcherを優先し、無ければpythonを使用します。
if (Get-Command py -ErrorAction SilentlyContinue) {
    & py -3 $ImportScript
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    & python $ImportScript
} else {
    throw "Python 3 が見つかりません。Pythonをインストールしてください。"
}

if ($LASTEXITCODE -ne 0) {
    throw "家計データの読み込みに失敗しました。終了コード: $LASTEXITCODE"
}

if (-not (Test-Path -LiteralPath $OutputFile)) {
    throw "budget-data.json が生成されませんでした。"
}

afterHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $OutputFile).Hash
if ($beforeHash -eq $afterHash) {
    Write-Host "変更なし。GitHubへの更新は不要です。" -ForegroundColor Green
    exit 0
}

# Codespaces側などで変更したファイルがある場合に上書きしないよう確認します。
$status = git status --porcelain
if ($status) {
    $nonBudgetChanges = $status | Where-Object { $_ -notmatch "^\s*[MADRCU?!]{1,2}\s+data/budget-data\.json$" }
    if ($nonBudgetChanges) {
        Write-Host "未コミットの別変更があるため、安全のため停止します。" -ForegroundColor Yellow
        $nonBudgetChanges | ForEach-Object { Write-Host $_ }
        exit 4
    }
}

git pull --rebase origin main
if ($LASTEXITCODE -ne 0) {
    throw "GitHubからの最新取得に失敗しました。"
}

git add -- data/budget-data.json

git diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
    Write-Host "変更なし。GitHubへの更新は不要です。" -ForegroundColor Green
    exit 0
}

git commit -m "Update budget data from OneDrive"
if ($LASTEXITCODE -ne 0) {
    throw "コミットに失敗しました。"
}

git push origin main
if ($LASTEXITCODE -ne 0) {
    throw "GitHubへのpushに失敗しました。GitHub認証を確認してください。"
}

Write-Host "家計.xlsm → budget-data.json → GitHub の更新が完了しました。" -ForegroundColor Green
