# 家計簿・支払い計画 Webアプリ試作

## GitHub Codespaces

新規CodespaceではNode.js 22、依存関係のインストール、Webサーバー起動、8080番ポートの自動転送を設定します。起動スクリプトは `0.0.0.0:8080` の応答を確認してから完了します。`onAutoForward: openBrowser` により通常はターミナル操作なしでブラウザプレビューを開きます。

既存Codespaceで設定変更が反映されない場合は、Command Paletteから **Codespaces: Rebuild Container** を一度実行してください。その後、Portsの8080から **Open in Browser** で確認できます。起動ログは `/tmp/household-budget-web-app.log` にあります。

## アプリ

リボ・分割・固定費・光熱費・その他の支払いを管理し、実績入力から残額・完済予定・月別必要額・収支を自動計算する試作Webアプリです。

## OneDriveの家計.xlsmを自動反映

毎回ファイルをChatGPTへアップロードする必要がないよう、Windows PC上の次のファイルを直接読み込めます。

`C:\Users\syrup\OneDrive\デスクトップ\家計.xlsm`

ローカルPCの処理は次の流れです。

`家計.xlsm` → `scripts/import_budget.py` → `data/budget-data.json` → GitHubへcommit/push → Webアプリに反映

### 初回セットアップ

1. このリポジトリをWindows PCへcloneします。
2. WindowsでPython 3、Gitを用意します。
3. GitHubへpushできるようGitの認証を済ませます（GitHub DesktopまたはGit Credential Managerでも構いません）。
4. リポジトリのPowerShellで次を1回実行します。

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\Install-HouseholdBudgetSync.ps1
```

これで `HouseholdBudget-OneDrive-Sync` というWindowsタスクが登録され、**5分ごと**にOneDriveの `家計.xlsm` を確認します。

変更がなければGitHubへcommit/pushしません。変更があった場合だけ `data/budget-data.json` を更新してGitHubへ反映します。

### 手動テスト

自動タスクを待たずに、リポジトリのPowerShellから次を実行できます。

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\Sync-HouseholdBudget.ps1
```

### 注意

- `家計.xlsm` がExcelやOneDriveの同期処理で一時的に読み取りできない場合、その回はスキップします。
- ローカルPCのGitに未コミットの別変更がある場合、安全のため同期を停止します。
- `2026年7月分` 以降のA/B/C列は収入として扱いません。
- カード・サービスの引落日／締め日は、ブックのI/J/K列ではなく公式HPの情報を使用します。
- 個別契約によって返済日等が異なる商品は、公式HPに掲載されている標準条件を表示します。
