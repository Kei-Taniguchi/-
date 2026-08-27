# 家計簿・支払い計画 Webアプリ試作

## GitHub Codespaces

新規CodespaceではNode.js 22、依存関係のインストール、Webサーバー起動、8080番ポートの自動転送を設定します。起動スクリプトは `0.0.0.0:8080` の応答を確認してから完了します。`onAutoForward: openBrowser` により通常はターミナル操作なしでブラウザプレビューを開きます。

既存Codespaceで設定変更が反映されない場合は、Command Paletteから **Codespaces: Rebuild Container** を一度実行してください。その後、Portsの8080から **Open in Browser** で確認できます。起動ログは `/tmp/household-budget-web-app.log` にあります。

## アプリ

リボ・分割・固定費・光熱費・その他の支払いを管理し、実績入力から残額・完済予定・月別必要額・収支を自動計算する試作Webアプリです。
