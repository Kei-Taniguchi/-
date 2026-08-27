# Codespaces プレビュー

このリポジトリは静的なHTML/CSS/JavaScriptアプリですが、CodespacesではNode.jsの `http-server` で配信します。

## 自動設定

- Node.js 22 Bookworm
- `npm install` をCodespace作成時に実行
- Codespace起動時に `.devcontainer/start-server.sh` を実行
- `0.0.0.0:8080` でWebサーバーを起動
- ポート8080を自動転送
- HTTPとして認識し、転送完了時にブラウザを自動オープン

## 既存Codespaceで設定が反映されない場合

`.devcontainer/devcontainer.json` はコンテナ作成時の設定です。既存Codespaceを使用している場合は、Codespacesの「Rebuild Container（コンテナーの再ビルド）」を一度実行してください。新規Codespaceでは自動的に適用されます。

## ブラウザが開かない場合

ポート一覧で `8080` / `家計簿 Webアプリ` が表示されれば、そこから「ブラウザーで開く」を選択できます。サーバー起動に失敗した場合は `/tmp/household-budget-web-app.log` に原因が記録されます。
