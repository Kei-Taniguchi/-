# Codespaces preview recovery

設定変更後に既存のCodespaceでプレビューが開かない場合は、コンテナ設定を再適用する必要があります。

1. Codespaceを開く
2. Command Paletteから **Codespaces: Rebuild Container** を実行
3. 再構築完了まで待つ
4. 自動的にポート8080が転送され、ブラウザプレビューが開く

ポート一覧に `8080` が出ている場合は、その行の「Open in Browser」でも確認できます。
