# 玉川測量設計 現場データ生成アプリ（Web版）

## このアプリは何をするの？
filtered_data.xlsx をアップロードすることで、テンプレートに情報を反映した「現場データ(申請地の所在).xlsx」ファイルを生成します。

## 必要ファイル
- 入力シート生成_app.py（アプリ本体）
- 現場データ_テンプレート.xlsx（出力テンプレート）
- requirements.txt（依存ライブラリ）

## デプロイ手順（Streamlit Cloud）

1. GitHubにこのフォルダをアップ
2. https://streamlit.io/cloud にログイン
3. 「New app」 → GitHubリポジトリを選択
4. 起動ファイルに `入力シート生成_app.py` を指定
5. 「Deploy」ボタンで公開完了！
