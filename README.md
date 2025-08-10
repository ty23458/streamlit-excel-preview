# Streamlit Excel Preview

Excel/CSVファイルをアップロードして中身をプレビューできるシンプルなStreamlitアプリです。
- シート選択（Excel）
- 列範囲 `usecols` 指定（例: `A:D` または `A,C,E`）
- ヘッダー行指定、スキップ行、最大表示行数
- 列型の確認、プレビューのCSVダウンロード

## ローカルで動かす
```bash
pip install -r requirements.txt
streamlit run app.py
```

## デプロイ（Streamlit Community Cloud）
1. GitHubにこのフォルダの内容をアップロード（Publicリポ）
2. Streamlit Community Cloudで「New app」→ リポジトリと `app.py` を指定して Deploy
3. 完了すると専用URLが発行されます

## 注意
- `.xlsx` は `openpyxl`、`.xls` は `xlrd`、`.xlsb` は `pyxlsb` で読み込みます
- CSVの場合は `usecols` 指定は無視（pandasのExcel専用指定のため）
