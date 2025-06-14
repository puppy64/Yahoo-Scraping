# Yahoo-Scraping

このリポジトリには、Yahoo!ファイナンスから投資信託の月末基準価額を取得してExcelに記録するスクリプト `TestJP_5.py` が含まれています。

## 必要環境
- Python 3.8 以上
- Google Chrome および ChromeDriver
- `selenium`、`openpyxl` パッケージ

依存パッケージは以下でインストールできます。

```bash
pip install selenium openpyxl
```

ChromeDriver は Chrome のバージョンに合わせてダウンロードし、`PATH` が通った場所に配置してください。

## 使い方
1. `JP_Stocks.xlsx` をリポジトリのルートに配置します。"Price" シートの A 列に証券コード、1 行目の B 列以降に `yy/mm` 形式の年月を記入してください。
2. スクリプトを実行します。

```bash
python TestJP_5.py
```

処理が完了すると Excel ファイルに月末(休日の場合は直前の平日)の基準価額が書き込まれ、上書き保存されます。

## 注意
- Web からデータを取得するため、実行にはインターネット接続が必要です。
- Yahoo!ファイナンスのページ構成変更により動作しなくなる可能性があります。
