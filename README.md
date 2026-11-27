# ExcelをHTMLで読むツール

表データ（.xlsx / .csv / .tsv / .json）を読みやすく表示・点検するためのツール集です。
単一HTMLファイルで動く Web 版、コマンドラインで使う Python 版、API として使う Flask バックエンド版の3つが入っています。
いずれも「読む前に全体像と要注意ポイントが分かる」概要プロファイル（列ごとの型・空欄率・ユニーク数・最小/最大/平均・最頻値、空欄セル・重複行・型混在・外れ値の検出）を備えています。

## 構成

| パス | 内容 |
|---|---|
| `ExcelをHTMLで読む.html` | Web版（単一HTMLファイル、完全オフライン動作） |
| `python/profile_table.py` | Python CLI版（プロファイルを JSON / HTML で出力） |
| `backend/app.py` | Flask バックエンド版（アップロード→プロファイルJSONを返すAPI） |

## Web版

HTMLファイルをブラウザで開き、表データのファイルを画面にドラッグ＆ドロップするだけです。
処理はすべてブラウザ内で完結し、ファイルが外部に送信されることはありません。サンプルデータも内蔵しています。

主な機能:

- シート切り替え／見出し行の選択／検索・並び替え・数値列の合計と平均
- 概要ダッシュボード（シート一覧サマリ、列プロファイル、要注意ポイントの表示と該当セルへのジャンプ）
- データチェック（重複・空欄・型混在の検出）
- 横スクロール対策（列固定・列の絞り込み・転置表示・カード表示）
- 画像・テキストボックス・図形の表示（埋め込み画像のサムネイル一覧と拡大表示、テキストボックスの本文抽出、配置セルのマーカー表示）
- 音声読み上げ（Web Speech API）
- HTML・CSVへの書き出し
- 概要プロファイルのJSON書き出し（Python CLI版と同一構造で、結果を共有できます）

対応形式: .xlsx / .csv / .tsv（文字コードは UTF-8 / Shift_JIS を自動判定、区切り文字はカンマ／セミコロン／タブを内容から自動判定）/ .json（オブジェクトの配列形式）

## Python CLI版

```bash
pip install openpyxl   # xlsxを読む場合のみ（csv/tsv/jsonは標準ライブラリのみ）
python3 python/profile_table.py 資料.xlsx                 # JSONで出力
python3 python/profile_table.py 資料.csv --format html --out profile.html

# 一括（複数ファイル／フォルダ）：合計サマリ付きの1つのJSON/HTMLを出力
python3 python/profile_table.py 資料1.xlsx 資料2.csv --out batch.json
python3 python/profile_table.py 資料フォルダ/ --format html --out batch.html
```

一括では、読めないファイルが混ざっていてもそのファイルは error 付きで記録して残りの処理を続ける（全件読めなかった場合のみ終了コード1）。

## バックエンド版

```bash
cd backend
pip install -r requirements.txt
python3 app.py   # http://127.0.0.1:5000
```

- `GET /` : 動作確認用のアップロードページ
- `GET /api/health` : 死活確認（`{"status": "ok"}` を返します）
- `POST /api/profile` : multipart/form-data（フィールド名 `file`）でファイルを受け取り、プロファイルJSONを返します
- `POST /api/profiles` : multipart/form-data（フィールド名 `files`、複数可）で複数ファイルを受け取り、`{"files": [...], "summary": {...}}` を返します。読めないファイルはエラーとして記録し、残りの処理を続けます

## 注意

収録されているサンプルはすべてダミーデータです。実データの取り扱いは各ツールの処理方式（Web版はブラウザ内完結、CLI版・バックエンド版はローカル実行）を確認のうえ自己責任でお願いします。

## ライセンス

MIT License
