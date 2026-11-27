#!/usr/bin/env python3
"""表データのプロファイルを返す Flask バックエンド。

Web版（単一HTML）はブラウザ内で完結するが、こちらはサーバー側で処理する版。
大きなファイルのバッチ処理や、他システムからのAPI利用を想定している。

起動:
    pip install -r requirements.txt
    python3 app.py            # http://127.0.0.1:5000

API:
    POST /api/profile  (multipart/form-data, フィールド名: file)
        xlsx / csv / tsv / json を受け取り、プロファイルJSONを返す。
    GET  /
        動作確認用の簡易アップロードページ。
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "python"))
from profile_table import profile_file  # noqa: E402

ALLOWED = {".xlsx", ".csv", ".tsv", ".json"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MBまで


@app.get("/")
def index():
    return (
        "<!DOCTYPE html><html lang='ja'><meta charset='utf-8'>"
        "<title>profile API</title>"
        "<h1>表データ プロファイルAPI</h1>"
        "<form action='/api/profile' method='post' enctype='multipart/form-data'>"
        "<input type='file' name='file' accept='.xlsx,.csv,.tsv,.json'>"
        "<button type='submit'>プロファイルを作成</button></form>"
        "<p>対応形式: xlsx / csv / tsv / json（JSONはオブジェクトの配列）</p>"
    )


@app.post("/api/profile")
def api_profile():
    if "file" not in request.files:
        return jsonify({"error": "file フィールドが必要です"}), 400
    upload = request.files["file"]
    suffix = Path(upload.filename or "").suffix.lower()
    if suffix not in ALLOWED:
        return jsonify({"error": f"未対応の拡張子です: {suffix or '(なし)'}（xlsx / csv / tsv / json のみ）"}), 400
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        upload.save(tmp.name)
        tmp_path = tmp.name
    try:
        profile = profile_file(tmp_path)
        profile["file"] = upload.filename
        return jsonify(profile)
    except (ValueError, RuntimeError) as e:
        return jsonify({"error": str(e)}), 422
    finally:
        Path(tmp_path).unlink(missing_ok=True)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
