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
    POST /api/profiles  (multipart/form-data, フィールド名: files を複数)
        複数ファイルを受け取り、{"files": [...], "summary": {...}} を返す。
        読めないファイルは error 付きで記録し、残りの処理を続ける。
    GET  /
        動作確認用の簡易アップロードページ。
"""
from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "python"))
from profile_table import profile_file, summarize_profiles  # noqa: E402

ALLOWED = {".xlsx", ".csv", ".tsv", ".json"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MBまで


def _profile_upload(tmp_path: Path, original_name: str) -> dict:
    """アップロードされたファイルのプロファイルを作る。

    profile["file"] は元のファイル名にし、csv / tsv / json の単一表では
    シート名も元のファイル名（拡張子なし）にする。profile_file は一時
    ファイルのパスから作るため、そのままだとシート名が一時ファイル名
    （tmpXXXX）になってしまう。
    """
    profile = profile_file(tmp_path)
    profile["file"] = original_name
    suffix = Path(original_name).suffix.lower()
    if suffix in (".csv", ".tsv", ".json") and profile.get("sheets"):
        profile["sheets"][0]["name"] = Path(original_name).stem
    return profile


@app.get("/")
def index():
    return (
        "<!DOCTYPE html><html lang='ja'><meta charset='utf-8'>"
        "<title>profile API</title>"
        "<h1>表データ プロファイルAPI</h1>"
        "<form action='/api/profile' method='post' enctype='multipart/form-data'>"
        "<input type='file' name='file' accept='.xlsx,.csv,.tsv,.json'>"
        "<button type='submit'>プロファイルを作成</button></form>"
        "<h2>一括（複数ファイル）</h2>"
        "<form action='/api/profiles' method='post' enctype='multipart/form-data'>"
        "<input type='file' name='files' accept='.xlsx,.csv,.tsv,.json' multiple>"
        "<button type='submit'>一括プロファイルを作成</button></form>"
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
        return jsonify(_profile_upload(Path(tmp_path), upload.filename or tmp_path))
    except Exception as e:  # 壊れたファイル等は種類を問わずJSONエラーで返す
        return jsonify({"error": str(e)}), 422
    finally:
        Path(tmp_path).unlink(missing_ok=True)


@app.post("/api/profiles")
def api_profiles():
    uploads = request.files.getlist("files") or request.files.getlist("file")
    if not uploads:
        return jsonify({"error": "files フィールドが必要です"}), 400
    tmp_paths: list[Path] = []
    names: list[str] = []
    try:
        for upload in uploads:
            suffix = Path(upload.filename or "").suffix.lower()
            if suffix not in ALLOWED:
                tmp_paths.append(Path(""))
                names.append(upload.filename or "(名前なし)")
                continue
            with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                upload.save(tmp.name)
                tmp_paths.append(Path(tmp.name))
                names.append(upload.filename or tmp.name)
        profiles: list[dict] = []
        for tmp_path, name in zip(tmp_paths, names):
            if not tmp_path.name:
                profiles.append({"file": name, "error": "未対応の拡張子です（xlsx / csv / tsv / json のみ）"})
                continue
            try:
                profiles.append(_profile_upload(tmp_path, name))
            except Exception as e:  # openpyxlのBadZipFileなど、読み込み系の例外は種類を問わず記録する
                profiles.append({"file": name, "error": str(e)})
        return jsonify({"files": profiles, "summary": summarize_profiles(profiles)})
    finally:
        for tmp_path in tmp_paths:
            if tmp_path.name:
                tmp_path.unlink(missing_ok=True)


@app.get("/api/health")
def api_health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    # デバッグモードは環境変数 FLASK_DEBUG=1 のときだけ有効にする
    app.run(host="127.0.0.1", port=5000, debug=os.environ.get("FLASK_DEBUG") == "1")
