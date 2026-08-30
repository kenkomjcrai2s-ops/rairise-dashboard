#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Google Drive上の「らいらいず個人成績管理表」スプレッドシートを sheet.xlsx として
カレントディレクトリに取得する。GitHub Actions上での実行を想定し、認証には
サービスアカウントを使う（人手でのPAT提示は不要）。

環境変数:
  GDRIVE_SA_JSON   : サービスアカウントのJSON鍵の中身（文字列そのまま）
  SPREADSHEET_ID   : 対象スプレッドシートのファイルID

挙動:
  1. まず xlsx 形式で直接エクスポートを試みる。
  2. スプレッドシートが大きすぎて失敗した場合（Google Sheetsのエクスポート上限に
     ひっかかるケース）、ODS形式でエクスポートし、LibreOffice(headless)で
     xlsxに変換するフォールバックを行う。
     ※ このフォールバックを使った場合、"一覧"シートのINDIRECT数式のキャッシュが
     失われ空欄になることがあるが、これは generate_dashboard.py 側の
     フォールバック処理（個人タブのC4/C12から直接補完）で吸収される想定。
"""
import os
import sys
import json
import subprocess
import io

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
ODS_MIME = "application/vnd.oasis.opendocument.spreadsheet"


def get_service():
    sa_json = os.environ.get("GDRIVE_SA_JSON")
    if not sa_json:
        print("GDRIVE_SA_JSON が設定されていません。", file=sys.stderr)
        sys.exit(1)
    sa_info = json.loads(sa_json)
    creds = service_account.Credentials.from_service_account_info(sa_info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def export_file(service, file_id, mime_type, out_path):
    request = service.files().export_media(fileId=file_id, mimeType=mime_type)
    fh = io.FileIO(out_path, "wb")
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.close()


def main():
    file_id = os.environ.get("SPREADSHEET_ID")
    if not file_id:
        print("SPREADSHEET_ID が設定されていません。", file=sys.stderr)
        sys.exit(1)

    service = get_service()

    try:
        export_file(service, file_id, XLSX_MIME, "sheet.xlsx")
        print("[fetch_sheet] xlsx直接エクスポートに成功しました。")
        return
    except Exception as e:
        print(f"[fetch_sheet] xlsx直接エクスポート失敗: {e}", file=sys.stderr)
        print("[fetch_sheet] ODS経由のフォールバックを試みます。", file=sys.stderr)

    export_file(service, file_id, ODS_MIME, "sheet.ods")
    result = subprocess.run(
        ["soffice", "--headless", "--convert-to", "xlsx", "sheet.ods"],
        capture_output=True, text=True,
    )
    print(result.stdout)
    if result.returncode != 0 or not os.path.exists("sheet.xlsx"):
        print("[fetch_sheet] ODS→xlsx変換に失敗しました。", file=sys.stderr)
        print(result.stderr, file=sys.stderr)
        sys.exit(1)

    print("[fetch_sheet] ODS経由のエクスポート・変換に成功しました。")


if __name__ == "__main__":
    main()
