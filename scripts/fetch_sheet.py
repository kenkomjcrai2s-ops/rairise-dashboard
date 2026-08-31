#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Google Drive上のGoogleスプレッドシートを
sheet.xlsx として取得する。

GitHub Actionsでは Workload Identity Federation (WIF) を使って
Google Cloudに認証する。

環境変数:
  SPREADSHEET_ID : 対象スプレッドシートのファイルID

挙動:
  1. xlsx形式で直接エクスポートを試みる。
  2. 失敗した場合はODS形式でエクスポートする。
  3. LibreOffice(headless)でODSをxlsxへ変換する。
"""

import io
import os
import subprocess
import sys

import google.auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly"
]

XLSX_MIME = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

ODS_MIME = (
    "application/vnd.oasis.opendocument.spreadsheet"
)


def get_service():
    """
    GitHub ActionsのWIF認証情報を使用して
    Google Drive APIへ接続する。
    """

    print("[fetch_sheet] Google Cloud認証情報を取得しています。")

    creds, project_id = google.auth.default(
        scopes=SCOPES
    )

    print(
        f"[fetch_sheet] Google Cloud認証に成功しました。"
        f" project={project_id}"
    )

    return build(
        "drive",
        "v3",
        credentials=creds,
    )


def export_file(service, file_id, mime_type, out_path):
    """
    Google Driveのファイルを指定形式でエクスポートする。
    """

    request = service.files().export_media(
        fileId=file_id,
        mimeType=mime_type,
    )

    with io.FileIO(out_path, "wb") as fh:
        downloader = MediaIoBaseDownload(
            fh,
            request,
        )

        done = False

        while not done:
            status, done = downloader.next_chunk()

            if status:
                print(
                    f"[fetch_sheet] ダウンロード中: "
                    f"{int(status.progress() * 100)}%"
                )

    if not os.path.exists(out_path):
        raise RuntimeError(
            f"エクスポート後のファイルが見つかりません: {out_path}"
        )

    file_size = os.path.getsize(out_path)

    if file_size == 0:
        raise RuntimeError(
            f"エクスポートされたファイルが空です: {out_path}"
        )

    print(
        f"[fetch_sheet] エクスポート成功: "
        f"{out_path} ({file_size:,} bytes)"
    )


def convert_ods_to_xlsx():
    """
    LibreOfficeを使用してsheet.odsをsheet.xlsxへ変換する。
    """

    if not os.path.exists("sheet.ods"):
        raise RuntimeError(
            "sheet.ods が存在しません。"
        )

    # 既存ファイルがあれば削除
    if os.path.exists("sheet.xlsx"):
        os.remove("sheet.xlsx")

    result = subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to",
            "xlsx",
            "--outdir",
            ".",
            "sheet.ods",
        ],
        capture_output=True,
        text=True,
    )

    if result.stdout:
        print(result.stdout)

    if result.stderr:
        print(
            result.stderr,
            file=sys.stderr,
        )

    if result.returncode != 0:
        raise RuntimeError(
            "LibreOfficeによるODS→XLSX変換に失敗しました。"
        )

    if not os.path.exists("sheet.xlsx"):
        raise RuntimeError(
            "LibreOffice実行後にsheet.xlsxが作成されませんでした。"
        )

    file_size = os.path.getsize("sheet.xlsx")

    if file_size == 0:
        raise RuntimeError(
            "作成されたsheet.xlsxが空です。"
        )

    print(
        f"[fetch_sheet] ODS→XLSX変換成功: "
        f"sheet.xlsx ({file_size:,} bytes)"
    )


def main():
    file_id = os.environ.get(
        "SPREADSHEET_ID",
        ""
    ).strip()

    if not file_id:
        print(
            "[fetch_sheet] エラー: "
            "SPREADSHEET_ID が設定されていません。",
            file=sys.stderr,
        )
        sys.exit(1)

    print(
        f"[fetch_sheet] 対象スプレッドシートID: {file_id}"
    )

    try:
        service = get_service()

    except Exception as e:
        print(
            "[fetch_sheet] Google Cloud認証に失敗しました。",
            file=sys.stderr,
        )
        print(
            f"[fetch_sheet] {type(e).__name__}: {e}",
            file=sys.stderr,
        )
        sys.exit(1)

    # --------------------------------------------------
    # 1. XLSX直接エクスポート
    # --------------------------------------------------

    try:
        if os.path.exists("sheet.xlsx"):
            os.remove("sheet.xlsx")

        print(
            "[fetch_sheet] XLSX直接エクスポートを開始します。"
        )

        export_file(
            service,
            file_id,
            XLSX_MIME,
            "sheet.xlsx",
        )

        print(
            "[fetch_sheet] "
            "XLSX直接エクスポートに成功しました。"
        )

        return

    except Exception as e:
        print(
            "[fetch_sheet] "
            "XLSX直接エクスポートに失敗しました。",
            file=sys.stderr,
        )

        print(
            f"[fetch_sheet] {type(e).__name__}: {e}",
            file=sys.stderr,
        )

        print(
            "[fetch_sheet] "
            "ODS経由のフォールバックを試みます。",
            file=sys.stderr,
        )

    # --------------------------------------------------
    # 2. ODSエクスポート
    # --------------------------------------------------

    try:
        if os.path.exists("sheet.ods"):
            os.remove("sheet.ods")

        print(
            "[fetch_sheet] ODSエクスポートを開始します。"
        )

        export_file(
            service,
            file_id,
            ODS_MIME,
            "sheet.ods",
        )

    except Exception as e:
        print(
            "[fetch_sheet] "
            "ODSエクスポートにも失敗しました。",
            file=sys.stderr,
        )

        print(
            f"[fetch_sheet] {type(e).__name__}: {e}",
            file=sys.stderr,
        )

        sys.exit(1)

    # --------------------------------------------------
    # 3. ODS → XLSX
    # --------------------------------------------------

    try:
        convert_ods_to_xlsx()

    except Exception as e:
        print(
            "[fetch_sheet] "
            "ODS→XLSX変換に失敗しました。",
            file=sys.stderr,
        )

        print(
            f"[fetch_sheet] {type(e).__name__}: {e}",
            file=sys.stderr,
        )

        sys.exit(1)

    print(
        "[fetch_sheet] "
        "スプレッドシートの取得が完了しました。"
    )


if __name__ == "__main__":
    main()
