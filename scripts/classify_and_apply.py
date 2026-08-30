#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_dashboard.py を実行し、コンソールに出る「確認事項」を
SKILL.md記載のルールで分類する:

  - "[自動補完・要確認不要]" で始まるメッセージ           → ブロックしない
  - それ以外の確認事項（段位ラベル不一致・個人タブ欠落等） → ブロックする（要レビュー）
  - スクリプトがエラー終了                                → ブロックする
  - 人数が前回から大きく増減（目安: ±2名を超える）         → ブロックする
  - 生成後の人数が0名                                      → ブロックする

判定結果は GITHUB_OUTPUT に need_review=true/false として書き出し、
理由は review_reasons.txt に書き出す（PR本文・ログ用）。
"""
import subprocess
import sys
import os
import json
import re

XLSX = "sheet.xlsx"
OUT_HTML = "index.new.html"
PREV_HTML = "index.html"


def load_players(path):
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as f:
        c = f.read()
    m = re.search(r'<script id="rawData"[^>]*>(.*?)</script>', c, re.S)
    if not m:
        return []
    try:
        return json.loads(m.group(1)).get("players", [])
    except Exception:
        return []


def main():
    reasons = []

    proc = subprocess.run(
        [sys.executable, "generate_dashboard.py", XLSX, OUT_HTML],
        capture_output=True, text=True,
    )
    print(proc.stdout)
    if proc.stderr:
        print(proc.stderr, file=sys.stderr)

    if proc.returncode != 0:
        reasons.append("generate_dashboard.py がエラーで終了しました（ログを確認してください）。")
    else:
        lines = proc.stdout.splitlines()
        confirm_lines = []
        in_section = False
        for line in lines:
            if line.strip() == "--- 確認事項 ---":
                in_section = True
                continue
            if in_section and line.startswith("・"):
                confirm_lines.append(line[1:])

        for line in confirm_lines:
            if not line.startswith("[自動補完・要確認不要]"):
                reasons.append(line)

        old_players = load_players(PREV_HTML)
        new_players = load_players(OUT_HTML)
        old_n, new_n = len(old_players), len(new_players)

        if new_n == 0:
            reasons.append("生成後の人数が0名でした。生成物が壊れている可能性があります。")
        elif old_n > 0:
            delta = new_n - old_n
            if abs(delta) > 2:
                reasons.append(
                    f"人数が前回({old_n}名)から{delta:+d}名変化しており、"
                    f"通常運用の増減幅(±2名)を超えています。"
                )

    need_review = len(reasons) > 0

    with open("review_reasons.txt", "w", encoding="utf-8") as f:
        if reasons:
            f.write("以下の確認事項があるため、mainへの直接反映はせずPRを作成しました。\n\n")
            for r in reasons:
                f.write(f"- {r}\n")
        else:
            f.write("確認事項はありませんでした（自動補完のみ、または該当なし）。\n")

    gh_output = os.environ.get("GITHUB_OUTPUT")
    if gh_output:
        with open(gh_output, "a", encoding="utf-8") as f:
            f.write(f"need_review={'true' if need_review else 'false'}\n")

    print("=== 判定結果 ===")
    print("need_review:", need_review)
    for r in reasons:
        print(" -", r)


if __name__ == "__main__":
    main()
