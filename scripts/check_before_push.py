#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
上傳前安全檢查　check_before_push.py

在 git push 之前跑一次。檢查工作目錄裡有沒有東西不該被推上公開網路：
對照表、鹽值、真實資料、殘留的個資。

    python scripts/check_before_push.py

有任何一項不通過就會以離開碼 1 結束，訊息會說明是哪一個檔案、為什麼。
"""

import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# 一旦出現在版控裡就是事故
FATAL_NAMES = {
    "salt.txt": "鹽值。外流等於所有代碼都可以被反推。",
    "mapping.csv": "代碼與真實姓名的對照表。",
    "aliases.csv": "姓名整併規則，含真實姓名。",
    "name_warnings.csv": "姓名寫法異常紀錄，含真實姓名。",
    "retired_codes.csv": "退役代碼對照，含真實姓名。",
    "output_sources.csv": "輸出來源對照，可能含檔名層級的個資。",
}
FATAL_DIRS = {
    "_private": "對照表與鹽值所在的資料夾。",
    "output": "假名化後的產出，不需要進版控。",
}

# 只有示範專案的 input 可以進版控
ALLOWED_INPUT = "專案/示範專案/input/"

# 疑似個資的樣式（用於掃描文字檔）
PII = [
    (re.compile(r"\b[A-Z][12]\d{8}\b"), "疑似身分證字號"),
    (re.compile(r"\b09\d{2}[- ]?\d{3}[- ]?\d{3}\b"), "疑似手機號碼"),
    (re.compile(r"[\w.+-]+@(?!example\.)[\w-]+\.[\w.]+"), "疑似 Email"),
]
TEXT_EXT = {".py", ".txt", ".md", ".csv", ".bat", ".json", ".yml", ".yaml"}


def tracked_files():
    try:
        out = subprocess.run(["git", "-c", "core.quotepath=false", "ls-files"], cwd=ROOT,
                             capture_output=True, text=True, encoding="utf-8", check=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        return None
    return [p for p in out.stdout.splitlines() if p.strip()]


def drop_ignored(paths):
    """把 .gitignore 已經擋掉的路徑濾掉，避免重複報同一件事。"""
    if not paths:
        return paths
    try:
        out = subprocess.run(["git", "-c", "core.quotepath=false",
                                "check-ignore", "--stdin"], cwd=ROOT,
                             input="\n".join(paths), capture_output=True, text=True, encoding="utf-8")
    except FileNotFoundError:
        return paths
    ignored = set(out.stdout.splitlines())
    return [p for p in paths if p not in ignored]


def check_bat_cjk_adjacency():
    """.bat 裡相鄰兩行都含中文時，cmd 在真實主控台會把後一行從行中間開始讀，
    那一行說明就不見了，換成一行 is not recognized 錯誤。中間空一行即可避免。
    （行尾是 ^ 的續行不能插空行，會把指令切斷，所以跳過。）"""
    bad = []
    for p in sorted(ROOT.rglob("*.bat")):
        try:
            lines = p.read_bytes().decode("utf-8").split("\r\n")
        except UnicodeDecodeError:
            continue
        for i in range(len(lines) - 1):
            a, b = lines[i], lines[i + 1]
            if not a.isascii() and not b.isascii() and not a.rstrip().endswith("^"):
                bad.append(f"{p.relative_to(ROOT)}:{i + 1}")
    return bad


def main():
    problems = []
    notes = []

    files = tracked_files()
    if not files:
        # 還沒 git init，或還沒有任何 commit／暫存——此時 git ls-files 是空的。
        # 不能因此判定通過，改為掃描整個資料夾。
        print("! 版控中尚無檔案，改為掃描整個資料夾。\n")
        files = [str(p.relative_to(ROOT)).replace("\\", "/")
                 for p in ROOT.rglob("*") if p.is_file()
                 and not str(p.relative_to(ROOT)).replace("\\", "/").startswith(".git/")]
        files = drop_ignored(files)

    print(f"檢查 {len(files)} 個檔案……\n")

    for rel in files:
        parts = rel.split("/")
        name = parts[-1]

        if name in FATAL_NAMES:
            problems.append(f"[對照表] {rel}\n         {FATAL_NAMES[name]}")
        for d, why in FATAL_DIRS.items():
            if d in parts[:-1]:
                problems.append(f"[資料夾] {rel}\n         位於 {d}/ 之下：{why}")
                break

        if "/input/" in rel and not rel.startswith(ALLOWED_INPUT):
            problems.append(f"[原始檔] {rel}\n         只有「{ALLOWED_INPUT}」底下的示範檔可以進版控。")

        p = ROOT / rel
        if p.suffix.lower() in TEXT_EXT and p.exists():
            try:
                text = p.read_text(encoding="utf-8", errors="ignore")
            except Exception:
                continue
            for rx, label in PII:
                hits = rx.findall(text)
                if hits:
                    notes.append(f"[{label}] {rel}　例：{str(hits[0])[:24]}")

    bat_bad = check_bat_cjk_adjacency()

    print("=" * 62)
    if problems:
        print(f"不通過：{len(problems)} 項必須處理\n")
        for i, s in enumerate(problems, 1):
            print(f"{i:2d}. {s}")
        print("\n處理方式：")
        print("  1. git rm --cached <檔案>        （從版控移除，本機保留）")
        print("  2. 確認 .gitignore 有擋住它")
        print("  3. 若已經 push 過，改鹽值並視為已外洩處理——")
        print("     git 歷史會永久保留，刪掉最新版本沒有用")
        print("=" * 62)
        return 1

    print("通過：沒有發現對照表、鹽值或非示範的原始檔。")
    if bat_bad:
        print(f"\n另有 {len(bat_bad)} 處 .bat 相鄰兩行都含中文，"
              "在真實主控台會吃掉一行說明；中間請空一行：")
        for s2 in bat_bad[:20]:
            print("  ·", s2)
    if notes:
        print(f"\n另有 {len(notes)} 項需要你自己看一眼（不一定是問題）：")
        for s in notes[:20]:
            print("  ·", s)
    print("=" * 62)
    return 0


if __name__ == "__main__":
    sys.exit(main())
