#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
發佈打包　make_release.py

把工具資料夾打包成 Release 用的 zip，並自動驗證。

    python scripts/make_release.py Pseudonymiser_v3.0_NCCUORD

檔名規則（每次改版都照這個）
----------------------------
    Release 附件與最外層資料夾一律用英文：

        Pseudonymiser_v<版號>_NCCUORD.zip
        └── Pseudonymiser_v<版號>_NCCUORD/
            └── 假名化工具/          ← 裡面維持中文，同仁看得懂

    GitHub 的下載網址會把中文檔名轉成一長串 %E5%81... 的編碼，
    部分瀏覽器與郵件系統會再轉一次，同仁拿到的檔名就變成亂碼。
    最外層用英文可以避開這一段；工具內部的中文檔名不受影響
    （程式會去讀「使用說明.txt」「欄位設定.txt」等名稱，不能改）。

    這支腳本會擋下非 ASCII 的輸出檔名。

為什麼要用這支而不是直接壓縮
--------------------------------
Linux 與 macOS 的 `zip` 指令，對非 ASCII 檔名不會設定
general purpose flag 的第 11 位元（0x800）。繁體中文 Windows 因此改用
Big5 解讀 UTF-8 檔名，檔案總管回報「壓縮資料夾無效」，內建 tar 回報
`Invalid empty pathname`，整包解不開——v2.9 的 Release 就是這樣。

Python 的 zipfile 會自動替非 ASCII 檔名設定旗標，所以改用它打包。

用法
----
    # 在 repo 根目錄
    python scripts/make_release.py <發佈資料夾名稱> [輸出檔名]

    發佈資料夾要放在 repo 的上一層，結構是
        <發佈資料夾名稱>/假名化工具/...
"""

import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# 絕對不可以出現在發佈包裡
FORBIDDEN_NAMES = {"salt.txt", "mapping.csv", "aliases.csv",
                   "name_warnings.csv", "retired_codes.csv",
                   "output_sources.csv", "last_project.txt"}
FORBIDDEN_DIRS = {"_private", "output", "__pycache__", ".git"}


def check_tree(folder: Path):
    bad = []
    for p in folder.rglob("*"):
        if not p.is_file():
            continue
        rel = p.relative_to(folder)
        if p.name in FORBIDDEN_NAMES:
            bad.append(f"{rel}　（對照表或狀態檔，不可發佈）")
        if set(rel.parts[:-1]) & FORBIDDEN_DIRS:
            bad.append(f"{rel}　（位於不該發佈的資料夾）")
        if p.name == ".DS_Store":
            bad.append(f"{rel}　（macOS 殘留檔）")
    return bad


def verify_zip(zip_path: Path):
    """回傳問題清單；空清單代表通過。"""
    problems = []
    with zipfile.ZipFile(zip_path) as z:
        infos = z.infolist()
        noflag = [i.filename for i in infos
                  if not i.filename.isascii() and not i.flag_bits & 0x800]
        if noflag:
            problems.append(f"{len(noflag)} 個非 ASCII 檔名沒有標示 UTF-8"
                            "（繁中 Windows 會解不開）")
        broken = z.testzip()
        if broken:
            problems.append(f"CRC 檢查失敗：{broken}")
        for i in infos:
            if Path(i.filename).name in FORBIDDEN_NAMES:
                problems.append(f"包含不該發佈的檔案：{i.filename}")
    return problems


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1

    folder_name = sys.argv[1]
    src = (ROOT.parent / folder_name).resolve()
    if not src.is_dir():
        print(f"找不到發佈資料夾：{src}")
        print("請先把要發佈的檔案整理到 repo 上一層的這個資料夾裡。")
        return 1

    out_name = sys.argv[2] if len(sys.argv) > 2 else folder_name
    # 不能用 with_suffix("")：版號裡的點會被當成副檔名（v3.0_NCCUORD → v3）
    if out_name.lower().endswith(".zip"):
        out_name = out_name[:-4]
    if not out_name.isascii():
        print(f"輸出檔名含中文：{out_name}.zip")
        print("Release 附件請用英文檔名，例如 Pseudonymiser_v3.0_NCCUORD.zip")
        print("（GitHub 的下載網址會把中文轉成 %E5%81… 編碼，同仁拿到的檔名容易變亂碼）")
        return 1
    out = ROOT.parent / out_name

    print(f"發佈資料夾：{src}")
    bad = check_tree(src)
    if bad:
        print("\n不通過，以下項目不該出現在發佈包裡：")
        for b in bad:
            print("  ・" + b)
        return 1
    n = sum(1 for p in src.rglob("*") if p.is_file())
    print(f"檢查通過，共 {n} 個檔案。\n")

    zip_path = Path(shutil.make_archive(str(out), "zip",
                                        root_dir=str(src.parent),
                                        base_dir=src.name))
    print(f"已打包：{zip_path}")

    problems = verify_zip(zip_path)
    print("\n" + "=" * 58)
    if problems:
        print("驗證不通過：")
        for p in problems:
            print("  ・" + p)
        print("=" * 58)
        return 1
    size = zip_path.stat().st_size / 1024
    inner = sum(1 for i in zipfile.ZipFile(zip_path).infolist() if not i.filename.isascii())
    print(f"驗證通過：UTF-8 旗標正確、CRC 正確、沒有對照表　（{size:.0f} KB）")
    print(f"　　外層檔名純英文；內部 {inner} 個中文檔名已正確標示 UTF-8")
    print("可以當成 Release 附件上傳。")
    print("=" * 58)
    return 0


if __name__ == "__main__":
    sys.exit(main())
