#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
====================================================================
 政大研發處　新進教師計畫參與分析
 上傳前假名化工具  anonymize.py   v1.4
====================================================================

用途
----
把研發處、人事室、國科會系統匯出的原始檔中的「教師姓名」欄位，
在本機替換成固定長度的代碼（例如「王小明」→「T-3F9A21B8C4」），
再交給外部 AI 進行分析。對照表只留在本機，不隨檔案外流。

特性
----
* 同一個姓名在所有檔案、所有年度都會得到「相同」的代碼，
  因此跨檔比對、跨年追蹤完全不受影響。
* 代碼由「姓名 + 本機鹽值」以 SHA-256 產生，無鹽值無法反推。
* 鹽值只產生一次並存於本機 salt.txt，請勿外流、勿刪除。
  （刪除或更換鹽值，會導致新舊年度的代碼對不起來。）
* 產生對照表 mapping.csv 供內部還原使用，同樣不得外流。

v1.3 修正（依 2026-09 檢測報告）
--------------------------------
1. output 只放「工具真正處理過的檔案」。非試算表格式不再原樣複製進 output。
2. 姓名欄與敏感欄的比對改為不分大小寫，並擴充清單；另可用同目錄的
   「欄位設定.txt」自行增補，不必改程式。
3. 偵測到「疑似姓名／敏感欄但不在清單內」時提出警告。
4. 一張工作表支援多段表頭。
5. 代碼長度 6 → 10 個十六進位字元，並加入碰撞檢查。
6. 輸出檔名衝突偵測（原本同名的 .xls 與 .xlsx 會靜默互相覆蓋）。
7. 寫檔失敗不再中斷整批作業；對照表改為原子寫入。
8. 姓名寫法異常的提醒不再把真實姓名印在畫面上。
9. 標準輸出改為 UTF-8，修正輸出被重導向時的 UnicodeEncodeError。

v1.4 修正（依 2026-09-15 對 v1.3 的兩次驗證報告）
------------------------------------------------
1. 標題列判定：v1.3 只要一格內容等於欄名就把整列當標題列，資料格剛好寫
   「共同主持人」「電話」時整張表零替換。現在第二段以後的標題列要通過多道
   檢查才採用，而且每一格一定再疊上 v1.2 的「單一標題列、以下全部處理」
   判定當保底：任何一種判定認為是姓名欄的格子都換成代碼（包括數字格），
   只有整格等於欄名的格子保留原樣。
2. 兩階段處理：先把整批檔案的姓名欄都換成代碼，再用「整批」的已知姓名
   掃描備註、說明、表頭上方文字、工作表名稱與檔名——結果不受檔案排序影響。
   已知姓名只收「像人名」的值（不收單位、職稱、合計、經費項目）。
3. 認不出任何姓名欄、疑似個資欄位、欄名清單型表格（直式表單、異動紀錄）、
   姓名欄全部沒換到等狀況，一律計入「需要確認」，視窗不會說完成。
4. output 盤點：非 .xlsx 檔、檔名含已知姓名的舊檔、同一原始檔改名後留下的
   舊檔，移到 _private\\output_隔離\\（不刪除）；沒有來源紀錄的舊檔計入需要確認。
5. 對照表在寫出任何 output 之前就存檔；mapping.csv 被 Excel 開著時一開始就停下。
   salt.txt 不見但 mapping.csv 還在時停下，不會靜默產生新鹽值。
6. 檢視模式完全不寫檔（不建立 salt.txt、不改 name_warnings.csv）。
7. 畫面訊息（檔名、工作表名、清單）遮蔽已知姓名與看起來像人名的片段。
8. 欄位設定.txt 寫錯時明確提示並計入結果；接受全形【】、行尾 # 註解與 cp950。
9. 文字裡「承辦人：王小明」「聯絡人 王小明」「王小明老師」這幾種寫法的姓名也換成代碼；
   一格寫好幾個人（王小明、李大華）時個別姓名也配代碼，供第二階段掃描使用。
10. 表頭判定補強：「承辦人｜王小明」資訊列不當標題列（右邊的人名另外換碼）；
   欄名帶簡短括號註記（姓名(英)、電話（公））也認得；有三個以上認得欄名的列，
   上面不必有空白分隔列也能當標題列（前 8 列以外、第二段以後都適用）。

v1.4 第五輪補充（依 2026-09-16 第三次驗證報告）
----------------------------------------------
11. 資料格只有「整格等於欄名」才保留原樣（「講師（王大明）」照樣換碼）；欄名比對先做
    全形半形、簡繁、零寬字元正規化。
12. 保底判定分兩種：v1.2 清單命中的標題列照舊處理到表尾（和 v2.4 處理的格子相同）；
    新清單的保底只到下一段標題列為止，而且那一欄底下要有像姓名的值才採用。
    表頭可含日期、年度欄名；簽核列、單格小標題、段內資料列不當標題列。
13. 文字寫法只換 3～4 字姓名（人員欄名接冒號、空白、括號，或緊接姓名再接「負責、代理」
    這類字；「王小明老師」稱謂前可空一格）。兩字候選、「感謝王小明老師」只提醒。
    已知姓名：三字以上才做句中替換；兩字姓名與單字英文名在句中只提醒。
14. 工作表名、檔名只改「確定像人名」的片段，畫面一律遮蔽；遮蔽後撞名加編號。
15. 停下不處理：鹽值與對照表對不上或有大寫、mapping.csv／retired_codes.csv 被存成 ANSI
    而姓名有「?」、retired_codes.csv 讀寫失敗、output 有代碼但 mapping.csv 不見。
16. output 已有同名檔、但沒有來源紀錄時，覆蓋前先移到 _private\\output_隔離\\；
    mapping.csv 有多餘欄位或不是 UTF-8 時先備份再改寫；一格多人整格代碼、
    「欄名｜姓名」列可能是另一段表頭、欄位設定.txt 放錯位置，都會提醒。

使用方式
--------
    python anonymize.py                 # 處理 ./input 內所有檔案
    python anonymize.py --dry-run       # 只檢視會處理哪些欄位，不寫檔
    python anonymize.py --decode 代碼   # 用對照表把代碼還原成姓名

資料夾結構
----------
    anonymize.py
    input/     ← 把原始檔全部丟進來（子資料夾也可以）
    output/    ← 假名化後的檔案（這個資料夾的內容才可以上傳）
    _private/  ← salt.txt、mapping.csv（絕對不要外流）
====================================================================
"""

import argparse
import csv
import datetime as dt
import hashlib
import os
import re
import secrets
import unicodedata
import sys
from pathlib import Path

try:
    import numpy as np
    import pandas as pd
    import warnings
    warnings.filterwarnings("ignore", message="Cannot parse header or footer.*")
except ImportError:
    sys.exit("請先安裝 pandas 與 openpyxl：\n    pip install pandas openpyxl xlrd")

# 基準資料夾：
#   一般執行時 = anonymize.py 所在資料夾
#   被 PyInstaller 打包成 .exe 時 = .exe 所在資料夾
if getattr(sys, "frozen", False):
    TOOL_ROOT = Path(sys.executable).resolve().parent
else:
    TOOL_ROOT = Path(__file__).resolve().parent

BASE = TOOL_ROOT
IN_DIR = BASE / "input"
OUT_DIR = BASE / "output"
PRIV_DIR = BASE / "_private"
SALT_FILE = PRIV_DIR / "salt.txt"
MAP_FILE = PRIV_DIR / "mapping.csv"
ALIAS_FILE = PRIV_DIR / "aliases.csv"             # 姓名整併表（同一人的不同寫法）
SOURCES_FILE = PRIV_DIR / "output_sources.csv"   # output 檔各是由哪個 input 檔產生
RETIRED_FILE = PRIV_DIR / "retired_codes.csv"    # 被整併掉的舊代碼（仍可還原）

# 使用者可自行增補欄名的設定檔（放在工具資料夾，所有專案共用）
COLUMN_CONFIG = TOOL_ROOT / "欄位設定.txt"

# 整併規則：{錯誤或變體寫法: 正式姓名}。main() 開始時載入。
ALIASES = {}

# Windows 自動產生的系統檔，不算使用者的資料
SYSTEM_FILES = {"desktop.ini", "thumbs.db", ".ds_store"}


def set_workdir(path):
    """切換工作資料夾（多專案用）。

    每個專案有自己的 input / output / _private，因此也有自己的
    salt.txt——不同專案的代碼各自獨立，互不相干。
    視窗介面會在每次執行前呼叫這個函式。
    """
    global BASE, IN_DIR, OUT_DIR, PRIV_DIR, SALT_FILE, MAP_FILE, ALIAS_FILE, SOURCES_FILE, RETIRED_FILE
    BASE = Path(path).resolve()
    IN_DIR = BASE / "input"
    OUT_DIR = BASE / "output"
    PRIV_DIR = BASE / "_private"
    SALT_FILE = PRIV_DIR / "salt.txt"
    MAP_FILE = PRIV_DIR / "mapping.csv"
    ALIAS_FILE = PRIV_DIR / "aliases.csv"
    SOURCES_FILE = PRIV_DIR / "output_sources.csv"
    RETIRED_FILE = PRIV_DIR / "retired_codes.csv"
    return BASE


# --------------------------------------------------------------------
# 需要假名化的欄位名稱。比對方式為「去除空白並轉小寫後完全相符」。
# 若日後系統改版新增欄位，可加進這個清單，或寫進「欄位設定.txt」。
# --------------------------------------------------------------------
BASE_NAME_COLUMNS = frozenset({
    "姓名", "中文姓名", "英文姓名", "教師姓名", "教師中文姓名", "員工姓名",
    "主持人", "計畫主持人", "主持人姓名", "計畫主持人姓名",
    "共同主持人", "共同主持人姓名", "協同主持人", "協同主持人姓名",
    "申請人", "申請人姓名",
    "指導教授", "指導教授姓名", "指導老師", "指導老師姓名",
    "學生姓名", "研究生姓名",
    "承辦人", "經辦人", "填表人", "聯絡人", "聯絡人姓名", "緊急聯絡人", "連絡人", "連絡人姓名",
    "教師", "老師", "授課教師", "講師", "講師姓名", "業師", "業師姓名",
    "作者", "作者姓名", "第一作者", "通訊作者", "共同作者", "撰稿人",
    "審查委員", "評審委員", "委員姓名",
    "負責人", "計畫負責人", "執行人", "計畫執行人",
    "研究人員", "研究人員姓名", "參與人員", "團隊成員", "成員姓名",
    "兼任助理", "專任助理", "助理姓名",
    "得獎人", "獲獎人", "受獎人", "受訪者", "發明人",
    "全名", "真實姓名", "本名",
})

# 這些欄位含個人識別資訊但分析用不到，直接清空
BASE_DROP_COLUMNS = frozenset({
    "員工編號", "職員編號", "人事編號", "教職員編號",
    "申請人代號", "教師代號", "人員代號",
    "身分證字號", "身份證字號", "身分證號", "身份證號", "身分證統一編號", "統一編號",
    "護照號碼", "居留證號", "居留證號碼",
    "電子郵件", "電子信箱", "e-mail", "email", "mail", "電郵",
    "聯絡電話", "電話", "連絡電話", "手機", "手機號碼", "行動電話", "分機",
    "住家電話", "辦公室電話", "公務電話", "聯絡手機", "緊急聯絡電話", "傳真",
    "學號", "學生證號", "生日", "出生日期", "出生年月日",
    "地址", "通訊地址", "戶籍地址", "聯絡地址", "住址", "戶籍地", "居住地址",
    "銀行帳號", "帳號", "郵局帳號", "帳戶",
})

# v1.2（視窗版 v2.4）的欄名清單，只用來重現舊版「選哪一列當標題列」的判斷
LEGACY_NAME_COLUMNS = frozenset({"姓名", "主持人", "申請人", "共同主持人", "協同主持人", "指導教授", "學生姓名"})
LEGACY_DROP_COLUMNS = frozenset({"員工編號", "申請人代號", "身分證字號", "電子郵件", "e-mail", "email",
                                 "聯絡電話", "手機"})

NAME_COLUMNS = set(BASE_NAME_COLUMNS)
DROP_COLUMNS = set(BASE_DROP_COLUMNS)
CONFIG_ADDED = {}          # 欄位設定.txt 加進來的欄名：{比對鍵: 原文}，結束時列出沒用到的
CONFIG_USED = set()

EXCEL_EXT = {".xlsx", ".xls", ".xlsm"}

# 主標題列只在前幾列找（沿用 v1.2 的做法，實務上表頭都在最上面）
MAX_SCAN = 8

# 去空白轉小寫後的比對用集合，load_column_config() 會重算
NAME_KEYS = set()
DROP_KEYS = set()


def norm_header(x) -> str:
    """去掉所有空白（含全形空白）。顯示用。"""
    return re.sub(r"\s", "", str(x))


_ZERO_WIDTH = re.compile(r"[​-‍⁠﻿]")
# 欄名常見的簡體字（「手机」「电子邮件」「联络电话」）；只用在欄名比對，不改資料
_SIMPLIFIED = str.maketrans("机电话邮号码联络职称学证帐户址员师长编单项类别时间备注说明资讯讯", "機電話郵號碼聯絡職稱學證帳戶址員師長編單項類別時間備註說明資訊訊")


_norm_cache = {}


def norm_key(x) -> str:
    """比對用：全形轉半形（ＮＡＭＥ、Ｅ－ｍａｉｌ）、去空白與零寬字元、轉小寫、去掉結尾冒號（「承辦人：」）。"""
    s = x if isinstance(x, str) else str(x)
    r = _norm_cache.get(s)
    if r is None:
        if len(_norm_cache) > 300000:
            _norm_cache.clear()
        t = unicodedata.normalize("NFKC", s).translate(_SIMPLIFIED)
        r = _norm_cache[s] = _ZERO_WIDTH.sub("", re.sub(r"\s", "", t)).lower().rstrip(":")
    return r


def _rebuild_keys():
    global NAME_KEYS, DROP_KEYS
    NAME_KEYS = {norm_key(c) for c in NAME_COLUMNS}
    DROP_KEYS = {norm_key(c) for c in DROP_COLUMNS}
    if "_person_cache" in globals():
        _person_cache.clear()
    for c in ("_core_cache", "_isperson_cache", "_parts_cache"):
        if c in globals():
            globals()[c].clear()


def _read_text_any(path: Path):
    """依序嘗試 UTF-16（有 BOM）、UTF-8、cp950（記事本存成「ANSI」時）。"""
    raw = path.read_bytes()
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16"), "UTF-16"
    try:
        return raw.decode("utf-8-sig"), "UTF-8"
    except UnicodeDecodeError:
        pass
    try:
        return raw.decode("cp950"), "ANSI（cp950）"
    except UnicodeDecodeError:
        return None, None


def load_column_config(verbose: bool = True) -> int:
    """讀取同目錄的「欄位設定.txt」，讓使用者不改程式也能增補欄名。

    格式（# 之後是註解，括號可用半形 [ ] 或全形【 】）：
        [姓名欄]
        指導老師
        [清空欄]
        校內分機

    回傳「有問題、沒有生效」的項目數，由 main() 計入結果讓視窗介面示警。
    """
    global NAME_COLUMNS, DROP_COLUMNS, CONFIG_ADDED
    # 每次都從內建清單重建：視窗介面同一次開啟會重複呼叫，
    # 使用者刪掉設定檔裡的某一行後，不應該還殘留上一次加進去的欄名。
    NAME_COLUMNS = set(BASE_NAME_COLUMNS)
    DROP_COLUMNS = set(BASE_DROP_COLUMNS)
    CONFIG_ADDED = {}
    CONFIG_USED.clear()
    _rebuild_keys()
    problems = []
    # 記事本存檔時常變成「欄位設定.txt.txt」，工具就讀不到
    for odd in TOOL_ROOT.glob("欄位設定.txt.*"):
        problems.append(f"找到「{odd.name}」，但工具只讀「欄位設定.txt」，請把檔名改正")
    if not COLUMN_CONFIG.exists():
        if verbose and problems:
            _print_config_problems(problems)
        return len(problems)

    text, enc = None, None
    try:
        text, enc = _read_text_any(COLUMN_CONFIG)
        if text is None:
            problems.append("檔案編碼無法辨識，請用記事本「另存新檔」並選擇 UTF-8")
    except OSError as e:
        problems.append(f"無法讀取：{e}")

    added_name, added_drop = [], []
    if text is not None:
        section = None
        for no, raw in enumerate(text.splitlines(), 1):
            line = raw.split("#", 1)[0].strip()
            if not line:
                continue
            if line[0] in "[【［":
                if line[-1] not in "]】］":
                    section = None
                    problems.append(f"第 {no} 行「{line[:20]}」少了右括號，底下的欄名不會生效")
                    continue
                tag = norm_header(line[1:-1])
                is_name = "姓名" in tag or "假名" in tag
                is_drop = any(w in tag for w in ("清空", "敏感", "刪除", "個資"))
                if is_name == is_drop:
                    section = None
                    problems.append(f"第 {no} 行的區塊名稱「{line[:20]}」無法辨識，"
                                    "底下的欄名不會生效（請寫成 [姓名欄] 或 [清空欄]）")
                else:
                    section = "name" if is_name else "drop"
                continue
            line = line.rstrip("：:").strip()
            if not line:
                continue
            if section == "name":
                NAME_COLUMNS.add(line)
                added_name.append(line)
            elif section == "drop":
                DROP_COLUMNS.add(line)
                added_drop.append(line)
            else:
                problems.append(f"第 {no} 行「{line[:20]}」不在 [姓名欄] 或 [清空欄] 區塊底下，沒有生效")
                continue
            CONFIG_ADDED[norm_key(line)] = line
        if not (added_name or added_drop):
            problems.append("檔案裡沒有任何一個欄名生效")
    _rebuild_keys()

    if verbose and (added_name or added_drop):
        parts = []
        if added_name:
            parts.append(f"姓名欄 +{len(added_name)}")
        if added_drop:
            parts.append(f"清空欄 +{len(added_drop)}")
        print(f"[欄位設定] 已套用 欄位設定.txt（{'、'.join(parts)}；編碼 {enc}）")
    if verbose and problems:
        _print_config_problems(problems)
    return len(problems)


def _print_config_problems(problems):
    print("  ! 欄位設定.txt 有下列問題，這些設定「沒有生效」：")
    for p in problems[:15]:
        print(f"    ・{p}")
    if len(problems) > 15:
        print(f"    ・…（其餘 {len(problems) - 15} 項）")


_rebuild_keys()


# ====================== 鹽值與對照表 ======================

class StopRun(Exception):
    """一開始就必須停下的狀況（鹽值、對照表有問題）；訊息會原樣印給同仁看。"""


class MappingSaveError(Exception):
    """對照表存不進去（通常是 mapping.csv 正被 Excel 開著）。"""


def load_salt(dry: bool, mapping: dict) -> str:
    if SALT_FILE.exists():
        text, _enc = _read_text_any(SALT_FILE)
        salt = (text or "").strip()
        if not re.fullmatch(r"[0-9a-fA-F]{32}", salt):
            raise StopRun("salt.txt 的內容格式不對（應該是 32 個英數字）。\n"
                          "請用備份還原這個檔案；在還原之前不要執行，否則代碼會全部改變。")
        if salt != salt.lower():
            raise StopRun("salt.txt 裡有大寫英文字母，和工具產生的鹽值不同（工具只產生小寫），可能是抄寫時改到。\n"
                          "請用原始備份檔還原；在還原之前不要執行，否則新加入的人會拿到另一套代碼。")
        _verify_salt(salt, mapping)
        return salt
    if mapping:
        raise StopRun("找不到 salt.txt，但 mapping.csv 裡已經有代碼。\n"
                      "這表示鹽值被刪除或沒有一起複製過來。請先用備份還原 _private\\salt.txt；\n"
                      "若直接執行，會產生一套新的代碼，今年與去年就對不起來。")
    if dry:
        print("[檢視模式] 這個專案還沒有鹽值；正式執行時才會建立。")
        return secrets.token_hex(16)
    PRIV_DIR.mkdir(parents=True, exist_ok=True)
    salt = secrets.token_hex(16)
    SALT_FILE.write_text(salt, encoding="utf-8")
    print(f"[初次執行] 已產生新的鹽值並存於 {SALT_FILE}")
    print("           請妥善保存；更換或遺失會導致新舊年度代碼不一致。")
    return salt


def _verify_salt(salt: str, mapping: dict):
    """用對照表抽樣驗證鹽值：從紙本或備份還原時抄錯一個字，工具要擋下來。"""
    sample = [(c, n) for c, n in list(mapping.items())[:200]][:40]
    if not sample:
        return
    ok = 0
    for code, name in sample:
        full = make_code(name, salt)
        if code == full or (len(code) == 8 and full.startswith(code)):
            ok += 1
    if ok == 0:
        raise StopRun("salt.txt 和 mapping.csv 對不上（用這個鹽值算不出對照表裡的任何一個代碼）。\n"
                      "可能是鹽值抄錯或放錯專案。請確認還原的是這個專案的 salt.txt；在確認之前不要執行。")


MAPPING_ENC = "UTF-8"


def load_mapping() -> dict:
    if not MAP_FILE.exists():
        return {}
    global MAPPING_ENC
    text, enc = _read_text_any(MAP_FILE)
    if text is None:
        raise StopRun("mapping.csv 的編碼無法辨識（可能被 Excel 另存過）。\n"
                      "請用備份還原，或用記事本另存為 UTF-8 後再執行。")
    rows = list(csv.DictReader(text.splitlines()))
    if rows and not {"代碼", "姓名"} <= set(rows[0]):
        raise StopRun("mapping.csv 的欄位不對（應該有「代碼」「姓名」兩欄），可能被改過。\n請用備份還原。")
    out = {r["代碼"]: r["姓名"] for r in rows if r.get("代碼")}
    MAPPING_ENC = enc if not (rows and set(rows[0]) - {"代碼", "姓名"}) else "多餘欄位"
    if enc != "UTF-8" and any("?" in (n or "") for n in out.values()):
        raise StopRun("mapping.csv 被存成 ANSI（Big5）格式，而且有姓名的字變成「?」（Big5 沒有這些字，例如堃、珉、喆）。\n"
                      "繼續執行會讓這些人拿到新代碼、舊代碼永遠查不回。請用 UTF-8 的備份還原 mapping.csv。")
    return out


def _tmp_path(path: Path) -> Path:
    return path.with_name(path.name + ".tmp")


def save_mapping(mapping: dict):
    """原子寫入：先寫暫存檔再置換，避免寫到一半失敗毀掉累積多年的對照表。

    置換失敗時一定刪掉暫存檔——暫存檔裡就是完整的姓名對照。
    """
    PRIV_DIR.mkdir(parents=True, exist_ok=True)
    if MAPPING_ENC != "UTF-8" and MAP_FILE.exists():
        bak = MAP_FILE.with_name(f"mapping.csv.bak_{dt.datetime.now():%Y%m%d_%H%M%S}")
        try:
            bak.write_bytes(MAP_FILE.read_bytes())
        except OSError as e:
            raise MappingSaveError(f"無法備份原本的 mapping.csv：{e}") from e
    tmp = _tmp_path(MAP_FILE)
    try:
        with tmp.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["代碼", "姓名"])
            for code, name in sorted(mapping.items(), key=lambda x: x[1]):
                w.writerow([code, name])
        os.replace(tmp, MAP_FILE)
    except OSError as e:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        raise MappingSaveError(str(e)) from e


def mapping_writable() -> bool:
    """mapping.csv 被 Excel 開著時，Windows 不允許寫入。開始處理前先試一次。"""
    if not MAP_FILE.exists():
        return True
    try:
        with MAP_FILE.open("a", encoding="utf-8"):
            pass
        return True
    except OSError:
        return False


def load_aliases():
    """讀取姓名整併表，回傳 ({變體寫法: 正式姓名}, [形成循環的 (變體寫法, 原本寫的正式姓名)])。

    支援串接：若 A→B 且 B→C，最終會解析成 A→C。
    """
    if not ALIAS_FILE.exists():
        return {}, []
    raw = {}
    text, _enc = _read_text_any(ALIAS_FILE)
    if text is None:
        raise StopRun("aliases.csv（姓名整併表）的編碼無法辨識。請用記事本另存為 UTF-8，或在「姓名整併」重新儲存。")
    reader = csv.DictReader(text.splitlines())
    if reader.fieldnames and not {"變體寫法", "正式姓名"} <= set(reader.fieldnames):
        raise StopRun("aliases.csv（姓名整併表）的欄位不對（應該有「變體寫法」「正式姓名」兩欄）。請在「姓名整併」重新儲存。")
    for r in reader:
        a = (r.get("變體寫法") or "").strip()
        b = (r.get("正式姓名") or "").strip()
        if a and b and a != b:
            raw[a] = b
    out, cycles = {}, []
    for a in raw:
        seen, cur = {a}, raw[a]
        while cur in raw:
            if cur in seen:
                cycles.append((a, raw[a]))
                break
            seen.add(cur)
            cur = raw[cur]
        else:
            out[a] = cur
    return out, cycles


def load_aliases_raw() -> dict:
    """姓名整併表原本寫的規則（整併視窗顯示與儲存用，不做串接解析）。"""
    if not ALIAS_FILE.exists():
        return {}
    text, _enc = _read_text_any(ALIAS_FILE)
    if text is None:
        raise StopRun("aliases.csv（姓名整併表）的編碼無法辨識。請用記事本另存為 UTF-8。")
    raw = {}
    for r in csv.DictReader(text.splitlines()):
        x = (r.get("變體寫法") or "").strip()
        y = (r.get("正式姓名") or "").strip()
        if x and y and x != y:
            raw[x] = y
    return raw


def save_aliases(pairs: dict):
    """pairs = {變體寫法: 正式姓名}"""
    PRIV_DIR.mkdir(parents=True, exist_ok=True)
    with ALIAS_FILE.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["變體寫法", "正式姓名"])
        for a, b in sorted(pairs.items(), key=lambda x: (x[1], x[0])):
            w.writerow([a, b])


def load_retired() -> dict:
    """{被整併掉的舊代碼: 當時的寫法}"""
    if not RETIRED_FILE.exists():
        return {}
    text, _enc = _read_text_any(RETIRED_FILE)
    reader = csv.DictReader((text or "").splitlines())
    if text is None or (reader.fieldnames and not {"代碼", "姓名"} <= set(reader.fieldnames)):
        raise StopRun("retired_codes.csv（被整併掉的舊代碼）讀不到或欄位不對，可能被 Excel 另存過。\n"
                      "請用備份還原；在還原之前不要執行，否則舊代碼可能永久查不回。")
    out = {r["代碼"]: r["姓名"] for r in reader if r.get("代碼")}
    if _enc != "UTF-8" and any("?" in (n or "") for n in out.values()):
        raise StopRun("retired_codes.csv 被存成 ANSI 格式，姓名有字變成「?」。請用 UTF-8 的備份還原。")
    return out


def save_retired(retired: dict):
    """原子寫入；失敗時丟 OSError，呼叫端必須停下（不能在舊代碼從 mapping 移除後才失敗）。"""
    PRIV_DIR.mkdir(parents=True, exist_ok=True)
    tmp = _tmp_path(RETIRED_FILE)
    try:
        with tmp.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["代碼", "姓名"])
            for code, name in sorted(retired.items()):
                w.writerow([code, name])
        os.replace(tmp, RETIRED_FILE)
    except OSError:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        raise


def canonical(name: str) -> str:
    """把變體寫法換成正式姓名；沒有規則就原樣回傳。"""
    return ALIASES.get(name, name)


# 代碼長度（十六進位字元數）。10 碼 = 40 bits。
# 原本為 6 碼（24 bits），數千人規模就會出現兩個人共用同一組代碼。
CODE_LEN = 10
_CODE_RE = re.compile(r"T-[0-9A-F]{6,}")


def make_code(name: str, salt: str) -> str:
    h = hashlib.sha256((salt + "|" + name).encode("utf-8")).hexdigest()
    return "T-" + h[:CODE_LEN].upper()


def assign_code(name: str, salt: str, mapping: dict, stats: dict = None, reserved=()) -> str:
    """配發代碼，並確保不會和「別人」的代碼相同（含被整併掉、仍保留可還原的舊代碼）。

    雜湊本來就有碰撞的可能。撞到時改用加了後綴的輸入重算，直到取得
    一個沒有被別人佔用的代碼；同一個姓名在同一組鹽值下結果仍然固定。
    """
    attempt = 0
    while attempt <= 64:
        seed = name if attempt == 0 else f"{name}\x00{attempt}"
        code = make_code(seed, salt)
        owner = mapping.get(code)
        held = reserved.get(code) if isinstance(reserved, dict) else ("\x00" if code in reserved else None)
        if (owner is None or owner == name) and held in (None, name):
            if attempt and stats is not None:
                stats["collisions"] = stats.get("collisions", 0) + 1
            return code
        attempt += 1
    # 訊息不放姓名：這一行會出現在畫面上，同仁常截圖求助
    raise RuntimeError("無法配發不重複的代碼，請聯絡維護人員。")


def load_sources():
    """回傳 ({output 相對路徑（小寫）: input 相對路徑}, 是否讀取成功)。"""
    if not SOURCES_FILE.exists():
        return {}, True
    try:
        with SOURCES_FILE.open(encoding="utf-8-sig", newline="") as f:
            return {r["output"].lower(): r["source"] for r in csv.DictReader(f)}, True
    except Exception:  # noqa: BLE001
        return {}, False


def save_sources(sources: dict) -> bool:
    try:
        PRIV_DIR.mkdir(parents=True, exist_ok=True)
        with SOURCES_FILE.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["output", "source"])
            for k in sorted(sources):
                w.writerow([k, sources[k]])
        return True
    except OSError:
        return False


# ====================== 儲存格特徵 ======================

def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str):
        return not v.strip()
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


# 欄名後面的簡短括號註記：「姓名(英)」「電話（公）」（norm_key 已把全形括號轉半形）
_ANNOT = re.compile(r"[(【\[]([^)】\]\d@:]{1,4})[)】\]]$")


def _annot_ok(inner: str) -> bool:
    """括號裡是欄名註記（英、中文、公、校內、en），不是人名（「主持人(黃宇軒)」）。"""
    if inner in ("廠商", "公司", "機關", "單位", "學校", "法人", "團體"):
        return False
    if re.fullmatch(r"[a-z.]{1,4}", inner):
        return True
    if not re.fullmatch(r"[㐀-䶿一-鿿豈-﫿\U00020000-\U0003134F]{1,4}", inner):
        return False
    return len(inner) <= 2 and inner[0] not in _SURNAMES or inner in _ANNOT_WORDS


_ANNOT_WORDS = {"中文", "英文", "校內", "校外", "必填", "選填", "全名", "正楷", "可複選", "公司", "住家", "手機", "市話"}


def _match_key(v) -> str:
    """判定「標題列」用的欄名：先整格比；不在清單時去掉結尾的簡短括號註記再比。
    資料格要不要保留原樣，一律用 _exact_role（整格等於欄名），不能用這個寬鬆比對。"""
    k = norm_key(v)
    if k in NAME_KEYS or k in DROP_KEYS:
        return k
    m = _ANNOT.search(k)
    if not m or not _annot_ok(m.group(1)):
        return k
    k2 = k[:m.start()]
    return k2 if k2 in NAME_KEYS or k2 in DROP_KEYS else k


def _role(v):
    """這一格若是認得的欄名（標題列判定用，可帶簡短括號註記），回傳 'name' 或 'drop'；否則 None。"""
    if not isinstance(v, str):
        return None
    k = _match_key(v)
    if k in NAME_KEYS:
        return "name"
    if k in DROP_KEYS:
        return "drop"
    return None


def _exact_role(v):
    """資料格整格等於欄名（重複表頭、欄位說明）才算；「講師（王大明）」不算，要照樣換碼。"""
    if not isinstance(v, str):
        return None
    k = norm_key(v)
    return "name" if k in NAME_KEYS else "drop" if k in DROP_KEYS else None


_DIGITS = re.compile(r"\d")
# 個資格式：Email、台灣電話／手機（可帶分機）、身分證字號、帳號這類長數字
_PII_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+"
                     r"|^(?:\+?886[-\s]?|0)\d{1,2}[-\s)]?\d{3,4}[-\s]?\d{3,4}(?:\s*(?:#|分機|ext\.?)\s*\d+)?$"
                     r"|^[A-Z][12]\d{8}$|^\d[\d-]{9,19}$", re.I)


def _pii_like(v) -> bool:
    """看起來是聯絡方式或證號（不是一般金額、年度）。"""
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        if isinstance(v, float) and not v.is_integer():
            return False
        s = str(int(v))
        return (len(s) == 9 and s[0] == "9") or len(s) >= 10     # 手機存成數字（少了開頭的 0）、帳號
    return isinstance(v, str) and bool(_PII_RE.search(v.strip()))


def _data_like(v) -> bool:
    """看起來是「資料值」而不是欄名：數字、日期、Email、以數字為主的字串。"""
    if isinstance(v, (bool, int, float, dt.date, dt.datetime, dt.time, pd.Timestamp)):
        return True
    s = str(v).strip()
    if "@" in s:
        return True
    d = len(_DIGITS.findall(s))
    return d >= 3 and d * 2 >= len(re.sub(r"\s", "", s))


# 欄名常見的用語與結尾字。資料列（人名、系所名、職稱）很少整列都長這樣，
# 用來分辨「第二段表格的標題列」與「剛好有一格等於欄名的資料列」。
_HEADERISH_WORDS = ("名稱", "編號", "代號", "代碼", "序號", "項目", "內容", "類別",
                    "欄位", "日期", "時間", "年度", "學年", "學期", "金額", "經費",
                    "單位", "系所", "職稱", "職級", "備註", "說明", "狀態", "性別",
                    "身分", "身份", "資料", "來源", "合計", "小計", "總計")
# 不含「容」：子容、若容是常見名字結尾（「內容」由上面的用語涵蓋）
_HEADERISH_TAIL = tuple("別數率號期間度費額位所稱級註態目碼序日")


def _headerish(v) -> bool:
    if not isinstance(v, str):
        return False
    k = norm_key(v)
    if not k or len(k) > 20:
        return False
    if _role(v):
        return True
    return bool(_HEADERISH_WORDS_RE.search(k)) or k.endswith(_HEADERISH_TAIL)


def _is_str_mask(col: pd.Series) -> np.ndarray:
    return col.map(lambda v: isinstance(v, str)).to_numpy(dtype=bool)


def _hit_matrix(df: pd.DataFrame) -> np.ndarray:
    """每一格是不是認得的欄名（逐欄向量化）。"""
    keys = NAME_KEYS | DROP_KEYS
    hit = np.zeros(df.shape, dtype=bool)
    for ci in range(df.shape[1]):
        col = df.iloc[:, ci]
        is_str = _is_str_mask(col)
        if not is_str.any():
            continue
        s = col[is_str].astype(str).map(norm_key)
        found = s.isin(keys).to_numpy(dtype=bool, copy=True)     # pandas 3 會回傳唯讀陣列
        annot = s.map(lambda x: _ANNOT.search(x) is not None).to_numpy(dtype=bool) & ~found
        if annot.any():                 # 「姓名(英)」這種帶註記的欄名
            found[annot] = [_match_key(x) in keys for x in s[annot]]
        hit[np.flatnonzero(is_str), ci] = found
    return hit


def _nonblank_matrix(df: pd.DataFrame) -> np.ndarray:
    # pandas 3 的 copy-on-write 會讓單欄表格回傳唯讀陣列，一定要自己複製一份
    nb = df.notna().to_numpy(dtype=bool, copy=True)
    for ci in range(df.shape[1]):
        ws = df.iloc[:, ci].map(lambda v: isinstance(v, str) and not v.strip())
        nb[ws.to_numpy(dtype=bool), ci] = False
    return nb


# ====================== 人名樣式（畫面遮蔽、已知姓名過濾） ======================

_SURNAMES = set("陳林黃張李王吳劉蔡楊許鄭謝洪郭邱曾廖賴徐周葉蘇莊呂江何蕭羅高潘簡朱鍾游彭詹胡施沈余"
                "盧梁趙顏柯翁魏孫戴范方宋鄧杜傅侯曹薛丁卓阮馬董溫唐藍石蔣古紀姚連馮歐程湯田康姜白汪鄒"
                "尤巫鐘黎涂龔嚴韓袁金童陸夏柳凃邵錢伍倪于譚駱熊任甘秦顧毛章史萬俞雷饒闕凌崔尹孔辛武"
                "辜陶段龍韋葛池孟褚殷麥賀賈莫"
                # 較少見、但幾乎不會是一般詞彙開頭的姓氏（不收「管、全、應、甄、關」這類常見詞首字）
                "郝靳邢聶耿閻鮑霍翟戚滕裴覃鄔竇繆屠臧冀粘鄞邰刁冉樊祝岳喬"
                "冷仉伏甯蔚闞酆郜鄺厲苟蒯戈璩逯桑敖姬符")
_COMPOUND_SURNAMES = {"歐陽", "司馬", "諸葛", "上官", "司徒", "皇甫", "東方", "長孫", "夏侯", "公孫",
                      "慕容", "令狐", "端木", "尉遲", "張簡", "范姜", "呼延", "宇文", "南宮", "西門", "百里",
                      "東郭", "聞人", "軒轅", "鍾離", "鐘離", "司空", "太史", "左丘", "拓跋", "澹臺", "張廖"}
_HONORIFICS = ("老師", "教授", "先生", "女士", "小姐", "主任", "博士", "同學", "校長", "院長",
               "所長", "組長", "團隊", "研究室", "實驗室")
_ORG_TAILS = tuple("院系所組處室科部會表單冊館校局")
# 費用、比率、編碼類的結尾（「鐘點費」「高鐵票價」不是人名；人名幾乎不會用這些字結尾）
_ITEM_TAILS = tuple("費額率號碼註態價")
_TOTAL_WORDS = {"合計", "小計", "總計", "共計", "平均", "總數", "總和"}
# 姓名欄裡常見、但不是人名的填寫值
_PLACEHOLDERS = {"待聘", "從缺", "不詳", "未定", "待定", "暫缺", "空缺", "同上", "未填", "保密",
                 "不公開", "none", "null", "n/a", "na", "無", "略", "缺", "待補", "tbd", "vacant",
                 "待確認", "尚未聘任", "未聘", "另聘", "外聘", "暫無", "待補人選", "另行通知", "未定案", "#n/a"}


_BRACKET_NOTE = re.compile(r"[（(【\[].*?[）)】\]]")


def _strip_person_decor(s: str) -> str:
    """去掉括號註記與稱謂：「王大明（系所主管）」「林志明老師」→ 王大明、林志明。"""
    t = _BRACKET_NOTE.sub("", s).strip() if ("(" in s or "（" in s or "[" in s or "【" in s) else s.strip()
    while t.endswith(_HONORIFICS):          # 「陳美玲教授團隊」要連續去掉兩層
        for h in _HONORIFICS:
            if t.endswith(h) and len(t) > len(h):
                t = t[:-len(h)].strip()
                break
        else:
            break
    return t


_GROUP_TAILS = ("研究室", "實驗室", "團隊")


def _cjk_person(s: str) -> bool:
    t = re.sub(r"[\s・‧·]", "", _strip_person_decor(s))
    # 「王小明研究室」「陳美玲教授團隊」是人名；「高齡研究室」「金融科技實驗室」是主題
    if s.rstrip().endswith(_GROUP_TAILS) and not (len(t) == 3 or (len(t) == 4 and t[:2] in _COMPOUND_SURNAMES)):
        return False
    if not re.fullmatch(r"[㐀-䶿一-鿿豈-﫿𠀀-𱍏]{2,4}", t):
        return False
    if t.lower() in NAME_KEYS or t.lower() in DROP_KEYS or t.endswith(_ORG_TAILS) or t.endswith(_ITEM_TAILS):
        return False
    if t in _TOTAL_WORDS or _HEADERISH_WORDS_RE.search(t) or _PERSON_WORDS_RE.search(t):
        return False
    return t[:2] in _COMPOUND_SURNAMES or t[0] in _SURNAMES


# 英文姓名的單字：Wang、O'Neil、McDonald；可夾小寫的姓氏前綴（Ludwig van Beethoven）
_LATIN_TOK = r"(?:[A-Z]['’])?[A-Z][a-z]+(?:[A-Z][a-z]+)?"
_LATIN_NAME = (rf"{_LATIN_TOK}(?:[ \-](?:[A-Z]\. )?(?:(?:van|von|de|der|den|da|di|du|del|la|le) ){{0,2}}"
               rf"{_LATIN_TOK}){{1,3}}")


def _latin_person(s: str) -> bool:
    t = _strip_person_decor(s)
    if not re.fullmatch(_LATIN_NAME, t):
        return False
    tokens = set(re.findall(r"[a-z]+", t.lower()))
    return not tokens & (_LATIN_PERSON | _LATIN_PII | _LATIN_BENIGN | _LATIN_MONEY_STATUS | _LATIN_TITLE_WORDS)


# 常見的英文工作表名、檔名用字（「Raw Data」「Summary Table」不是人名）
_LATIN_TITLE_WORDS = {"data", "raw", "summary", "sheet", "list", "report", "table", "overview", "sample",
                      "template", "form", "record", "records", "detail", "details", "info", "information",
                      "note", "notes", "results", "stats", "statistics", "budget", "schedule", "index",
                      "contents", "appendix", "draft", "final", "version", "update", "copy", "main", "test",
                      "chart", "pivot", "output", "input", "backup", "annual", "monthly", "research",
                      "grant", "grants", "faculty", "office", "center", "college", "university", "english",
                      "chinese", "international", "new", "old", "all",
                      "machine", "learning", "human", "resources", "resource", "visiting", "scholars", "scholar",
                      "north", "south", "east", "west", "campus", "board", "meeting", "meetings", "minutes",
                      "agenda", "committee", "conference", "workshop", "seminar", "lecture", "class", "courses",
                      "alumni", "graduate", "undergraduate", "admission", "admissions", "finance", "accounting",
                      "policy", "strategy", "planning", "projects", "team", "teams", "group", "groups", "services",
                      "service", "lab", "laboratory", "library", "national", "taiwan", "public", "private",
                      "social", "science", "sciences", "art", "arts", "history", "language", "economics", "law",
                      "business", "management", "marketing", "education", "health", "medical", "digital",
                      "technology", "innovation", "global", "local", "regional", "community", "industry",
                      "partnership", "development", "evaluation", "progress", "quarterly", "spring", "summer",
                      "fall", "autumn", "winter", "first", "second", "third", "part", "phase", "stage", "round",
                      "roster", "calendar", "timeline", "exchange", "program", "programs", "fund", "funding"}


_person_cache = {}


def looks_like_person(s) -> bool:
    if not isinstance(s, str):
        return False
    r = _person_cache.get(s)
    if r is None:
        if len(_person_cache) > 200000:
            _person_cache.clear()
        r = _person_cache[s] = _cjk_person(s) or _latin_person(s)
    return r


def mask_text(s: str, known=None) -> str:
    """畫面用：已知姓名換成代碼，其餘看起來像人名的片段遮成「＊＊」。"""
    if known is not None:
        s = known.replace_tokens(s)[0]
    # 英文姓名是好幾個單字，要先整段判斷（切開後單字本身不像人名）
    s = re.sub(_LATIN_NAME, lambda mo: "＊＊" if _latin_person(mo.group(0)) else mo.group(0), s)
    # 「原主持人丁柏宏」「郝承恩老師」「承辦老師 符振邦先生」這種黏在一起的寫法
    s = _mask_embedded(s)
    parts = re.split(r"([_\-\s（）()【】\[\]、，,.\\/]+)", s)
    for i, p in enumerate(parts):
        if p and looks_like_person(p):
            parts[i] = "＊＊"
    return mask_title("".join(parts))


_HONOR_WORDS = ("老師", "教授", "先生", "女士", "小姐")


def _mask_embedded(s: str) -> str:
    """畫面遮蔽用：人員用語後面、或稱謂前面的 2～4 字人名遮成「＊＊」（只影響畫面，不改 output）。"""
    if not _CJK.search(s):
        return s
    roles = tuple(sorted(set(_PERSON_WORDS) | set(_MENTION_EXTRA_ROLES)
                         | {k for k in NAME_KEYS if re.fullmatch(f"[{_CJK_CLASS}]{{2,8}}", k)}, key=len, reverse=True))
    out, i = [], 0
    while i < len(s):
        got = 0
        if _CJK.match(s[i]):
            before = s[:i].rstrip(" 　:：")
            role_before = before.endswith(roles)
            for L in (4, 3, 2):
                t = s[i:i + L]
                if len(t) < L or not all(_CJK.match(c) for c in t):
                    continue
                after = s[i + L:]
                hon_after = after.lstrip(" 　").startswith(_HONOR_WORDS)
                boundary = not after or not _CJK.match(after) or after.startswith(_MENTION_AFTER)
                if (hon_after or (role_before and boundary)) and looks_like_person(t):
                    got = L
                    break
        if got:
            out.append("＊＊")
            i += got
        else:
            out.append(s[i])
            i += 1
    return "".join(out)


_parts_cache = {}


def _name_parts(val: str):
    r = _parts_cache.get(val)
    if r is None:
        if len(_parts_cache) > 100000:
            _parts_cache.clear()
        r = _parts_cache[val] = _name_parts_calc(val)
    return r


def _name_parts_calc(val: str):
    """一格寫了好幾個人（「王小明、李大華」「汪傑 Peter Garcia」「程靜怡 劉張霞」「鄭宇軒\n林森」）或
    英文姓在前（「Garcia, Peter」）時，拆出個別的姓名；拆不出來回傳空清單。"""
    s = val.strip().translate(_APOSTROPHES)
    m = re.fullmatch(r"([A-Z][A-Za-z'\-]+)\s*,\s*([A-Z][A-Za-z'\-]+)", s)
    if m:
        return [f"{m.group(2)} {m.group(1)}"]
    pieces = [p.strip() for p in re.split(r"\s*[、，,/／;；&＆\n]\s*", s) if p.strip()]
    if len(pieces) >= 2:
        out = []
        for p in pieces:
            core = _person_core(p)
            if core and (_CJK.search(core) or len(core.split()) >= 2):
                out.append(core)
        return out
    # 只用空白分隔：中英並列「汪傑 Peter Garcia」，或每一段都像人名的中文「程靜怡 劉張霞」「鄭宇軒 林森」
    mixed = re.split(f"(?<=[{_CJK_CLASS}])\\s+(?=[A-Za-z])|(?<=[A-Za-z])\\s+(?=[{_CJK_CLASS}])", s)
    if len(mixed) >= 2:
        return [c for c in (_person_core(p) for p in mixed) if c]
    cjk = s.split()
    if len(cjk) >= 2 and all(re.fullmatch(f"[{_CJK_CLASS}]{{2,4}}", p) and _cjk_person(p) for p in cjk):
        return cjk
    return []


_isperson_cache = {}


def _is_person_value(name: str) -> bool:
    """姓名欄裡被換成代碼的值，要「像人名」才拿去掃描其他地方。"""
    r = _isperson_cache.get(name)
    if r is None:
        if len(_isperson_cache) > 200000:
            _isperson_cache.clear()
        r = _isperson_cache[name] = _is_person_value_calc(name)
    return r


def _is_person_value_calc(name: str) -> bool:
    k = name.strip()
    if not k or len(k) > 40 or _DIGITS.search(k) or _CODE_RE.fullmatch(k):
        return False
    if norm_key(k) in NAME_KEYS or norm_key(k) in DROP_KEYS or k.lower() in _PLACEHOLDERS:
        return False
    if _CJK.search(k):
        core = re.sub(r"[\s・‧·]", "", k)
        if not re.fullmatch(f"[{_CJK_CLASS}]{{2,8}}", core) or core in _PLACEHOLDERS:
            return False
        if core in _TOTAL_WORDS or core.endswith(_ORG_TAILS) or core.endswith(("費", "額")):
            return False
        if _HEADERISH_WORDS_RE.search(core) or _PERSON_WORDS_RE.search(core):
            return False
        # 「推薦人」「指導者」這種角色詞（不是姓氏開頭、以人員字結尾）不是人名
        if core.endswith(_PERSON_TAILS) and not (core[0] in _SURNAMES or core[:2] in _COMPOUND_SURNAMES):
            return False
        return True
    if not re.fullmatch(r"[A-Za-z][A-Za-z.'\-]*(?:\s+[A-Za-z][A-Za-z.'\-]*){0,4}", k):
        return False
    tokens = set(re.findall(r"[a-z]+", k.lower()))
    return not tokens <= (_LATIN_PERSON | _LATIN_PII | _LATIN_BENIGN | _LATIN_MONEY_STATUS)


# ====================== 表格版面判定 ======================

class Layout:
    """一張工作表的判定結果。

    blocks  ：每段表格的標題區 (首列, 末列, {欄: 'name'/'drop'}, {欄: 原始欄名})
    segments：每段的資料列範圍 (起, 迄(不含), 所屬 block 索引)
    """

    def __init__(self):
        self.blocks = []
        self.segments = []
        self.header_rows = set()
        self.doubtful_rows = []        # 看起來像標題列、但沒有採用的列
        self.hit = None
        self.fallback = False          # 主標題列是退回「命中最多的列」選出來的（該列不像標題）
        self.extra_name = []           # [(列, [姓名欄])]：沒採用、但像標題且含姓名欄名的列（以下照樣換碼）


_DATE_LABEL = re.compile(r"\d{2,4}[年/.\-]\d{1,2}(?:月(?:\d{1,2}日)?|[/.\-]\d{1,2}日?)?"
                         r"|\d{1,2}[/.\-]\d{1,2}|\d{1,2}月(?:\d{1,2}日)?")


def _date_label(v) -> bool:
    """簽到表、參與矩陣常把日期或年度當欄名：Excel 日期格、2023、「4/2」「113年3月」。"""
    if isinstance(v, bool):
        return False
    if isinstance(v, (dt.date, dt.datetime, pd.Timestamp)):
        return True
    if isinstance(v, (int, float)):
        return float(v).is_integer() and 1900 <= v <= 2100
    return isinstance(v, str) and bool(_DATE_LABEL.fullmatch(v.strip()))


def _label_like_row(df, nb, i) -> bool:
    """整列沒有數字、日期、Email 這類資料值，也沒有看起來像人名的格子（表頭不會寫人名）。
    日期、年度欄名（「序號｜姓名｜單位｜4/2｜4/9」）要同一列另有兩格以上欄名用語才算。"""
    dates = heads = 0
    for ci in np.flatnonzero(nb[i]):
        v = df.iat[i, ci]
        if _date_label(v) or (isinstance(v, str) and len(v.strip()) <= 30 and "@" not in v
                              and _data_like(v) and _CJK.search(v) and not _pii_like(v)):
            dates += 1                      # 「113年」「110-112年累計」：帶數字的欄名
            continue
        if not isinstance(v, str) or len(v.strip()) > 30 or _data_like(v) or looks_like_person(v):
            return False
        heads += _headerish(v)
    return not dates or heads >= 2


def _strong_header(df, nb, hit, i) -> bool:
    """整列都是欄名用語、至少兩格是認得的欄名——不需要空白列分隔也能當新表頭。"""
    idx = np.flatnonzero(nb[i])
    return (hit[i].sum() >= 2 and len(idx) >= 2 and _label_like_row(df, nb, i)
            and all(_headerish(df.iat[i, ci]) for ci in idx))


def _separated(nb, hit, i, header_rows, loose=False, kv_rows=frozenset(), is_info=None) -> bool:
    """上一列是空白列、只有一格的小標題（例如「114年度」）、「承辦人｜王小明」資訊列；
    loose=True 時也接受兩格以內、沒有欄名的說明列（例如「製表單位｜製表日期」）。"""
    if i == 0:
        return True
    if (i - 1) in header_rows:
        return False
    cnt = nb[i - 1].sum()
    return (cnt <= 1 or (i - 1) in kv_rows or (loose and cnt <= 2 and not hit[i - 1].any())
            or (is_info is not None and is_info(i - 1)))


def _next_data_row(nb, i):
    n = nb.shape[0]
    j = i + 1
    while j < n and j <= i + 3 and not nb[j].any():
        j += 1
    return j if j < n and nb[j].any() else None


def _below_plausible(df, nb, hit, i) -> bool:
    """候選標題列底下那一列，內容要像「資料」。

    例：「經費項目」欄某一格剛好寫「電話」，它底下是「影印費」而不是電話號碼；
    「身分別」欄寫「共同主持人」，它底下又是「主持人」——都不是真的標題列。
    """
    rows, j = [], i
    while len(rows) < 6:                # 每欄看底下前三個有填的值（姓名欄偶爾填員工編號，一格不準）
        j = _next_data_row(nb, j)
        if j is None:
            break
        rows.append(j)
    if not rows:
        return False
    pos = neg = 0
    filled = False
    for ci in np.flatnonzero(hit[i]):
        is_name = _role(df.iat[i, ci]) == "name"
        seen = 0
        for r in rows:
            if not nb[r, ci]:
                continue
            if hit[r, ci]:
                if r == rows[0]:
                    return False        # 底下又是欄名 → 這一欄其實是「欄名清單」
                break
            filled = True
            w = df.iat[r, ci]
            if _data_like(w):
                pos, neg = (pos, neg + 1) if is_name else (pos + 1, neg)
            elif is_name:
                # 底下是欄名用語或長標題（簽核列底下接表名、真正的表頭）不算姓名。
                # 不能用 _headerish：它看結尾字（日、期），「郭旭日」「陳子期」會被當成欄名
                if isinstance(w, str) and (_label_word(w) or len(re.sub(r"\s", "", w)) > 20):
                    neg += 1
                else:
                    pos += 1
            seen += 1
            if seen == 3:
                break
    return (pos > 0 and pos >= neg) or not filled


def _accept_new_header(df, nb, hit, i, roles_now, known_names) -> bool:
    """第二段以後的標題列（或前幾列找不到時的主標題列）：全部條件成立才採用。"""
    if nb[i].sum() < 2 or not _label_like_row(df, nb, i):
        return False
    idx = np.flatnonzero(nb[i])
    labels = [ci for ci in idx if not _date_label(df.iat[i, ci])]
    if sum(_headerish(df.iat[i, ci]) for ci in labels) < 0.6 * len(labels):
        return False
    for ci in idx:
        v = df.iat[i, ci]
        if isinstance(v, str) and (v.strip() in known_names or looks_like_person(v)):
            return False
    # 前一段的姓名欄，新標題列必須在同一個位置「寫著欄名」；
    # 否則多半是一列資料，採用後那一欄的姓名就不會被處理。
    # 整列都是欄名用語的強標題列（例如第二段換成「系所｜姓名｜職稱」）可以寫別的欄名用語。
    strong = _strong_header(df, nb, hit, i)
    j = _next_data_row(nb, i)
    # 兩層表頭的下層（上面是「聯絡方式｜人員」分類列）：認得三個以上欄名、自己的姓名欄底下是人名，
    # 而且上一段姓名欄的位置底下不是人名——那一欄已經換了意思
    moved = hit[i].sum() >= 3 and _name_cols_look_like_names(df, nb, hit, i)
    for ci, role in roles_now.items():
        if role != "name":
            continue
        below_person = j is not None and nb[j, ci] and looks_like_person(df.iat[j, ci])
        if nb[i, ci]:
            if not (hit[i, ci] or (strong and _headerish(df.iat[i, ci])) or (moved and not below_person)):
                return False
        elif j is not None and nb[j, ci] and not (moved and not below_person):
            return False
    return _below_plausible(df, nb, hit, i)


def _kv_row(df, nb, hit, i) -> bool:
    """「承辦人｜梁雅婷」這種資訊列：姓名類欄名右邊緊接著人名，而且欄名底下不是人名。
    這種列不是表頭；當成表頭會把底下整欄換成代碼，也會擋住真正的表頭（右邊的人名另由鍵值對處理）。"""
    m = df.shape[1]
    j = _next_data_row(nb, i)
    for ci in np.flatnonzero(hit[i]):
        if ci + 1 >= m or hit[i, ci + 1] or _role(df.iat[i, ci]) != "name":
            continue
        if looks_like_person(df.iat[i, ci + 1]) and not (j is not None and looks_like_person(df.iat[j, ci])):
            return True
    return False


def _info_row(df, nb, hit, i) -> bool:
    """表頭上方的資訊列（「承辦人｜管秀英｜填表人｜甄志強」「製表人｜倪佩君｜分機｜62817」）：
    6 格以內，有人員欄名，右邊（隔最多 2 格空白）是 2～4 字的名字或英文人名。"""
    idx = np.flatnonzero(nb[i])
    if not 2 <= len(idx) <= 6:
        return False
    for ci in idx:
        v = df.iat[i, ci]
        if not isinstance(v, str):
            continue
        k = norm_key(v)
        if not (hit[i, ci] and _role(v) == "name") and not _PERSON_WORDS_RE.search(k) \
                and not (len(k) <= 5 and k.endswith(("人", "者"))):
            continue
        vc = _right_value_col(nb, i, ci)
        if vc is None:
            continue
        w = df.iat[i, vc]
        if isinstance(w, str) and not hit[i, vc] and not _label_word(w) and (
                looks_like_person(w) or (len(idx) <= 4 and re.fullmatch(f"[{_CJK_CLASS}]{{2,4}}", w.strip()))):
            return True
    return False


def _right_value_col(nb, i, ci, reach=3):
    """同一列右邊第一個有填的格子（跳過合併儲存格造成的空白，最多往右 reach 格）。"""
    for vc in range(ci + 1, min(ci + 1 + reach, nb.shape[1])):
        if nb[i, vc]:
            return vc
    return None


def _keyset(df, nb, i):
    return frozenset(norm_key(df.iat[i, ci]) for ci in np.flatnonzero(nb[i]))


def _name_cols_look_like_names(df, nb, hit, i) -> bool:
    """這一列的姓名欄底下，第一筆資料看起來是人名（不是欄名、不是數字）。"""
    cols = [ci for ci in np.flatnonzero(hit[i]) if _role(df.iat[i, ci]) == "name"]
    j = _next_data_row(nb, i)
    if not cols or j is None or nb[i].sum() < 2:
        return False
    vals = [(ci, df.iat[j, ci]) for ci in cols if nb[j, ci]]
    return bool(vals) and all(isinstance(v, str) and not hit[j, ci] and not _data_like(v) for ci, v in vals) \
        and any(looks_like_person(v) or (re.fullmatch(f"[{_CJK_CLASS}]{{2,4}}", v.strip()) and not _label_word(v))
                for _ci, v in vals)


def detect_layout(df: pd.DataFrame, known_names=frozenset(), nb=None) -> Layout:
    lay = Layout()
    n, m = df.shape
    if n == 0 or m == 0:
        return lay
    hit = _hit_matrix(df)
    hits = hit.sum(axis=1)
    lay.hit = hit
    if not hits.any():
        return lay
    if nb is None:
        nb = _nonblank_matrix(df)
    info_cache = {}

    def is_info(r):
        if r not in info_cache:
            info_cache[r] = _info_row(df, nb, hit, r)
        return info_cache[r]

    # ---- 主標題列 ----
    kv_rows = frozenset(int(i) for i in np.flatnonzero(hits > 0) if _kv_row(df, nb, hit, int(i)))
    top = [i for i in range(min(MAX_SCAN, n)) if hits[i] > 0 and i not in kv_rows]
    clean = [i for i in top if _label_like_row(df, nb, i)]
    primary = -1
    for i in clean:                     # 由上往下第一個「明確是標題」的列
        if _separated(nb, hit, i, set(), loose=True, kv_rows=kv_rows, is_info=is_info) \
                and _accept_new_header(df, nb, hit, i, {}, known_names):
            primary = i
            break
    # 含姓名欄名、而且底下看起來是人名的列（「教師姓名｜服務機關｜現職」這種欄名用語不多的表頭）
    name_first = next((i for i in clean if _name_cols_look_like_names(df, nb, hit, i)), -1)
    if primary < 0 and clean:
        pool = [i for i in clean if nb[i].sum() >= 2] or clean     # 單格小標題（「教師」）最後才考慮
        primary = max(pool, key=lambda i: (hits[i], -i))    # v1.2：命中最多，同分取最上面
        if nb[primary].sum() < 2 and any(hits[i] >= 2 for i in top if i != primary):
            lay.fallback = True             # 只剩單格小標題可用，真正的表頭多半沒認出來
    if name_first >= 0 and name_first != primary and (primary < 0 or name_first < primary
                                                       or not _name_cols_look_like_names(df, nb, hit, primary)):
        primary = name_first
    if primary < 0:
        # 前 8 列找不到：整張表由上往下找。三個以上認得的欄名（兩層表頭的第二層、上面緊接分類列）不必有分隔列
        for i in np.flatnonzero(hits > 0):
            i = int(i)
            if (_separated(nb, hit, i, set(), loose=True, kv_rows=kv_rows, is_info=is_info) or hits[i] >= 3)                     and _accept_new_header(df, nb, hit, i, {}, known_names):
                primary = i
                break
    if primary < 0 and top:
        primary = max(top, key=lambda i: (hits[i], -i))
        lay.fallback = True             # 標題列不像標題（含資料值），判定不可靠，要提醒
    if primary < 0:
        lay.doubtful_rows = [int(i) for i in np.flatnonzero(hits >= 2)
                             if _label_like_row(df, nb, int(i))][:50]
        return lay

    def blank_in(rows_from, rows_to, ci):
        return not nb[rows_from:rows_to + 1, ci].any()

    def grow_up(first, last):
        # 補齊空白欄位的上層標題（例如「姓名」上下合併、第二層只寫電話／地址）
        while first - 1 >= 0 and hits[first - 1] > 0 and _label_like_row(df, nb, first - 1):
            if not any(hit[first - 1, ci] and blank_in(first, last, ci) for ci in range(m)):
                break
            first -= 1
        return first

    def grow_down(first, last):
        # 第二層標題：整列都是欄名用語（沒有人名、沒有資料值），並替本區還沒有欄名的欄位補上欄名
        while last + 1 < n and hits[last + 1] > 0:
            r = last + 1
            if not _label_like_row(df, nb, r):
                break
            idx = np.flatnonzero(nb[r])
            if not all(_headerish(df.iat[r, ci]) or suspect_label(df.iat[r, ci]) or _date_label(df.iat[r, ci])
                       or re.fullmatch(r"[A-Za-z][A-Za-z .\-/#]{1,20}", str(df.iat[r, ci]).strip())
                       for ci in idx):
                break
            fills = [ci for ci in idx if hit[r, ci] and not any(hit[first:last + 1, ci])]
            if not fills or (len(idx) == 1 and nb[first:last + 1, idx[0]].any()):
                break
            last = r
        return last

    def add_block(first, last):
        rows = list(range(first, last + 1))
        roles, labels = {}, {}
        for ci in range(m):
            for r in reversed(rows):            # 每一欄取最下面那個有字的格子
                if nb[r, ci]:
                    v = df.iat[r, ci]
                    labels[ci] = re.sub(r"\s+", " ", str(v)).strip()
                    role = _role(v)
                    if role:
                        roles[ci] = role
                    break
        lay.blocks.append((first, last, roles, labels))
        lay.header_rows.update(rows)
        return roles

    first = grow_up(primary, primary)
    last = grow_down(first, primary)
    r = last + 1
    if r < n and hits[r] == 0 and nb[r].sum() >= 1 and _label_like_row(df, nb, r) and all(
            suspect_label(df.iat[r, ci]) or _headerish(df.iat[r, ci]) or _date_label(df.iat[r, ci])
            or re.fullmatch(r"[A-Za-z][A-Za-z .\-/#]{1,20}", str(df.iat[r, ci]).strip())
            for ci in np.flatnonzero(nb[r])) and any(suspect_label(df.iat[r, ci]) for ci in np.flatnonzero(nb[r])):
        last = r
    roles_now = add_block(first, last)
    for r in range(first):
        if hits[r] >= 2 and _label_like_row(df, nb, r) and r not in kv_rows and len(lay.doubtful_rows) < 50:
            lay.doubtful_rows.append(int(r))
    primary_width = int(nb[primary].sum())
    keysets = {_keyset(df, nb, primary)}

    # ---- 第二段以後的標題列 ----
    def plain_header(r, roles_prev):
        """上面是空白列、整列都是欄名用語、蓋到上一段的姓名或清空欄，
        而且底下那一列在上一段的姓名欄不像人名、清空欄不像個資。"""
        if r == 0 or nb[r - 1].any() or nb[r].sum() < 2 or not roles_prev:
            return False
        idx = np.flatnonzero(nb[r])
        if not any(ci in roles_prev for ci in idx) or not _label_like_row(df, nb, r):
            return False
        if not all(_headerish(df.iat[r, ci]) or _date_label(df.iat[r, ci]) for ci in idx):
            return False
        j = _next_data_row(nb, r)
        if j is None:
            return False
        for ci, role in roles_prev.items():
            if nb[j, ci]:
                w = df.iat[j, ci]
                if (role == "name" and looks_like_person(w)) or (role == "drop" and _pii_like(w)):
                    return False
        return True

    i = last + 1
    while i < n:
        if hits[i] == 0:
            if plain_header(i, roles_now):
                roles_now = add_block(i, i)
            i += 1
            continue
        accepted = nb[i].sum() == primary_width and _keyset(df, nb, i) in keysets   # 每頁重複的表頭
        if not accepted and (_separated(nb, hit, i, lay.header_rows, kv_rows=kv_rows, is_info=is_info) or hits[i] >= 3
                             or _strong_header(df, nb, hit, i)):
            accepted = _accept_new_header(df, nb, hit, i, roles_now, known_names)
        if not accepted:
            label_like = _label_like_row(df, nb, i)
            name_cols = [int(ci) for ci in np.flatnonzero(hit[i]) if _role(df.iat[i, ci]) == "name"]
            candidate = label_like and nb[i].sum() >= 2 and name_cols and i not in kv_rows and (
                _separated(nb, hit, i, lay.header_rows, kv_rows=kv_rows, is_info=is_info)
                or _strong_header(df, nb, hit, i))
            if candidate:
                lay.extra_name.append((int(i), name_cols))
            if (hits[i] >= 2 and label_like or candidate) and len(lay.doubtful_rows) < 50:
                lay.doubtful_rows.append(int(i))
            i += 1
            continue
        last = grow_down(i, i)
        roles_now = add_block(i, last)
        keysets.add(_keyset(df, nb, i))
        i = last + 1

    # ---- 每段的資料列範圍；段落最後的單格小標題列不算資料 ----
    for bi, (_first, last_, _r, _l) in enumerate(lay.blocks):
        nxt = lay.blocks[bi + 1][0] if bi + 1 < len(lay.blocks) else n
        stop = nxt
        if bi + 1 < len(lay.blocks):
            k = nxt - 1
            while k > last_ and not nb[k].any():
                k -= 1
            if k > last_ and nb[k].sum() <= 1 and not hit[k].any():
                stop = k
        lay.segments.append((last_ + 1, stop, bi))
    return lay


def _label_word(v) -> bool:
    """資料格的內容是欄名用語（「序號」「單位」「姓名」）。只看完整用語，不看結尾字（「郭旭日」不是欄名）。"""
    if not isinstance(v, str) or looks_like_person(v):
        return False
    k = norm_key(v)
    return bool(_role(v)) or bool(_HEADERISH_WORDS_RE.search(k))


def _col_has_name(df, start, stop, ci) -> bool:
    for v in df.iloc[start:stop, ci].tolist():
        if isinstance(v, str) and (looks_like_person(v) or (
                re.fullmatch(f"[{_CJK_CLASS}]{{2,4}}", v.strip()) and not _label_word(v)
                and not v.strip().endswith(_ORG_TAILS + _ITEM_TAILS))):
            return True
    return False


def _names_below(df, nb, done, ri, ci) -> bool:
    """欄名底下接著幾列像人名、但沒有處理的格子。"""
    seen = 0
    for r in range(ri + 1, min(df.shape[0], ri + 8)):
        if not nb[r, ci]:
            continue
        if looks_like_person(df.iat[r, ci]) and not done[r, ci]:
            return True
        seen += 1
        if seen >= 3:
            break
    return False


def _legacy_marks(df: pd.DataFrame, hit: np.ndarray, data_rows=frozenset()):
    """保底判定：回傳 [(標題列, 要採用的角色 'all'／'name')]，每一列以下全部處理。

    1. v1.2 的判斷：前 8 列中用 v1.2 清單（完全相符）命中最多的一列。
    2. 同樣做法套用現在的清單（只採姓名欄，避免資料列剛好寫「電話」就整欄清空）。
    3. 前 8 列中第一個含姓名欄名、沒有資料值的列（v1.2 清單不認得的新欄名）。
    """
    out = []
    n = min(MAX_SCAN, df.shape[0])
    if not n:
        return out
    legacy = LEGACY_NAME_COLUMNS | LEGACY_DROP_COLUMNS
    best, best_hits = -1, 0
    for i in range(n):
        h = sum(1 for v in df.iloc[i] if norm_header(v) in legacy)
        if h > best_hits:
            best, best_hits = i, h
    if best >= 0:
        out.append((best, "legacy"))
    nb = _nonblank_matrix(df.iloc[:n])
    name_first = -1
    for i in range(n):
        if any(hit[i, ci] and _role(df.iat[i, ci]) == "name" for ci in range(df.shape[1])) \
                and _label_like_row(df, nb, i) and nb[i].sum() >= 2 and i not in data_rows:
            name_first = i
            break
    if name_first >= 0:
        out.append((name_first, "all"))
    counts = hit[:n].sum(axis=1)
    for i in np.flatnonzero(counts):
        # 「承辦人｜梁雅婷」資訊列、單格小標題、已判定表格裡的資料列，都不是標題列
        if _kv_row(df, nb, hit, int(i)) or nb[i].sum() < 2 or int(i) in data_rows:
            counts[i] = 0
    if counts.max() > 0:
        am = int(np.argmax(counts))
        # 那一列若在上面「姓名欄名列」的姓名欄位置寫著人名，它是資料列，不採用
        if not (name_first >= 0 and am > name_first and any(
                _role(df.iat[name_first, ci]) == "name" and looks_like_person(df.iat[am, ci])
                for ci in range(df.shape[1]))):
            out.append((am, "name"))
    return out


# ====================== 疑似個資欄位 ======================

# 判斷看的是欄名的「結構」而不只是有沒有某個字：
#   ・人員用語 → 這一欄多半是人名（推薦人、單位主管、講者、作者群）
#   ・個資用語 → 聯絡方式、證件、帳號（健保卡號、LINE ID）
#   ・欄名「結尾」是屬性（職稱、日期、人數、金額）→ 描述人的屬性，不是個資
#     （「主持人職稱」不示警，但「單位主管」結尾是人員用語，要示警）
#   ・人員或證件用語＋編號／代碼 → 可以把同一人串起來的識別碼，一律示警
_PERSON_WORDS = ("姓名", "名字", "人名", "筆名", "別名", "主持人", "申請人", "指導教授", "教授",
                 "教師", "老師", "講師", "講者", "業師", "導師", "作者", "撰稿", "委員", "審查人",
                 "評審", "助理", "研究人員", "研究員", "人員", "成員", "學員", "學生", "研究生",
                 "博士生", "碩士生", "負責人", "執行人", "聯絡人", "承辦人", "經辦人", "代理人",
                 "受訪", "填表人", "得獎", "受獎", "獲獎", "得主", "參與者", "參加者", "報名者",
                 "發明人", "簽核人", "核定人", "主管", "主任", "院長", "校長", "所長", "組長",
                 "召集人", "校友", "受款人", "領款人", "申請者", "配偶", "眷屬", "家長", "監護人",
                 "口委", "窗口", "簽名", "全名", "本名", "承辦", "戶名", "助教")
_PERSON_TAILS = ("人", "者", "員", "師", "生", "長")          # 推薦人、保證人、講者、職員、醫師
_PII_WORDS = ("信箱", "email", "e-mail", "mail", "郵件", "電郵", "電話", "手機", "行動", "分機", "聯絡方式", "連絡方式", "門牌",
              "傳真", "地址", "住址", "戶籍", "通訊處", "學號", "身分證", "身份證", "護照", "居留證",
              "生日", "出生", "帳號", "帳戶", "卡號", "車牌", "網頁", "網址", "臉書", "健保", "統一證號")
_IDENT_WORDS = ("編號", "代號", "代碼", "證號", "號碼", "帳號", "卡號")
# 「任一詞出現在字串裡」：逐詞比對很慢（每格都要跑幾十次），改成一次正規表示式
_HEADERISH_WORDS_RE = re.compile("|".join(map(re.escape, _HEADERISH_WORDS)))
_PERSON_WORDS_RE = re.compile("|".join(map(re.escape, _PERSON_WORDS)))
_PII_WORDS_RE = re.compile("|".join(map(re.escape, _PII_WORDS)))
# 識別碼前面接這些字也算個人識別碼（人事代碼、使用者帳號、身心障礙手冊號碼）
_IDENT_CONTEXT = ("人事", "個人", "員工", "職員", "教職員", "使用者", "手冊", "病歷")
# 結尾是這些屬性用語的欄名不示警
_BENIGN_TAILS = ("職稱", "職級", "職務", "職別", "職等", "單位", "系所", "學院", "科系", "部門",
                 "別", "類型", "數", "數量", "率", "比例", "金額", "經費", "費", "薪", "酬", "津貼",
                 "補助", "禮金", "日期", "日", "時間", "年度", "學年", "學期", "狀態", "領域", "專長", "機構",
                 "學校", "現職", "備註", "說明", "等級", "排名", "排序", "順序", "序", "分數", "成績",
                 "年資", "題目", "所屬", "關係", "來源", "結果", "完成", "情形", "方式", "上限", "下限",
                 "意見", "人次", "容量")
_BENIGN_ANYWHERE = ("是否", "有無", "委員會")
_LATIN_PERSON = {"name", "surname", "person", "contact", "author", "authors", "advisor", "supervisor",
                 "pi", "investigator", "employee", "staff", "student", "stu", "member", "applicant",
                 "speaker", "reviewer", "mentor", "instructor", "lecturer", "interviewee", "participant"}
_LATIN_PII = {"email", "mail", "mobile", "phone", "tel", "fax", "address", "addr", "birthday", "birth",
              "dob", "line", "orcid", "scopus", "facebook", "passport", "ext", "cellphone"}
_LATIN_BENIGN = {"count", "number", "time", "year", "status", "type", "title", "total", "allowance",
                 "fee", "amount", "rate", "project", "file", "course", "journal", "unit", "department",
                 "dept", "school", "program", "item", "award", "plan", "company", "order", "result",
                 "institution"}
_LATIN_MONEY_STATUS = {"allowance", "fee", "amount", "rate", "count", "status", "total"}


def suspect_label(label) -> bool:
    """欄名看起來可能含姓名或個資、但不在處理清單內。label 請傳原始欄名（保留空白）。"""
    if not isinstance(label, str):
        return False
    label = unicodedata.normalize("NFKC", label).translate(_SIMPLIFIED)
    k = norm_key(label)
    if (not k or len(k) > (40 if k.isascii() else 20) or k in NAME_KEYS or k in DROP_KEYS
            or _data_like(label)):
        return False
    # 合併欄名：「姓名/職稱」「共同主持人及職稱」「姓名職稱」「指導教授/系所」——其中一段是姓名欄名就示警
    parts = [x for x in re.split(r"[/、,&＆]|及|與|和", k) if x]
    if len(parts) >= 2 and any(x in NAME_KEYS or x in DROP_KEYS for x in parts):
        return True
    # 「姓名職稱」「姓名單位」：姓名欄名後面直接接屬性（「主持人職稱」是主持人的職稱，不算）
    if any(k.startswith(key) and 0 < len(k) - len(key) <= 4 for key in NAME_KEYS if "姓名" in key)             and not any(w in k for w in _BENIGN_ANYWHERE):
        return True
    if k in ("聯絡方式", "連絡方式", "聯繫方式"):
        return True
    k_attr = re.sub(r"\([^)]*\)$", "", k) or k       # 「電話費(元)」看的是「電話費」
    if len(k) <= 4 and ("姓" in k or k in ("性名", "名子")):
        return True
    words = re.findall(r"[a-z]+", label.lower())
    tokens = set(words)
    person = bool(_PERSON_WORDS_RE.search(k)) or k.endswith(_PERSON_TAILS) or bool(tokens & _LATIN_PERSON)
    pii = bool(_PII_WORDS_RE.search(k)) or bool(tokens & _LATIN_PII)
    # 識別碼：欄名以「編號／代碼／ID／No」收尾才算（「帳號申請日期」描述的是日期）
    ident = k.endswith(_IDENT_WORDS) or (bool(words) and words[-1] in ("id", "no"))
    if words[:2] in (["id", "no"], ["id", "number"]) or k in ("id", "idno", "idnumber"):
        return True
    if ident and (person or pii or any(w in k for w in _IDENT_CONTEXT)):
        return True
    benign = k_attr.endswith(_BENIGN_TAILS) or any(w in k for w in _BENIGN_ANYWHERE)
    # 「教師名稱」「主持人名稱」就是姓名（但「委員會名稱」是單位）
    if person and not benign and k.endswith(("名稱", "名")):
        return True
    # 金額、狀態類英文用語一律算屬性（Phone Allowance）；其他英文屬性詞遇到個資用語時不豁免（Date of Birth）
    if tokens & _LATIN_MONEY_STATUS or (tokens & _LATIN_BENIGN and not tokens & _LATIN_PII):
        benign = True
    return (person or pii) and not benign


# ====================== 已知姓名比對（姓名欄以外） ======================

_CJK_CLASS = "㐀-䶿一-鿿豈-﫿\U00020000-\U0003134F"
_CJK = re.compile(f"[{_CJK_CLASS}]")
_LATIN_WORD = re.compile(r"[A-Za-z]+(?:['.\-][A-Za-z]+)*")
_HAS_LATIN = re.compile(r"[A-Za-z]")
_APOSTROPHES = str.maketrans({"’": "'", "‘": "'", "ʼ": "'", "`": "'", "ı": "i", "İ": "I"})

# 文字裡「人員欄名＋冒號或空白＋姓名」的寫法（這個人沒有出現在姓名欄也換得掉）：
#   「承辦人：王小明　分機 1234」「聯絡人 王小明、李大華」
# 只收 3 個字（或複姓開頭 4 個字）的中文姓名：兩個字、四個字的一般詞（黃金、高教深耕、金門大學）誤判太多。
_MENTION_EXTRA_ROLES = ("製表人", "審核人", "核稿人", "單位主管", "主席", "報告人", "代理人",
                        "組長", "系主任", "主任", "所長", "院長", "處長", "秘書", "專員", "副本", "會辦")
# 人員欄名後面沒有冒號、直接接姓名時，姓名後面要接這些字才算（「原主持人阮婷怡退休」）
_MENTION_VERBS = ("已", "於", "負責", "退休", "代理", "主持", "擔任", "協助", "另", "借調", "離職", "請假", "因",
                  "處理", "辦理", "彙整", "聯繫", "確認", "審核", "審查", "報告", "說明", "討論")
_MENTION_AFTER = ("分機", "電話", "手機", "信箱", "老師", "教授", "先生", "女士", "小姐")
# 職稱結尾：「林專員」「陳組長」不是完整姓名
_TITLE_WORDS = ("專員", "組員", "科員", "辦事員", "秘書", "經理", "主任", "組長", "科長", "課長", "處長",
                "老師", "教授", "同學", "先生", "小姐", "女士", "助理")
_mention_cache = {}


def _mention_role_re():
    # 欄名清單重建時 NAME_KEYS 會換成新的 set；同一個物件就沿用上次編好的正規表示式
    if _mention_cache.get("src") is not NAME_KEYS:
        roles = frozenset(k for k in NAME_KEYS if re.fullmatch(f"[{_CJK_CLASS}]{{2,8}}", k)) | set(_MENTION_EXTRA_ROLES)
        alt = "|".join(sorted(map(re.escape, roles), key=len, reverse=True))
        _mention_cache.clear()
        _mention_cache["src"] = NAME_KEYS
        _mention_cache["re"] = re.compile(f"(?:{alt})(?P<sep>[ 　]*[:：][ 　]*|[ 　]*[（(][ 　]*|[ 　]+|)")
    return _mention_cache["re"]


def _mention_name_ok(t: str) -> bool:
    if any(w in t for w in _TITLE_WORDS) or not _cjk_person(t):
        return False
    if len(t) == 3:
        return True
    # 四個字：複姓（歐陽、張簡）或冠夫姓（周黃淑芬）
    return len(t) == 4 and (t[:2] in _COMPOUND_SURNAMES or t[1] in _SURNAMES)         and not t.endswith(_TITLE_NON_NAME_TAILS) and not (set(t) & _NON_NAME_CHARS)


_HONOR_RE = None


def _mention_spans(s: str, unsure=None):
    """回傳 [(起, 迄, 姓名)]：
    a. 人員欄名後面、以頓號逗號斜線或空白分隔的 3～4 字姓名（「承辦人：王小明、李大華」）
    b. 格子開頭或標點、空白之後的「3～4 字姓名＋老師／教授／先生／女士／小姐」（「王小明老師退休」）
    兩個字的候選（「聯絡人 柯霞」「嚴磊老師」）不換碼，放進 unsure 讓畫面提醒（「承辦人：黃金」這種一般詞太多）。"""
    global _HONOR_RE
    if _HONOR_RE is None:
        _HONOR_RE = re.compile(f"(?<![{_CJK_CLASS}])([{_CJK_CLASS}]{{3,4}})(?=[ 　]?(?:老師|教授|先生|女士|小姐))")
    out = []
    for mo in _HONOR_RE.finditer(s):
        if _mention_name_ok(mo.group(1)):
            out.append((mo.start(1), mo.end(1), mo.group(1)))
    for mo in re.finditer(rf"({_LATIN_NAME})[ 　]?(?=老師|教授|先生|女士|小姐)", s):
        if _latin_person(mo.group(1)):
            out.append((mo.start(1), mo.end(1), mo.group(1)))
    if unsure is not None:
        for mo in re.finditer(f"(?<![{_CJK_CLASS}])([{_CJK_CLASS}]{{2}})(?=老師|教授|先生|女士|小姐)", s):
            if _cjk_person(mo.group(1)):
                unsure.append(mo.group(1))
        for mo in re.finditer(f"(?<=[{_CJK_CLASS}])([{_CJK_CLASS}]{{3}})(?=老師|教授|先生|女士|小姐)", s):
            if _mention_name_ok(mo.group(1)):
                unsure.append(mo.group(1))
    for mo in _mention_role_re().finditer(s):
        pos = mo.end()
        glued = not mo.group("sep")               # 人員欄名後面直接接字
        while pos < len(s):
            got = None
            for L in (4, 3):
                t = s[pos:pos + L]
                if len(t) != L or not _CJK.fullmatch(t[-1]) or not all(_CJK.fullmatch(c) for c in t):
                    continue
                rest = s[pos + L:]
                # 冒號、空白、括號後面：只接受兩個字以上的動詞（「主持人 曾任教於本校」的「於」不算）
                verbs = _MENTION_VERBS if glued else tuple(w for w in _MENTION_VERBS if len(w) >= 2)
                after_ok = rest.startswith(_MENTION_AFTER) or rest.startswith(verbs)
                if glued and not (rest == "" or after_ok):
                    continue
                if rest and _CJK.match(rest) and not after_ok:
                    if L == 3 and unsure is not None and mo.group("sep").strip() and _mention_name_ok(t):
                        unsure.append(t)            # 「承辦人：林佩珊○○」後面接的字不確定：不換，提醒
                    continue
                if _mention_name_ok(t):
                    got = t
                    break
            if not got:
                if glued:
                    break
                t = s[pos:pos + 2]
                rest = s[pos + 2:]
                if unsure is not None and len(t) == 2 and all(_CJK.fullmatch(c) for c in t)                         and not (rest and _CJK.match(rest) and not rest.startswith(_MENTION_AFTER))                         and _cjk_person(t) and not any(w in t for w in _TITLE_WORDS):
                    unsure.append(t)
                break
            out.append((pos, pos + len(got), got))
            pos += len(got)
            sep = re.match(r"[ 　]*[、，,/／]?[ 　]*", s[pos:])
            if not sep or not sep.group(0) or pos + len(sep.group(0)) >= len(s):
                break
            pos += len(sep.group(0))
    # 兩種寫法可能抓到同一段，去重並依位置排序
    return sorted(set(out))


_HONOR_HINT = re.compile("老師|教授|先生|女士|小姐")
_MULTI_HINT = re.compile(r"[、，,/／;；&＆]|\s")


def _glue(code: str, following: str, sep: str) -> str:
    """代碼後面緊接英數字時補分隔，避免「T-XXXXXXXXXX2024」變成查不回的長代碼。"""
    return code + sep if following[:1].isascii() and following[:1].isalnum() else code


_core_cache = {}


def _person_core(name: str):
    """姓名欄的值去掉註記、稱謂、「等3人」、英文稱謂前綴、零寬字元後的核心姓名；不像人名回傳 None。"""
    key = str(name)
    if key in _core_cache:
        return _core_cache[key]
    if len(_core_cache) > 200000:
        _core_cache.clear()
    _core_cache[key] = r = _person_core_calc(key)
    return r


_ETC_PEOPLE = re.compile(r"等\s*\d*\s*人$")
_TITLE_PREFIX = re.compile(r"^(?:dr|prof|mr|mrs|ms|miss)\.?\s+", re.I)
_CORE_STRIP = re.compile(r"[\s・‧·•．.]")


def _person_core_calc(name: str):
    t = str(name)
    if not unicodedata.is_normalized("NFKC", t):
        t = unicodedata.normalize("NFKC", t)
    t = _ZERO_WIDTH.sub("", t).translate(_APOSTROPHES).strip() if "\u200b" in t or "\ufeff" in t \
        else t.translate(_APOSTROPHES).strip()
    if "等" in t:
        t = _ETC_PEOPLE.sub("", t).strip()
    if "." in t or " " in t:
        t = _TITLE_PREFIX.sub("", t)
    t = _strip_person_decor(t)
    if _CJK.search(t):
        core = _CORE_STRIP.sub("", t)
        return core if _is_person_value(core) else None
    t = " ".join(t.split())
    return t if t and _is_person_value(t) else None


# 常見的非姓名填寫值用字：姓名欄裡的「待確認」「尚未聘任」「專案經理」不能拿去換別處的文字
_NON_NAME_CHARS = set("待確認未尚聘任經理專案另行通知暫無缺補定填詳及與或等之的是否請洽詢問")


def _substring_safe(core: str) -> bool:
    """中文姓名要夠「像人名」才做句中子字串替換；其餘只做整格比對。"""
    if _cjk_person(core):
        return True
    return len(core) == 3 and not (set(core) & _NON_NAME_CHARS) and not core.endswith(_PERSON_TAILS)


_SHORT_BEFORE = frozenset("由與和跟請給向交經陪找及同讓被派邀")
_SHORT_AFTER = _MENTION_VERBS + _MENTION_AFTER


class KnownNames:
    """把「已經配發過代碼、而且像人名」的姓名做成查詢表，用來掃描姓名欄以外的地方。"""

    def __init__(self, rev: dict):
        self.exact, subs, short = {}, {}, {}
        self.latin_single, self.latin_multi, self.latin_rev = {}, {}, {}
        items = []
        for n_, c in rev.items():
            core = _person_core(n_) if isinstance(n_, str) else None
            if core:
                items.append((core, c))
        for al, f in ALIASES.items():
            core = _person_core(al)
            if core and f in rev:
                items.append((core, rev[f]))
        for name, code in items:
            k = name.strip()
            if _CJK.search(k):
                core = re.sub(r"\s", "", k)
                self.exact.setdefault(core, code)
                if len(core) >= 3:
                    if _substring_safe(core):
                        subs.setdefault(core, code)   # 夾在句子裡的中文姓名：3 個字以上
                else:
                    short.setdefault(core, code)      # 兩個字：句子裡不自動換，只提醒
            else:
                toks = k.split()
                key = " ".join(toks).lower()
                self.exact.setdefault(k, code)
                if len(toks) >= 2:
                    self.latin_multi.setdefault(key, code)
                    if len(toks) == 2:          # 姓在前的整格寫法：「Garcia, Peter」
                        self.latin_rev.setdefault(f"{toks[1]}, {toks[0]}".lower(), code)
                else:
                    self.latin_single.setdefault(key, code)   # 單字英文名只做整格比對
        self.subs, self.short = subs, short
        self.sub_lens = sorted({len(k) for k in subs}, reverse=True)
        self.min_sub = min(self.sub_lens) if self.sub_lens else 10 ** 9
        firsts = sorted({k[0] for k in subs})
        self.first_re = re.compile("[" + "".join(re.escape(c) for c in firsts) + "]") if firsts else None
        self.multi_lens = sorted({len(k.split()) for k in self.latin_multi}, reverse=True)
        self.heads = {k.split()[0] for k in self.latin_multi}
        # 兩字姓名：句子裡一側不是中文字，或前後是「由、與」「負責、代理」這類字才算提到（「兼顧國際化」不算）
        self.short_first = re.compile("[" + "".join(re.escape(c) for c in sorted({k[0] for k in short})) + "]") \
            if short else None

    def __bool__(self):
        return bool(self.exact)

    def _latin(self, s: str, sep: str):
        if not self.latin_multi or not _HAS_LATIN.search(s):
            return s, 0
        norm = s.translate(_APOSTROPHES)
        heads = self.heads
        toks = _LATIN_WORD.findall(norm)
        lows = [w.lower() for w in toks]
        if not any(w in heads for w in lows):            # 查表比 5000 個名字的正規表示式快很多
            return s, 0
        words, pos = [], 0                               # 位置依序用 find 算（比建立 match 物件快）
        for w, low in zip(toks, lows):
            st = norm.find(w, pos)
            pos = st + len(w)
            words.append((st, pos, low))
        spans, i = [], 0
        while i < len(words):
            if words[i][2] not in heads:
                i += 1
                continue
            for L in self.multi_lens:
                if i + L > len(words):
                    continue
                seg = words[i:i + L]
                if any(norm[seg[t][1]:seg[t + 1][0]].strip() for t in range(L - 1)):
                    continue                            # 單字之間只能是空白
                code = self.latin_multi.get(" ".join(w for _a, _b, w in seg))
                if code:
                    spans.append((seg[0][0], seg[-1][1], code))
                    i += L
                    break
            else:
                i += 1
        if not spans:
            return s, 0
        out, last = [], 0
        for a_, b_, code in spans:
            out.append(s[last:a_])
            out.append(_glue(code, s[b_:], sep))
            last = b_
        out.append(s[last:])
        return "".join(out), len(spans)

    def replace(self, s: str, sep: str = " "):
        """回傳 (新字串, 替換處數)。整格相符直接換；句子裡的姓名逐一換。"""
        t = s.strip()
        if _CJK.search(t):
            nows = re.sub(r"\s", "", t)
            if nows in self.exact and len(nows) >= 3:  # 「王 小明」「王小明」視為同一人
                return self.exact[nows], 1
        elif t.translate(_APOSTROPHES) in self.exact and " " in t.strip():
            return self.exact[t.translate(_APOSTROPHES)], 1
        low = " ".join(t.translate(_APOSTROPHES).split()).lower()
        if low in self.latin_multi:
            return self.latin_multi[low], 1
        if self.latin_rev and "," in low:
            rev_key = re.sub(r"\s*,\s*", ", ", low)
            if rev_key in self.latin_rev:
                return self.latin_rev[rev_key], 1
        s, cnt = self._latin(s, sep)
        if self.first_re is None or len(s) < self.min_sub:
            return s, cnt
        out, last, pos = [], 0, 0
        found = 0
        m = self.first_re.search(s, pos)
        while m:
            i, step = m.start(), 1
            for L in self.sub_lens:
                seg = s[i:i + L]
                if len(seg) == L and seg in self.subs:
                    out.append(s[last:i])
                    out.append(_glue(self.subs[seg], s[i + L:], sep))
                    last, step, found = i + L, L, found + 1
                    break
            m = self.first_re.search(s, i + step)
        if not found:
            return s, cnt
        out.append(s[last:])
        return "".join(out), cnt + found

    def replace_tokens(self, s: str):
        """檔名、工作表名用：句子比對之外，以分隔符切開後「整段」等於已知姓名（含兩個字）也換。"""
        new, cnt = self.replace(s, sep="_")
        parts = re.split(r"([_\-\s（）()【】\[\]、，,.]+)", new)
        for idx, p in enumerate(parts):
            if p in self.exact:
                parts[idx] = self.exact[p]
                cnt += 1
            elif p in self.short:
                parts[idx] = self.short[p]
                cnt += 1
            elif p.lower() in self.latin_single:
                parts[idx] = self.latin_single[p.lower()]
                cnt += 1
        return "".join(parts), cnt

    def whole(self, s: str, deep: bool = True):
        """整格就是已知姓名（含兩個字、單字英文名；deep 時含「王明(退休)」這種帶註記的）時回傳代碼。"""
        t = s.strip()
        if not t:
            return None
        if _CJK.search(t):
            code = self.exact.get(re.sub(r"\s", "", t))
            if code is None and deep and len(t) <= 16:
                core = _person_core(t)
                code = self.exact.get(core) if core else None
            return code
        low = " ".join(t.translate(_APOSTROPHES).split()).lower()
        return self.latin_single.get(low) or self.latin_multi.get(low)

    def short_mentions(self, s: str) -> int:
        """句子裡出現兩個字的已知姓名、或整格等於單字英文名（沒有自動替換）的次數。"""
        n_ = 0
        if self.short_first is not None:
            m = self.short_first.search(s)
            while m:
                i = m.start()
                if s[i:i + 2] in self.short:
                    b, a_ = s[i - 1:i], s[i + 2:i + 3]
                    if (not b or not _CJK.match(b) or b in _SHORT_BEFORE
                            or not a_ or not _CJK.match(a_) or s.startswith(_SHORT_AFTER, i + 2)):
                        n_ += 1
                        m = self.short_first.search(s, i + 2)
                        continue
                m = self.short_first.search(s, i + 1)
        low = s.strip().lower()
        if low and low in self.latin_single:
            n_ += 1
        return n_


# 工作表名／檔名裡「確定像人名」的片段（改名用；比畫面遮蔽嚴格，避免把「紀錄」「高雄場」改掉）
_TITLE_NON_NAME_TAILS = tuple("場區隊部班團組室館處系所院校會社季期年版表冊單檔案論學報誌展賽營課科計畫"
                              "書簿圖函稿集錄據卷帳局")


def _title_person(tok: str) -> bool:
    t = tok.strip()
    if _latin_person(t) or (t.isupper() and _latin_person(t.title())):
        return True
    # 原住民族名、外籍譯名常用間隔號：「達魯斯‧巴萬」「卡照‧伊斯坦大」
    if re.fullmatch(f"[{_CJK_CLASS}]{{1,6}}[‧・·．•][{_CJK_CLASS}]{{1,8}}", t) and not _headerish(t):
        return True
    m = re.fullmatch(f"(?:{_title_role_alt()})([{_CJK_CLASS}]{{3,4}})", t)
    if m and _mention_name_ok(m.group(1)):     # 「原主持人伏家瑋」
        return True
    m = re.fullmatch(f"([{_CJK_CLASS}]{{2,4}})(?:老師|教授|先生|女士|小姐)", t)
    if m:
        t = m.group(1)
        return _cjk_person(t) and len(t) <= 3 or t[:2] in _COMPOUND_SURNAMES
    if not re.fullmatch(f"[{_CJK_CLASS}]{{3,4}}", t) or not _cjk_person(t):
        return False
    if len(t) == 4:
        return t[:2] in _COMPOUND_SURNAMES
    return not t.endswith(_TITLE_NON_NAME_TAILS) and not (set(t) & _NON_NAME_CHARS)


_title_role_cache = {}


def _title_role_alt():
    if _title_role_cache.get("src") is not NAME_KEYS:
        roles = sorted(set(_MENTION_EXTRA_ROLES) | {k for k in NAME_KEYS if re.fullmatch(f"[{_CJK_CLASS}]{{2,8}}", k)},
                       key=len, reverse=True)
        _title_role_cache["src"] = NAME_KEYS
        _title_role_cache["alt"] = "(?:原|前|新|現任)?(?:" + "|".join(map(re.escape, roles)) + ")"
    return _title_role_cache["alt"]


def mask_title(s: str) -> str:
    s = re.sub(f"(?<![{_CJK_CLASS}])([{_CJK_CLASS}]{{1,2}})[ 　]([{_CJK_CLASS}]{{1,3}})(?![{_CJK_CLASS}])",
               lambda mo: "＊＊" if 3 <= len(mo.group(1) + mo.group(2)) <= 4
               and _cjk_person(mo.group(1) + mo.group(2)) else mo.group(0), s)
    # 字間有空白的中文姓名（「李 大 華」）先合起來判斷
    s = re.sub(f"(?<![{_CJK_CLASS}])([{_CJK_CLASS}])[ 　]([{_CJK_CLASS}])[ 　]([{_CJK_CLASS}])(?![{_CJK_CLASS}])",
               lambda mo: "＊＊" if _cjk_person("".join(mo.groups())) else mo.group(0), s)
    s = re.sub(r"[A-Z][A-Z'\-]+(?: [A-Z][A-Z'\-]+){1,2}",
               lambda mo: "＊＊" if _latin_person(mo.group(0).title()) else mo.group(0), s)
    s = re.sub(_LATIN_NAME, lambda mo: "＊＊" if _latin_person(mo.group(0)) else mo.group(0), s)
    parts = re.split(r"([_\-\s（）()【】\[\]、，,.\\/]+)", s)
    return "".join("＊＊" if p and _title_person(p) else p for p in parts)


def _col_letter(ci: int) -> str:
    s = ""
    ci += 1
    while ci:
        ci, r = divmod(ci - 1, 26)
        s = chr(65 + r) + s
    return s


# ====================== 主處理 ======================

def process_sheet(df: pd.DataFrame, salt: str, mapping: dict, rev: dict, stats: dict,
                  dry: bool, known_before=frozenset(), reserved=()):
    """就地處理一張工作表的姓名欄與清空欄（df 為無標題讀入的原始矩陣）。

    回傳 (df, info)。姓名欄以外的已知姓名掃描在第二階段做（見 finish_excel）。
    """
    n, m = df.shape
    info = dict(touched=[], suspects=[], name_nonblank=0, replaced=0, blocks=0, doubtful=[],
                label_cols=[], done=np.zeros((n, m), dtype=bool), no_header=False,
                already_coded=0, names=[], mentions=0, mention_unsure=0, fallback=False, data_coded={}, kv_doubt=[], multi_cells=0)
    nbm = _nonblank_matrix(df)          # 處理前的有值格：清空、換碼之後判斷位置仍以原始內容為準
    lay = detect_layout(df, known_before, nbm)
    info["blocks"] = len(lay.blocks)
    info["doubtful"] = lay.doubtful_rows
    done = info["done"]

    # ---- 每一格要做什麼 ----
    # 多段判定 ∪ 保底判定：任何一種判定認為是姓名欄的格子都處理，多段判定只能「多做」。
    name_any = np.zeros((n, m), dtype=bool)
    drop_any = np.zeros((n, m), dtype=bool)
    touched = []

    def mark(start, stop, roles, labels, only_name=False):
        for ci, role in sorted(roles.items()):
            if only_name and role != "name":
                continue
            (name_any if role == "name" else drop_any)[start:stop, ci] = True
            touched.append((norm_header(labels.get(ci, ""))[:20], "假名化" if role == "name" else "清空"))
            if norm_key(labels.get(ci, "")) in CONFIG_ADDED:
                CONFIG_USED.add(norm_key(labels.get(ci, "")))

    for start, stop, bi in lay.segments:
        mark(start, stop, lay.blocks[bi][2], lay.blocks[bi][3])
    hit = lay.hit if lay.hit is not None else _hit_matrix(df)
    block_starts = sorted(b[0] for b in lay.blocks)
    for r, cols in lay.extra_name:
        stop = next((b for b in block_starts if b > r), n)
        mark(r + 1, stop, {ci: "name" for ci in cols}, {ci: str(df.iat[r, ci]) for ci in cols})
    data_rows = frozenset(r for start, stop, _bi in lay.segments for r in range(start, stop))
    for h, which in _legacy_marks(df, hit, data_rows):
        roles, labels = {}, {}
        for ci in range(m):
            v = df.iat[h, ci]
            if which == "legacy":
                # v1.2 的判斷與清單：完全相符（只去空白）才算，和 v2.4 處理的格子一模一樣
                key = norm_header(v) if isinstance(v, str) else None
                role = "name" if key in LEGACY_NAME_COLUMNS else "drop" if key in LEGACY_DROP_COLUMNS else None
            else:
                role = _role(v)
            if role:
                roles[ci], labels[ci] = role, str(v)
        # v1.2 列照舊處理到表尾；新清單的保底停在下一個採用的標題列（那一段有自己的欄位）
        stop = n if which == "legacy" else next((b for b in block_starts if b > h), n)
        if which != "legacy":
            # 簽核列「承辦人｜單位主管」底下是表名、序號：那一欄沒有像姓名的值就不採用
            roles = {ci: r for ci, r in roles.items() if r != "name" or _col_has_name(df, h + 1, stop, ci)}
        mark(h + 1, stop, roles, labels, only_name=(which == "name"))

    nb_rows = None

    def nbm_row_count(ri):
        nonlocal nb_rows
        if nb_rows is None:
            nb_rows = nbm.sum(axis=1)
        return int(nb_rows[ri])

    def clear_cell(ri, ci):
        if not dry:
            df.iat[ri, ci] = ""
            vals0[ri, ci] = ""

    def code_of(name):
        """查出或配發代碼（檢視模式不配發，回傳 None）。"""
        if dry:
            return None
        code = rev.get(name)
        if code is None:
            code = assign_code(name, salt, mapping, stats, reserved)
            rev[name] = code
            mapping[code] = name
            stats["new_names"] += 1
            info["names"].append(name)
        return code

    def code_cell(ri, ci, cv):
        # 姓名：和 v1.2 一樣，任何值（含數字、日期）都換成代碼，只有占位值與既有代碼保留
        val = str(cv).strip()
        if not val or val in ("nan", "-", "—") or val.lower() in _PLACEHOLDERS:
            return
        if _CODE_RE.fullmatch(val):
            info["already_coded"] += 1
            return
        if re.fullmatch(r"[【\[［(（].{1,12}[】\]］)）]", val) and nbm_row_count(ri) == 1:
            return                 # 「【甲組】」這種分組小標題
        name = canonical(val)      # 套用姓名整併表
        # 一格寫好幾個人、或姓名帶註記稱謂（王小明(召集人)）時，個別／核心姓名也配代碼，
        # 第二階段才能在備註、標題裡找到他們
        parts = [canonical(p) for p in _name_parts(val)] if _MULTI_HINT.search(val) else []
        if len(parts) >= 2:
            info["multi_cells"] += 1
        if not parts:
            core = _person_core(val)
            if core and core != val:
                parts = [canonical(core)]
        if dry:
            info["names"] += [name] + parts
            return
        df.iat[ri, ci] = vals0[ri, ci] = code_of(name)
        for p in parts:
            code_of(p)
        stats["cells"] += 1
        info["replaced"] += 1
        if _data_like(cv):
            info["data_coded"][ci] = info["data_coded"].get(ci, 0) + 1

    # 和 df 同步的值（寫入時兩邊一起改）；讀陣列比 df.iat 快很多。
    # 🔴 一定要 copy=True：pandas 3 單欄表格的 to_numpy() 是唯讀陣列，寫入會 ValueError（整個檔處理失敗）
    vals0 = df.to_numpy(dtype=object, copy=True)
    for ri, ci in zip(*np.nonzero(name_any | drop_any)):
        cv = vals0[ri, ci]
        done[ri, ci] = True
        if _is_blank(cv):
            continue
        is_name = name_any[ri, ci]
        if is_name:
            info["name_nonblank"] += 1
        # 等於欄名的格子（重複表頭、欄位說明）不是人名也不是個資，保留原樣
        if isinstance(cv, str) and _exact_role(cv):
            continue
        # 只標為清空；或兩種判定衝突（這段是學號欄、保底說是姓名欄）而值是數字或編號 → 清空
        if drop_any[ri, ci] and (not is_name or not isinstance(cv, str) or _data_like(cv)):
            clear_cell(ri, ci)
            continue
        code_cell(ri, ci, cv)
    info["touched"] = list(dict.fromkeys(touched))

    # 「欄名清單」型的欄位：同一欄出現好幾個欄名，而且有個資欄名（直式表單、異動紀錄、欄位說明）
    for ci in range(m):
        rows = np.flatnonzero(hit[:, ci])
        if len(rows) < 3:
            continue
        found = {_match_key(df.iat[r, ci]) for r in rows}
        if len(found) >= 2 and found & DROP_KEYS:
            info["label_cols"].append(ci)

    # ---- 鍵值對：左邊寫欄名、右邊寫值（「計畫主持人｜王小明」「電話｜0912…」）----
    # 只看兩種位置，避免誤傷一般表格：表頭上方的資訊列（右邊要像人名才換），以及直式表單的欄名欄。
    in_seg = np.zeros(n, dtype=bool)
    for start, stop, _bi in lay.segments:
        in_seg[start:stop] = True
    # 直式表單的欄名欄一定有「姓名類欄名＋右邊是人名」；核銷清冊的經費項目欄剛好寫「電話、手機」不算
    form_cols = {ci for ci in info["label_cols"] if ci + 1 < m and any(
        _role(df.iat[r, ci]) == "name" and looks_like_person(df.iat[r, ci + 1])
        for r in np.flatnonzero(hit[:, ci]))}
    for ri, ci in zip(*np.nonzero(hit)):
        vc = _right_value_col(nbm, ri, ci)          # 合併儲存格造成的空白往右跳過
        if vc is None or done[ri, vc]:
            continue
        v = df.iat[ri, vc]
        if _is_blank(v) or _exact_role(v):
            continue
        form = ci in form_cols
        role = _role(df.iat[ri, ci])
        above = not in_seg[ri] and ri not in lay.header_rows and bool(lay.segments)   # 表頭上方的資訊列
        if role == "drop":
            # 右邊是電話、Email、證號格式：任何位置都清空（兩列直式表單、表尾簽核列）；
            # 其他資料值只在直式表單或表頭上方資訊列清空（「經費項目｜電話｜1200」不能清）
            if _pii_like(v) or form or (above and _data_like(v)):
                done[ri, vc] = True
                clear_cell(ri, vc)
        elif role == "name":
            if ri in lay.header_rows and isinstance(v, str) and (
                    _headerish(v) or len(re.sub(r"\s", "", v)) <= 2 or suspect_label(v)):
                continue
            plain = isinstance(v, str) and not _data_like(v)
            short_name = plain and re.fullmatch(f"[{_CJK_CLASS}]{{2,4}}", v.strip()) and not _label_word(v)                 and not v.strip().endswith(_ORG_TAILS + _ITEM_TAILS)
            if looks_like_person(v) or (form and plain) or (above and short_name):
                done[ri, vc] = True
                info["name_nonblank"] += 1
                code_cell(ri, vc, v)
                if not form and nbm[ri].sum() >= 3 and _names_below(df, nbm, done, ri, ci):
                    info["kv_doubt"].append(int(ri))

    # 清單外的人員欄名（「製表人｜程瑜蓉」「單位主管｜林志明」）：只看表頭上方或沒有表頭的稀疏列，右邊要像人名
    for ri in range(n):
        if in_seg[ri] or ri in lay.header_rows or not 2 <= nbm[ri].sum() <= 6:
            continue
        for ci in np.flatnonzero(nbm[ri]):
            k = df.iat[ri, ci]
            if hit[ri, ci] or not isinstance(k, str):
                continue
            kk = norm_key(k)
            if not (kk in _MENTION_EXTRA_ROLES or (len(kk) <= 5 and kk.endswith(("人", "者"))
                                                    and not _headerish(k)) or kk in ("單位主管", "主管")):
                continue
            vc = _right_value_col(nbm, ri, ci)
            if vc is None or done[ri, vc]:
                continue
            v = df.iat[ri, vc]
            if looks_like_person(v):
                done[ri, vc] = True
                info["name_nonblank"] += 1
                code_cell(ri, vc, v)

    for ri in range(n):
        if not 4 <= nbm[ri].sum() <= 10 or not done[ri].any() or ri in lay.header_rows:
            continue
        signed = any(done[ri, vc] and vc - 1 >= 0 and isinstance(vals0[ri, vc - 1], str)
                     and (hit[ri, vc - 1] or norm_key(vals0[ri, vc - 1]).endswith(("人", "主管")))
                     for vc in np.flatnonzero(done[ri]))
        if not signed:
            continue
        for ci in np.flatnonzero(nbm[ri] & ~done[ri]):
            v = vals0[ri, ci]
            if ci - 1 >= 0 and isinstance(vals0[ri, ci - 1], str) and not _data_like(vals0[ri, ci - 1]) \
                    and looks_like_person(v):
                done[ri, ci] = True
                code_cell(ri, ci, v)

    # ---- 文字裡的姓名寫法：「承辦人：王小明」「聯絡人 王小明、李大華」----
    role_re = _mention_role_re()
    for ci in range(m):
        col = df.iloc[:, ci]
        cand = _is_str_mask(col) & ~done[:, ci]
        if not cand.any():
            continue
        rows = np.flatnonzero(cand)
        hits_ = col.iloc[rows].map(lambda s_: role_re.search(s_) is not None
                                   or _HONOR_HINT.search(s_) is not None).to_numpy(dtype=bool)
        for ri in rows[hits_]:
            v = df.iat[ri, ci]
            if _exact_role(v) or _CODE_RE.fullmatch(v.strip()):
                continue
            unsure = []
            spans = _mention_spans(v, unsure)
            info["mention_unsure"] += len(unsure)
            if not spans:
                continue
            new = v
            for st, en, nm in reversed(spans):
                name = canonical(nm)
                if dry:
                    info["names"].append(name)
                code = code_of(name) if not dry else "T-（新代碼）"
                new = new[:st] + _glue(code, new[en:], " ") + new[en:]
            info["mentions"] += len(spans)
            df.iat[ri, ci] = vals0[ri, ci] = new  # 檢視模式不寫檔，但第二階段要看到換過的內容
            if not dry:
                stats["cells"] += 1

    info["fallback"] = lay.fallback
    if not touched or lay.fallback:
        info["no_header"] = not touched and n > 0 and m > 0
        # 沒有任何認得的欄名：仍然對前幾列「像標題」的列檢查疑似個資欄名
        nb = _nonblank_matrix(df.iloc[:MAX_SCAN]) if n else None
        for r in range(min(MAX_SCAN, n)):
            if nb[r].sum() >= 2 and _label_like_row(df, nb, r):
                info["suspects"] += [re.sub(r"\s+", " ", df.iat[r, ci]).strip()
                                     for ci in np.flatnonzero(nb[r]) if suspect_label(df.iat[r, ci])]
                break

    for start, stop, bi in lay.segments:
        _first, _last, roles, labels = lay.blocks[bi]
        # 疑似個資欄位：印出前還會過濾掉看起來像人名的標籤（見 finish_excel）
        for ci, lab in labels.items():
            if ci not in roles and suspect_label(lab):
                # 左邊緊鄰姓名欄名的格子多半是「姓名｜伏家生」的值，只印欄位位置
                left = roles.get(ci - 1)
                info["suspects"].append(f"第 {_col_letter(ci)} 欄" if left == "name" else lab)

    info["suspects"] = list(dict.fromkeys(info["suspects"]))
    return df, info


def scan_known_names(df: pd.DataFrame, done: np.ndarray, known: KnownNames, dry: bool):
    """姓名欄以外的格子若出現已知姓名，換成代碼。回傳 ({欄索引: 處數}, 兩字姓名提醒數)。"""
    hits, short = {}, 0
    if not known:
        return hits, short
    arr = df.to_numpy(dtype=object)
    todo = ~done
    # 同一欄有兩格以上、而且三成以上的格子「整格就是已知姓名」：多半是沒認出來的姓名欄，
    # 這一欄裡的兩字姓名、單字英文名也整格換（分類欄剛好有一格等於兩字姓名時不動）
    namecols = set()
    for ci in range(arr.shape[1]):
        rows = [ri for ri in np.flatnonzero(todo[:, ci]) if isinstance(arr[ri, ci], str) and arr[ri, ci].strip()]
        if len(rows) < 2:
            continue
        found = sum(1 for ri in rows if known.whole(arr[ri, ci], deep=False))
        if found >= 2 and found * 10 >= len(rows) * 3:
            namecols.add(ci)
    for ri, ci in zip(*np.nonzero(todo)):
        v = arr[ri, ci]
        if not isinstance(v, str) or not v.strip():
            continue
        t = v.strip()
        if _CODE_RE.fullmatch(t) or _exact_role(t):
            continue
        if ci in namecols:
            code = known.whole(v)
            if code is not None:
                hits[ci] = hits.get(ci, 0) + 1
                if not dry:
                    df.iat[ri, ci] = code
                continue
        new, cnt = known.replace(v)
        if cnt:
            hits[ci] = hits.get(ci, 0) + cnt
            if not dry:
                df.iat[ri, ci] = new
        short += known.short_mentions(new)
    return hits, short


_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _clean(v):
    # openpyxl 無法寫入 ASCII 控制字元（如 \x0b），這類字元偶爾混在
    # 計畫名稱等自由文字欄位中，會導致寫檔失敗。上傳前一併清除。
    if isinstance(v, str):
        return _ILLEGAL.sub("", v)
    return v


def _stem(name: str, suffix: str) -> str:
    """去掉副檔名（不能用 with_suffix：「114.09.15_名單」會被當成副檔名 .15_名單）。"""
    return name[:-len(suffix)] if suffix and name.lower().endswith(suffix.lower()) else name


def _out_rel(rel: Path, known) -> Path:
    parts = [known.replace_tokens(p)[0] if known else p for p in rel.parent.parts]
    stem = _stem(rel.name, rel.suffix)
    stem = known.replace_tokens(stem)[0] if known else stem
    return Path(*parts, stem + ".xlsx")


def _resolve_out_path(rel_out: Path, used: set, src_rel: str, sources: dict, planned=None):
    """若與「這次」或「先前批次」別的原始檔撞名，加上原副檔名區隔。回傳 (相對路徑, 是否新改名)。"""

    def taken(p: Path) -> bool:
        k = str(p).lower()
        if k in used:
            return True
        owner = sources.get(k)
        return owner is not None and owner.lower() != src_rel.lower() and (OUT_DIR / p).exists()

    renamed = False
    if taken(rel_out):
        key = str(rel_out).lower()
        owner = sources.get(key) or next((s for s, o in (planned or {}).items() if str(o).lower() == key), None)
        stem = _stem(rel_out.name, ".xlsx")
        src = Path(src_rel)
        if owner and _stem(Path(owner).name, Path(owner).suffix).lower() == _stem(src.name, src.suffix).lower():
            # 只差副檔名（查詢結果.xls／查詢結果.xlsx）：加上原副檔名區隔
            base, k = f"{stem}({src.suffix.lstrip('.').lower() or '檔'})", 2
            cand = rel_out.with_name(f"{base}.xlsx")
        else:
            # 檔名裡的姓名遮掉以後才撞名（113_王大明_簽到／113_李小華_簽到）：加編號
            base, k = stem, 3
            cand = rel_out.with_name(f"{stem}_2.xlsx")
        while taken(cand):
            cand = rel_out.with_name(f"{base}_{k}.xlsx")
            k += 1
        # 上一次就已經改存成這個檔名的話不再提示（重跑是常態，重複提醒會變成雜訊）
        renamed = (sources.get(str(cand).lower()) or "").lower() != src_rel.lower()
        rel_out = cand
    used.add(str(rel_out).lower())
    return rel_out, renamed


def _sheet_out_name(name: str, used: set) -> str:
    """Excel 工作表名最長 31 字；截斷時不可切到代碼中間（切出查不回的片段）。"""
    def cut(s, limit):
        if len(s) <= limit:
            return s
        s = s[:limit]
        m = re.search(r"T-[0-9A-F]*$", s)
        if m and not re.fullmatch(r"T-[0-9A-F]{%d}" % CODE_LEN, m.group(0)):
            s = s[:m.start()].rstrip("_- ") or "工作表"
        return s
    new = cut(name, 31)
    base, k = new, 2
    while new.lower() in used:
        new = f"{cut(base, 28)}~{k}"
        k += 1
    used.add(new.lower())
    return new


def _move_to_quarantine(paths, stats, label):
    """移到 _private\\output_隔離\\<時間>\\（同一個磁碟，用 os.replace，失敗時不會留下複本）。"""
    dest = PRIV_DIR / "output_隔離" / dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    failed = []
    for p in paths:
        rel = p.relative_to(OUT_DIR)
        try:
            (dest / rel).parent.mkdir(parents=True, exist_ok=True)
            os.replace(p, dest / rel)
            stats[label] += 1
        except OSError:
            failed.append(rel)
    stats["quarantine_failed"] += len(failed)
    return dest, failed


def quarantine_nonxlsx(dry: bool, stats: dict, known, later="按「進行假名化」時"):
    """output 裡「不是本工具產生的格式」（例如舊版原樣複製進去的 .csv）移出 output。"""
    if not OUT_DIR.exists():
        return
    risky = [p for p in sorted(OUT_DIR.rglob("*"))
             if p.is_file() and not p.name.startswith("~$") and p.name.lower() not in SYSTEM_FILES
             and p.suffix.lower() != ".xlsx"]
    if not risky:
        return
    print("\n" + "!" * 66)
    print(f"  ! output 裡有 {len(risky)} 個不是本工具產生的檔案（可能是舊版原樣複製進去的原始檔，")
    print("可能含真實姓名與個資）：")
    for p in risky[:20]:
        print(f"  ・{mask_text(str(p.relative_to(OUT_DIR)), known)}")
    if len(risky) > 20:
        print(f"  ...（其餘 {len(risky) - 20} 個）")
    if dry:
        stats["quarantine_pending"] += len(risky)
        print(f"{later}，會先把它們移到 _private\\output_隔離\\（不會刪除）。")
        print("!" * 66)
        return
    dest, failed = _move_to_quarantine(risky, stats, "quarantined")
    if stats["quarantined"]:
        print(f"已移到 {dest}（沒有刪除）。")
    if failed:
        print("  !! 下列檔案搬移失敗（可能正被開啟），仍在 output 裡，請勿上傳：")
        for r in failed[:20]:
            print(f"  ・{mask_text(str(r), known)}")
    print("如果先前曾把 output 整包上傳，請確認這些檔案是否含個資，必要時通報主管。")
    print("!" * 66)


def _ensure_utf8_stdout():
    """讓輸出被重導向到檔案或管線時也不會因為 cp950 無法編碼而中斷。
    視窗介面會把 stdout 換成自己的物件，那種情況下這裡會安靜略過。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass


def _new_stats() -> dict:
    return {"files": 0, "cells": 0, "new_names": 0, "failed": 0, "failed_files": [],
            "skipped": 0, "collisions": 0, "suspect_files": 0, "zero_replaced_files": 0,
            "no_name_files": 0, "name_hits_elsewhere": 0, "short_name_mentions": 0, "short_name_files": 0,
            "structure_warn_files": 0, "already_coded_files": 0, "person_names_in_titles": 0,
            "config_problems": 0,
            "quarantined": 0, "stale_quarantined": 0, "quarantine_failed": 0,
            "quarantine_pending": 0, "legacy_unverified": 0, "renamed_outputs": 0,
            "mapping_locked": False, "stop_reason": "", "mapping_missing": False,
            "sources_problem": False, "not_processed": 0, "mapping_rewritten": False,
            "name_dups": 0, "warnings_locked": False, "earlier_hits": 0}


def _finish_stats(stats: dict) -> dict:
    return stats


def _has_person_cell(df) -> bool:
    for v in df.to_numpy(dtype=object).ravel():
        if isinstance(v, str) and (looks_like_person(v) or _mention_spans(v)):
            return True
    return False


def finish_excel(entry, known, stats, dry, used_outputs, sources, planned):
    """第二階段：掃描姓名欄以外的已知姓名、改工作表名與檔名、印結果、寫檔。"""
    rel, sheet_names, sheets, infos = entry["rel"], entry["sheet_names"], entry["sheets"], entry["infos"]
    elsewhere, short = [], 0
    for sh in sheet_names:
        h, s = scan_known_names(sheets[sh], infos[sh]["done"], known, dry)
        elsewhere += [(sh, ci, cnt) for ci, cnt in sorted(h.items())]
        short += s

    used_sheet, out_names, renamed_sheets, person_sheets = set(), {}, 0, 0
    for idx, sh in enumerate(sheet_names):
        new, cnt = known.replace_tokens(sh) if known else (sh, 0)
        renamed_sheets += bool(cnt)
        if mask_title(new) != new:
            # 沒出現在姓名欄、但看起來像人名的工作表名（一位老師一個分頁）：output 改成編號
            new = f"工作表{idx + 1}"
            person_sheets += 1
        out_names[sh] = _sheet_out_name(new, used_sheet)
    # 畫面上的工作表名一律再遮蔽一次（output 只改確定的人名，畫面寧可多遮）
    disp = {sh: mask_text(out_names[sh], known) for sh in sheet_names}

    touched_any = False
    for sh in sheet_names:
        info = infos[sh]
        if info["touched"]:
            touched_any = True
            cols = [f"{mask_text(c, known)}（{a}）" for c, a in info["touched"]]
            more = f"…等 {len(cols)} 欄" if len(cols) > 20 else ""
            print(f"  · [{disp[sh]}] {'、'.join(cols[:20])}{more}")
            if info["blocks"] > 1:
                print(f"    （這張表偵測到 {info['blocks']} 段表格，各段分別判斷欄位）")
    mentions = sum(infos[sh]["mentions"] for sh in sheet_names)
    if mentions:
        verb = "執行時會換成代碼" if dry else "已換成代碼"
        print(f"  · 文字裡「承辦人：○○○」「○○○老師」這類寫法的姓名 {mentions} 處，{verb}")

    problem = False
    no_header = [sh for sh in sheet_names if infos[sh]["no_header"] and sheets[sh].notna().to_numpy().sum() >= 2]
    if not touched_any:
        print("  ! 這個檔案完全沒有偵測到姓名欄位——請人工確認它是否真的不含姓名；")
        print("    若有姓名欄，請把欄名寫進 欄位設定.txt 後重跑。")
        stats["no_name_files"] += 1
        problem = True
    elif no_header:
        risky = [sh for sh in no_header if infos[sh]["suspects"] or _has_person_cell(sheets[sh])]
        quiet = [sh for sh in no_header if sh not in risky]
        if risky:
            print("  ! 下列工作表沒有偵測到姓名欄位，已原樣保留，請人工確認：" +
                  "、".join(disp[sh] for sh in risky[:10]))
            stats["no_name_files"] += 1
            problem = True
        if quiet:
            print("  · 下列工作表沒有姓名欄位，也沒有看起來像人名的內容，原樣保留：" +
                  "、".join(disp[sh] for sh in quiet[:10]))

    zero = [sh for sh in sheet_names if infos[sh]["name_nonblank"] >= 2 and
            (len(infos[sh]["names"]) if dry else infos[sh]["replaced"]) == 0 and not infos[sh]["already_coded"]]
    if zero:
        print("  ! 偵測到姓名欄，但一格都沒有換成代碼（可能是直式表單或欄名清單）：" +
              "、".join(disp[sh] for sh in zero[:10]) + "——請人工確認")
        stats["zero_replaced_files"] += 1
        problem = True

    suspects, hidden = [], 0
    for sh in sheet_names:
        for lab in infos[sh]["suspects"]:
            flat = norm_header(lab)
            if mask_text(lab, known) != lab or mask_text(flat, known) != flat:
                hidden += 1                     # 看起來是資料（人名）被當成欄名，不照印
            elif not lab.isascii():
                shown = re.split(r"[（(【\[：:]", flat)[0]
                bare = re.fullmatch(f"[{_CJK_CLASS}]{{2,4}}", shown) and not any(
                    w in shown for w in _PERSON_WORDS + _PII_WORDS + _HEADERISH_WORDS)
                if bare:
                    hidden += 1                 # 「仉志人」這種可能是罕見姓名的格子，不照印
                else:
                    suspects.append((shown + ("（…）" if shown != flat else ""))[:20])
            else:
                suspects.append(lab[:40])
    suspects = list(dict.fromkeys(suspects))
    if suspects or hidden:
        print("  ! 下列欄位看起來可能含姓名或個資，但不在處理清單內，已「原樣保留」：")
        if suspects:
            print("    " + "、".join(suspects[:15]) + ("…" if len(suspects) > 15 else ""))
        if hidden:
            print(f"    （另有 {hidden} 個標題格的內容看起來像人名，表格結構可能特殊）")
        print("    若確實需要處理，請把欄名加進 欄位設定.txt 後重跑（寫法見使用說明）。")
        problem = True

    if elsewhere:
        total = sum(c for _s, _c, c in elsewhere)
        where = [f"[{disp[sh]}] 第 {_col_letter(ci)} 欄 {cnt} 處" for sh, ci, cnt in elsewhere[:10]]
        verb = "執行時會一併換成代碼" if dry else "已一併換成代碼"
        print(f"  ! 在姓名欄以外找到 {total} 處已知姓名，{verb}：")
        print("    " + "、".join(where) + ("…" if len(elsewhere) > 10 else ""))
        print("    這些位置附近可能還有工具不認得的姓名，請打開 output 人工確認。")
        stats["name_hits_elsewhere"] += total
        problem = True
    unsure = sum(infos[sh]["mention_unsure"] for sh in sheet_names)
    if unsure:
        print(f"  ! 文字裡有 {unsure} 處「承辦人：○○」「○○老師」「感謝○○○老師」這類寫法可能是姓名，"
              "工具沒有換（兩個字或前面緊接其他字，容易誤傷一般詞彙），請人工確認。")
        stats["short_name_mentions"] += unsure
        problem = True
    if short:
        print(f"  ! 句子裡有 {short} 處出現兩個字的已知姓名，工具沒有自動替換（避免誤傷一般詞彙），請人工確認。")
        stats["short_name_mentions"] += short
        problem = True
    if unsure or short:
        stats["short_name_files"] += 1

    label_cols = [(sh, ci) for sh in sheet_names for ci in infos[sh]["label_cols"]]
    if label_cols:
        print("  ! 下列欄位的內容看起來是「欄位名稱」（直式表單、異動紀錄或欄位說明），")
        print("    旁邊欄位裡的姓名、電話、地址工具無法判斷，請人工確認：")
        print("    " + "、".join(f"[{disp[sh]}] 第 {_col_letter(ci)} 欄" for sh, ci in label_cols[:10]))
        stats["structure_warn_files"] += 1
        problem = True

    doubtful = [(sh, infos[sh]["doubtful"]) for sh in sheet_names if infos[sh]["doubtful"]]
    if doubtful:
        rows = "；".join(f"[{disp[sh]}] 第 {'、'.join(str(r + 1) for r in rs[:5])} 列" for sh, rs in doubtful[:5])
        print(f"  ! {rows} 看起來像另一段表格的標題列，但無法確定（含姓名欄名的，底下的姓名欄已照樣換成代碼），"
              "請人工確認這幾列以下的姓名與個資是否都處理到。")
        if not label_cols:
            stats["structure_warn_files"] += 1
        problem = True

    kv_doubt = [(sh, sorted(set(infos[sh]["kv_doubt"]))) for sh in sheet_names if infos[sh]["kv_doubt"]]
    if kv_doubt:
        rows = "；".join(f"[{disp[sh]}] 第 {'、'.join(str(r + 1) for r in rs[:5])} 列" for sh, rs in kv_doubt[:5])
        print(f"  ! {rows} 是「欄名｜姓名」寫法，姓名已換成代碼；但這一列也可能是另一段表格的標題列，"
              "底下同一欄像姓名的格子沒有處理，請人工確認。")
        if not label_cols and not doubtful:
            stats["structure_warn_files"] += 1
        problem = True

    data_coded = [(sh, ci, c) for sh in sheet_names for ci, c in sorted(infos[sh]["data_coded"].items()) if c >= 3]
    if data_coded:
        print("  ! 下列姓名欄裡有編號或數字，已照樣換成代碼（可能是表頭判斷把別的欄當成姓名欄），請確認："
              + "、".join(f"[{disp[sh]}] 第 {_col_letter(ci)} 欄 {c} 格" for sh, ci, c in data_coded[:10]))
        problem = True
    fallback = [sh for sh in sheet_names if infos[sh]["fallback"]]
    if fallback:
        print("  ! 下列工作表找不到明確的標題列（被當成標題的那一列含資料值），姓名與個資可能沒有處理到，請人工確認："
              + "、".join(disp[sh] for sh in fallback[:10]))
        if not label_cols and not doubtful and not kv_doubt:
            stats["structure_warn_files"] += 1
        problem = True

    multi = sum(infos[sh]["multi_cells"] for sh in sheet_names)
    if multi:
        verb = "執行時會整格換成一個代碼" if dry else "已整格換成一個代碼"
        print(f"  ! 姓名欄有 {multi} 格一格寫了兩個人以上（例如「○○○、○○○」），{verb}；"
              "這個代碼和他們各自單獨出現時的代碼不同（查詢時會查回整格原文）。"
              "若需要逐人對應，請拆成一人一格後重跑。")
        problem = True

    coded = sum(infos[sh]["already_coded"] for sh in sheet_names)
    if coded:
        print(f"  ! 姓名欄裡有 {coded} 格已經是代碼（這個檔可能已經假名化過），保留原樣，請確認沒有放錯檔案。")
        stats["already_coded_files"] += 1
        problem = True

    if renamed_sheets:
        print(f"    （{renamed_sheets} 個工作表名稱含已知姓名，output 中已改成代碼）")

    rel_out = _out_rel(rel, known)
    masked_out = Path(mask_title(str(rel_out)))
    if person_sheets or masked_out != rel_out:
        print("  ! 工作表名稱或檔名看起來含人名（這些人沒有出現在姓名欄，工具無法換成代碼），"
              "output 中已改成「工作表N」或「＊＊」，請確認。")
        stats["person_names_in_titles"] += 1
        problem = True
        rel_out = masked_out
    if problem:
        stats["suspect_files"] += 1

    natural = _out_rel(rel, None)
    rel_out, renamed = _resolve_out_path(rel_out, used_outputs, str(rel), sources, planned)
    planned[str(rel)] = rel_out
    # 沒有紀錄這個 output 是哪個原始檔產生的（舊版升級、output_sources.csv 不見／損毀／被鎖）：
    # 可能是只差副檔名的另一份資料，不能直接覆蓋
    orphan = (OUT_DIR / rel_out).is_file() and str(rel_out).lower() not in sources
    if dry:
        if orphan:
            stats["quarantine_pending"] += 1
            print(f"  ! output 裡已經有「{mask_text(str(rel_out), known)}」，但沒有紀錄它是哪個原始檔產生的；"
                  "執行時會先把這個舊檔移到 _private\\output_隔離\\（不會刪除），再寫入新的結果。")
        return
    if orphan:
        dest, failed = _move_to_quarantine([OUT_DIR / rel_out], stats, "quarantined")
        if not failed:
            print(f"  ! output 裡原本就有「{mask_text(str(rel_out), known)}」，但沒有紀錄它是哪個原始檔產生的"
                  "（舊版升級、或 output_sources.csv 不見／損毀）；")
            print(f"    覆蓋前已把舊檔移到 {dest}（沒有刪除）。若它其實是另一份資料（例如只差副檔名的檔案），請從那裡取回。")
    if renamed:
        stats["renamed_outputs"] += 1
    if renamed or str(rel_out).lower() != str(natural).lower():
        print(f"    （output 存為：{mask_text(str(rel_out))}）")
    out_path = OUT_DIR / rel_out

    for sh in sheets:
        df0 = sheets[sh]
        for col in df0.columns:
            if df0[col].dtype == object:
                df0[col] = df0[col].map(_clean)
    try:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(out_path, engine="openpyxl") as w:
            for sh, df in sheets.items():
                df.to_excel(w, sheet_name=out_names[sh], header=False, index=False)
                ws = w.sheets[out_names[sh]]
                for row in ws.iter_rows():
                    for c in row:
                        if c.data_type == "f":          # 讀進來的都是值；「=」開頭的是文字
                            c.data_type = "s"
    except PermissionError:
        stats["failed"] += 1
        stats["failed_files"].append(mask_text(str(rel_out), known))
        print(f"  !! 寫檔失敗：{mask_text(str(rel_out), known)} 可能正被 Excel 開啟，請關閉後重跑這個檔案")
        return
    except Exception as e:  # noqa: BLE001
        stats["failed"] += 1
        stats["failed_files"].append(mask_text(str(rel_out), known))
        print(f"  !! 寫檔失敗：{type(e).__name__}")
        return
    sources[str(rel_out).lower()] = str(rel)
    stats["files"] += 1


def review_leftovers(dry, stats, known, sources_at_start, planned, entries):
    """output 裡這次沒有寫到的 .xlsx：含已知姓名的檔名、同一原始檔改名留下的舊檔 → 隔離；
    沒有來源紀錄的舊檔 → 檢查內容，查不出來就計入需要確認。"""
    if not OUT_DIR.exists():
        return
    produced = {str(p).lower() for p in planned.values()}
    processed = {str(e["rel"]).lower() for e in entries if e["kind"] == "ok"}
    failed_src = {str(e["rel"]).lower() for e in entries if e.get("write_failed")}
    stale, unverified, earlier, earlier_hit = [], [], 0, []
    for q in sorted(OUT_DIR.rglob("*.xlsx")):
        if not q.is_file() or q.name.startswith("~$"):
            continue
        rel_l = str(q.relative_to(OUT_DIR))
        if rel_l.lower() in produced:
            continue
        owner = sources_at_start.get(rel_l.lower())
        path_has_name = known and any(known.replace_tokens(p)[1] for p in Path(rel_l).parts)
        if path_has_name:
            stale.append(q)
        elif owner is not None and owner.lower() in processed and owner.lower() not in failed_src:
            stale.append(q)                       # 同一個原始檔這次改存成別的檔名，這是舊檔
        elif owner is not None:
            earlier += 1                          # 先前批次的成果，來源這次不在 input
            if known and not dry:
                try:
                    for df in pd.read_excel(q, sheet_name=None, header=None, dtype=object).values():
                        if any(isinstance(v, str) and known.replace(v)[1] for v in df.to_numpy().ravel()):
                            earlier_hit.append(q)
                            break
                except Exception:  # noqa: BLE001
                    pass
        else:
            hit = False
            if known:
                try:
                    for df in pd.read_excel(q, sheet_name=None, header=None, dtype=object).values():
                        if any(isinstance(v, str) and known.replace(v)[1] for v in df.to_numpy().ravel()):
                            hit = True
                            break
                except Exception:  # noqa: BLE001  讀不了的檔一樣列為需確認
                    pass
            (stale if hit else unverified).append(q)
    if stale:
        print("\n" + "!" * 66)
        print(f"output 裡有 {len(stale)} 個舊檔（檔名或內容含已知姓名，或同一個原始檔這次改存成別的檔名）：")
        for q in stale[:20]:
            print(f"  ・{mask_text(str(q.relative_to(OUT_DIR)), known)}")
        if dry:
            stats["quarantine_pending"] += len(stale)
            print("按「進行假名化」時，會把它們移到 _private\\output_隔離\\（不會刪除）。")
        else:
            dest, failed = _move_to_quarantine(stale, stats, "stale_quarantined")
            print(f"已移到 {dest}（沒有刪除）。")
            for r in failed[:20]:
                print(f"  !! 搬移失敗，仍在 output 裡，請勿上傳：{mask_text(str(r), known)}")
        print("!" * 66)
    if unverified:
        stats["legacy_unverified"] += len(unverified)
        print(f"\n  ! output 裡有 {len(unverified)} 個沒有來源紀錄的舊檔（可能是舊版或手動放進去的），"
              "工具無法確認內容是否都已假名化，請人工確認後再上傳"
              "（確認沒問題後，把原始檔放回 input 重跑一次，或把這些檔移出 output，這個提醒就會消失）：")
        for q in unverified[:20]:
            print(f"    ・{mask_text(str(q.relative_to(OUT_DIR)), known)}")
    if earlier:
        print(f"\n（output 裡另有 {earlier} 個先前批次產生的檔案，這次 input 沒有它們的原始檔。）")
    if earlier_hit:
        stats["earlier_hits"] += len(earlier_hit)
        print(f"\n  ! 先前批次的 {len(earlier_hit)} 個 output 裡，出現這次才在姓名欄看到的姓名（當時工具還不認得）：")
        for q in earlier_hit[:20]:
            print(f"    ・{mask_text(str(q.relative_to(OUT_DIR)), known)}")
        print("    請把這些檔案的原始檔放回 input，和這次的檔案一起重跑。")


def _stop(stats, e, n_files):
    stats["stop_reason"] = str(e).splitlines()[0].rstrip("。")
    stats["not_processed"] = n_files
    print("!" * 66)
    print("  !! " + str(e).replace("\n", "\n     "))
    print("  !! 這次沒有處理任何檔案。")
    print("!" * 66)
    return _finish_stats(stats)


def _check_private_consistency(mapping):
    """output 已經有成果、但對照表不見了：先停下，讓同仁從備份還原（事後才提醒時，output 已經被改寫）。"""
    if not (OUT_DIR.exists() and any(OUT_DIR.rglob("*.xlsx"))):
        return
    if not MAP_FILE.exists():
        # 不論 salt、sources 在不在都停：_private 整個不見、搬移中斷時也不能直接產生新鹽值改寫 output
        raise StopRun("output 裡已經有處理過的檔案，但 _private\\mapping.csv 不見了。\n"
                      "請先用備份還原 mapping.csv 再執行；直接執行會改寫 output，已上傳檔案的代碼可能永遠查不回。\n"
                      "（若 output 裡的檔案不要了，請先把它們移出 output。）")


def main(argv=None):
    """argv=None 時讀取命令列參數；視窗介面會直接傳入 ["--dry-run"] 或 []。"""
    _ensure_utf8_stdout()

    ap = argparse.ArgumentParser(description="上傳前假名化工具")
    ap.add_argument("--dry-run", action="store_true", help="只檢視，不寫檔")
    ap.add_argument("--decode", metavar="CODE", help="以對照表把代碼還原為姓名")
    args = ap.parse_args(argv)
    dry = args.dry_run

    if args.decode:
        code = args.decode.strip().upper()
        try:
            m = load_mapping()
        except StopRun as e:
            print(str(e))
            return
        if code in m:
            print(m[code])
        else:
            old = load_retired().get(code)
            print(f"{old}（這個代碼已整併，現在的代碼請查詢姓名整併表）" if old else "（查無此代碼）")
        return

    stats = _new_stats()
    stats["config_problems"] = load_column_config()
    for place, label in ((BASE, "專案資料夾"), (IN_DIR, "input 資料夾")):
        misplaced = place / "欄位設定.txt"
        if misplaced.exists() and misplaced.resolve() != COLUMN_CONFIG.resolve():
            print(f"  ! {label}裡有「欄位設定.txt」，但工具只讀工具資料夾（anonymize.py 所在的資料夾）裡的那一份，"
                  "這份沒有生效；請把它移到：" + str(TOOL_ROOT))
            stats["config_problems"] += 1

    files = [p for p in sorted(IN_DIR.rglob("*")) if p.is_file()
             and not p.name.startswith("~$") and p.name.lower() not in SYSTEM_FILES] if IN_DIR.exists() else []
    if not files:
        print(f"input 資料夾沒有可處理的檔案：{IN_DIR}")
        stats["stop_reason"] = "input 沒有可處理的檔案"
        return _finish_stats(stats)

    if not dry:
        # 上一次中斷留下的暫存對照表（內含姓名）一律清掉
        try:
            _tmp_path(MAP_FILE).unlink(missing_ok=True)
        except OSError:
            pass
        if not mapping_writable():
            stats["mapping_locked"] = True
            stats["not_processed"] = len(files)
            print("!" * 66)
            print("  !! 對照表 mapping.csv 目前無法寫入（通常是正被 Excel 開著）。")
            print("為了避免 output 出現查不回姓名的代碼，這次沒有處理任何檔案。")
            print("請先關閉開著 mapping.csv 的程式，再重新執行。")
            print("!" * 66)
            quarantine_nonxlsx(True, stats, None, later="關閉 mapping.csv 後重新執行時")
            return _finish_stats(stats)

    try:
        mapping = load_mapping()
        _check_private_consistency(mapping)
        salt = load_salt(dry, mapping)
    except StopRun as e:
        return _stop(stats, e, len(files))
    if MAPPING_ENC != "UTF-8" and MAP_FILE.exists():
        why = ("有「代碼、姓名」以外的欄位（例如自己加的備註）" if MAPPING_ENC == "多餘欄位"
               else f"是 {MAPPING_ENC} 編碼")
        later = "執行時存檔" if dry else "這次存檔"
        print(f"  ! mapping.csv {why}；{later}會改成只有「代碼、姓名」兩欄的 UTF-8，"
              "改寫前會先把原檔備份成 _private\\mapping.csv.bak_<日期時間>。")
        stats["mapping_rewritten"] = True

    global ALIASES
    try:
        ALIASES, cycles = load_aliases()
        retired = load_retired()
    except StopRun as e:
        return _stop(stats, e, len(files))
    if cycles:
        print(f"  ! 姓名整併表有 {len(cycles)} 條規則形成循環（A 併入 B、B 又併入 A），這些規則沒有套用，請修正。")
        stats["config_problems"] += len(cycles)
    if ALIASES:
        # 被整併掉的舊代碼從對照表移出，但保留在 retired_codes.csv：
        # 先前已上傳的 output 仍可還原，而且這個代碼不會再配發給別人
        stale = {c: n for c, n in mapping.items() if n in ALIASES}
        if stale and not dry:
            try:
                save_retired({**retired, **stale})      # 先確定寫得進去，才從對照表移除
            except OSError:
                return _stop(stats, StopRun("retired_codes.csv 無法寫入（可能正被 Excel 開著或設成唯讀）。\n"
                                            "姓名整併需要先把舊代碼記進這個檔，請關閉它後重新執行。"), len(files))
        for c in stale:
            del mapping[c]
        retired.update(stale)
        moved = (f"，{len(stale)} 個舊代碼移到 retired_codes.csv（仍可還原）" if not dry else
                 f"，執行時會把 {len(stale)} 個舊代碼移到 retired_codes.csv") if stale else ""
        print(f"[姓名整併] 已套用 {len(ALIASES)} 條整併規則{moved}")

    rev = {v: k for k, v in mapping.items()}
    sources, sources_ok = load_sources()
    if not sources_ok:
        stats["sources_problem"] = True
        print("  ! _private\\output_sources.csv 讀不到，這次無法判斷分批處理時的同名檔案。")
    sources_at_start = dict(sources)

    print("=" * 66)
    print(f"{'【檢視模式】' if dry else '【執行】'} 共 {len(files)} 個檔案")
    print(f"已知姓名代碼：{len(mapping)} 筆")
    short_codes = sum(1 for c in mapping if len(c) < 2 + CODE_LEN)
    if short_codes:
        print(f"（其中 {short_codes} 筆是舊版建立的 6 碼代碼，沿用不變以維持跨年一致；新加入的人使用 10 碼）")
    if OUT_DIR.exists() and any(OUT_DIR.rglob("*.xlsx")) and mapping and not SOURCES_FILE.exists():
        stats["sources_problem"] = True
        print("  ! 這個專案沒有 _private\\output_sources.csv（從舊版升級，或被刪除），這次無法判斷分批處理時的同名檔案；"
              "只差副檔名的同名檔可能會覆蓋先前的 output。")
    print("=" * 66)

    # ---- 第一階段：讀檔、處理姓名欄（整批的代碼都配發完，才掃描其他地方） ----
    person_names = {n.strip() for n in rev if _is_person_value(n)}
    reserved = dict(retired)
    entries, dry_names = [], set()
    for p in files:
        rel = p.relative_to(IN_DIR)
        ext = p.suffix.lower()
        if ext not in EXCEL_EXT:
            entries.append(dict(kind="skip", rel=rel, ext=ext))
            continue
        try:
            with pd.ExcelFile(p) as xl:
                sheet_names = list(xl.sheet_names)
                raw = {sh: pd.read_excel(xl, sheet_name=sh, header=None, dtype=object, keep_default_na=False,
                                         na_values=[""]) for sh in sheet_names}
        except Exception as e:  # noqa: BLE001
            entries.append(dict(kind="openfail", rel=rel, err=type(e).__name__))
            continue
        try:
            infos = {}
            known_before = frozenset(person_names | dry_names)
            for sh in sheet_names:
                raw[sh], infos[sh] = process_sheet(raw[sh], salt, mapping, rev, stats, dry,
                                                   known_before, reserved)
                for nm in infos[sh]["names"]:
                    if _is_person_value(nm):
                        (dry_names if dry else person_names).add(nm.strip())
            entries.append(dict(kind="ok", rel=rel, sheet_names=sheet_names, sheets=raw, infos=infos))
        except Exception as e:  # noqa: BLE001  單一檔案出錯不拖垮整批，但一定要講清楚
            entries.append(dict(kind="fail", rel=rel, err=f"{type(e).__name__}"))

    preview = dict(rev)
    if dry:
        for nm in dry_names:
            preview.setdefault(nm, "T-（新代碼）")
    known = KnownNames(preview)

    if not dry:
        try:
            save_mapping(mapping)       # 寫出任何 output 之前先存對照表
        except MappingSaveError as e:
            stats["mapping_locked"] = True
            stats["not_processed"] = sum(1 for e_ in entries if e_["kind"] == "ok")
            print("  !! 對照表 mapping.csv 存檔失敗，處理中止（沒有寫出任何 output）")
            print(f"     原因：{type(e).__name__}。請關閉開著 mapping.csv 的程式（通常是 Excel）後重新執行。")
            return _finish_stats(stats)

    quarantine_nonxlsx(dry, stats, known)

    # ---- 第二階段：逐檔掃描、改名、寫檔 ----
    used_outputs, planned, skipped_list = set(), {}, []
    for entry in entries:
        rel = entry["rel"]
        print(f"\n▸ {mask_text(str(rel), known)}")
        if entry["kind"] == "skip":
            stats["skipped"] += 1
            skipped_list.append(mask_text(str(rel), known))
            if entry["ext"] == ".pdf":
                print("  ! PDF 工具無法處理，不會放進 output。若內含姓名，請自行評估是否必要提供。")
            else:
                print("  ! 非試算表格式，工具無法處理，不會放進 output。")
            continue
        if entry["kind"] in ("openfail", "fail"):
            stats["failed"] += 1
            stats["failed_files"].append(mask_text(str(rel), known))
            what = "無法開啟" if entry["kind"] == "openfail" else "處理時發生錯誤"
            print(f"  !! {what}，這個檔案沒有寫出 output（{entry['err']}）")
            continue
        before = stats["failed"]
        finish_excel(entry, known, stats, dry, used_outputs, sources, planned)
        entry["write_failed"] = stats["failed"] > before
        entry["sheets"] = None                  # 寫完就釋放記憶體

    if not dry and not save_sources(sources):
        stats["sources_problem"] = True
        print("  ! _private\\output_sources.csv 無法寫入（可能正被開啟）。下次分批處理時，同名檔案可能會蓋掉這次的 output。")

    review_leftovers(dry, stats, known, sources_at_start, planned, entries)

    # 設定檔是所有專案共用的，沒用到不一定是寫錯，只提醒、不計入問題
    unused = [v for k, v in CONFIG_ADDED.items() if k not in CONFIG_USED]
    if unused:
        print(f"\n（欄位設定.txt 裡有 {len(unused)} 個欄名這次沒有出現在任何表頭：{'、'.join(unused[:10])}"
              "；若是寫法和 Excel 不同，請修正）")

    # ---- 未處理檔案：這是最容易讓真名外流的一條路，講清楚 ----
    if skipped_list:
        print("\n" + "!" * 66)
        print(f"下列 {len(skipped_list)} 個檔案「沒有」被處理，也沒有放進 output：")
        for r in skipped_list[:20]:
            print(f"  ・{r}")
        if len(skipped_list) > 20:
            print(f"  ...（其餘 {len(skipped_list) - 20} 個）")
        print("  ! 它們仍是原始資料，可能含真實姓名與個資，請勿上傳。")
        print("若需要處理，請先用 Excel 另存成 .xlsx 再放進 input 重跑。")
        print("!" * 66)

    # ---- 姓名寫法異常偵測（畫面只顯示代碼，姓名寫進檔案）----
    groups = {}
    for code, name in mapping.items():
        groups.setdefault(re.sub(r"\s", "", name), []).append((code, name))
    dups = {k: v for k, v in groups.items() if len(v) > 1}
    stats["name_dups"] = len(dups)
    warn_file = PRIV_DIR / "name_warnings.csv"
    if dups:
        print("\n" + "!" * 66)
        print(f"注意：偵測到 {len(dups)} 組疑似同一人的不同寫法（空白差異）")
        print("這會造成跨檔比對對不起來，建議請資料提供單位在來源端統一：")
        for k, v in list(dups.items())[:10]:
            print("  ・" + "　｜　".join(c for c, n in v))
        if len(dups) > 10:
            print(f"  ...（其餘 {len(dups) - 10} 組）")
        print("為避免截圖外流，畫面只顯示代碼。")
        if not dry:
            try:
                PRIV_DIR.mkdir(parents=True, exist_ok=True)
                with warn_file.open("w", encoding="utf-8-sig", newline="") as f:
                    w = csv.writer(f)
                    w.writerow(["正規化姓名", "代碼", "原始寫法"])
                    for k, v in dups.items():
                        for c, n in v:
                            w.writerow([k, c, n])
                print(f"對應的姓名請開啟 {warn_file}（該檔不可外流）。")
            except OSError:
                stats["warnings_locked"] = True
                print("  ! name_warnings.csv 正被開啟，這次沒有更新；請關閉後重跑。")
        print("!" * 66)
    elif not dry:
        # 全部整併完畢就清掉舊的警告檔，避免同仁誤讀過期資訊
        try:
            warn_file.unlink(missing_ok=True)
        except OSError:
            pass

    print("\n" + "=" * 66)
    if dry:
        print(f"檢視完成：掃描 {len(files)} 個檔案")
        print("（檢視模式不會寫出任何檔案，上面列出的是將會處理的欄位與需要確認的地方）")
    else:
        print(f"處理檔案　　：{stats['files']}")
        print(f"替換儲存格　：{stats['cells']}")
        print(f"本次新增姓名：{stats['new_names']}（累計 {len(mapping)}）")
        if stats["renamed_outputs"]:
            print(f"改名另存　　：{stats['renamed_outputs']} 個（和別的檔案同名，存檔名稱見上方「output 存為」）")
    if stats["skipped"]:
        print(f"未處理檔案　：{stats['skipped']}　← 不在 output，請勿上傳原檔")
    if stats["suspect_files"]:
        print(f"待人工確認　：{stats['suspect_files']} 個檔案（見上方 ! 開頭的說明）")
    if stats["config_problems"]:
        print(f"設定問題　　：{stats['config_problems']} 項沒有生效（見上方說明）")
    if stats["quarantined"] or stats["stale_quarantined"] or stats["quarantine_failed"] or stats["quarantine_pending"]:
        print(f"output 舊檔　：隔離 {stats['quarantined'] + stats['stale_quarantined']} 個"
              + (f"，搬移失敗 {stats['quarantine_failed']} 個" if stats["quarantine_failed"] else "")
              + (f"，待隔離 {stats['quarantine_pending']} 個" if stats["quarantine_pending"] else ""))
    if stats["legacy_unverified"]:
        print(f"待確認舊檔　：{stats['legacy_unverified']} 個（沒有來源紀錄）")
    if stats["collisions"]:
        print(f"代碼碰撞　　：{stats['collisions']}（已自動改配；請務必備份 mapping.csv，遺失後重建可能對不上）")
    if stats["failed"]:
        print(f"處理失敗　　：{stats['failed']}　← 這些檔案請勿上傳")
        for nm in stats["failed_files"][:10]:
            print(f"  ・{nm}")
    print("=" * 66)

    if not dry:
        print(f"可上傳的檔案在：{OUT_DIR}")
        print(f"請勿外流的檔案：{SALT_FILE}、{MAP_FILE}")

    # 回傳給視窗介面判斷要顯示「完成」還是警告；命令列執行時沒有人接收。
    return _finish_stats(stats)


if __name__ == "__main__":
    main()
