#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
====================================================================
 政大研發處　上傳前假名化工具　視窗介面  anonymize_gui.py  v2.9
====================================================================

用途
----
讓不熟悉命令列的同仁，用滑鼠完成假名化：
    ⓪ 選擇或新增專案
    ① 把檔案拖進該專案的 input 資料夾
    ② 按「進行假名化」
    ③ 按「檢視」核對結果，從 output 取出可上傳的檔案

多專案
------
同一位同仁可能同時處理好幾件工作。每個專案都是獨立的資料夾：

    假名化工具\\
    └── 專案\\
        ├── 國科會新進教師分析\\
        │   ├── input\\        原始檔
        │   ├── output\\       假名化後，可上傳
        │   └── _private\\     salt.txt、mapping.csv（不可外流）
        └── 產學合作經費分析\\
            ├── input\\ ...

每個專案有自己的 salt.txt，所以代碼各自獨立，互不相干。
同一個人在 A 專案和 B 專案會拿到不同代碼——這是刻意的，
可以避免不同業務的資料被串接起來。

反過來說，需要跨年追蹤的同一件業務，要一直用「同一個專案」，
不要每年新開一個，否則今年與去年的代碼會對不起來。

本程式不含任何處理邏輯，只是 anonymize.py 的外殼。
====================================================================
"""

import csv
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import traceback
from pathlib import Path

# ---- Windows 高解析度螢幕字體清晰化（放在建立視窗之前）----
if sys.platform == "win32":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog, ttk
except ImportError:
    sys.exit("此 Python 未包含 tkinter，請改用 python.org 的官方安裝版。")

# ---- 路徑：打包成 exe 後要用 exe 所在位置，不能用 __file__ ----
if getattr(sys, "frozen", False):
    ROOT = Path(sys.executable).resolve().parent
else:
    ROOT = Path(__file__).resolve().parent

PROJECTS_ROOT = ROOT / "專案"
LAST_FILE = ROOT / "last_project.txt"   # 記住上次用哪個專案，不含任何個資

# 目前專案的三個資料夾；選定專案後才有值
PROJ_DIR = None
IN_DIR = None
OUT_DIR = None
PRIV_DIR = None

FONT = ("Microsoft JhengHei UI", 10)
FONT_B = ("Microsoft JhengHei UI", 10, "bold")
FONT_H = ("Microsoft JhengHei UI", 13, "bold")
FONT_LOG = ("Consolas", 10)

INK = "#1f2328"
MUTED = "#6b7280"
AMBER = "#b45309"
TEAL = "#0f766e"
RED = "#b91c1c"
BG = "#faf9f7"

ICON_FILE = ROOT / "假名化工具.ico"      # 視窗與捷徑圖示（Windows）

# CopyNoright 標章（政大研發處設計）。以 base64 內嵌，不依賴外部檔案，
# 打包成 exe 時也不必額外處理。
BADGE_B64 = """
    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
"""


# ====================== 存放位置檢查 ======================
# 對照表 mapping.csv 含真實姓名。若工具資料夾放在會自動同步的雲端
# 資料夾內，等於把對照表送出這台電腦，假名化的防護就被繞過了。

# 服務品牌名稱：出現在路徑任何位置都算
CLOUD_KEYS = {
    "dropbox": "Dropbox",
    "onedrive": "OneDrive",
    "google drive": "Google 雲端硬碟",
    "googledrive": "Google 雲端硬碟",
    "google 雲端硬碟": "Google 雲端硬碟",
    "icloud": "iCloud",
    "box sync": "Box",
    "box drive": "Box",
    "pcloud": "pCloud",
    "megasync": "MEGA",
    "nextcloud": "Nextcloud",
    "seafile": "Seafile",
    "syncthing": "Syncthing",
    "resilio": "Resilio Sync",
    "webstorage": "ASUS WebStorage",
    "堅果雲": "堅果雲",
    "百度网盘": "百度網盤",
    "百度網盤": "百度網盤",
}

# 一般詞彙：整段資料夾名稱相符才算。
# Google Drive 在繁體中文 Windows 掛出來的預設名稱沒有「Google」字樣
# （實測 H:\我的雲端硬碟），但若用子字串比對，「D:\雲端硬碟備份說明\」也會被當成雲端。
CLOUD_FOLDER_NAMES = {
    "我的雲端硬碟": "Google 雲端硬碟",
    "共用雲端硬碟": "Google 雲端硬碟",
    "雲端硬碟": "雲端硬碟",
    "my drive": "Google 雲端硬碟",
    "shared drives": "Google 雲端硬碟",
    "同步空間": "雲端同步資料夾",
    "其他電腦": "Google 雲端硬碟",
    "computers": "Google 雲端硬碟",
    "box": "Box",
    "synologydrive": "Synology Drive",
    "synology drive": "Synology Drive",
}

# 不是雲端，但同樣不該長期存放對照表的位置（使用說明明文要求不要放桌面和文件）。
# 用 Windows 查詢這台電腦「桌面／文件／下載」的實際位置，
# 被重新導向到 D:\Desktop 也抓得到，網路磁碟上剛好叫 Documents 的共用資料夾則不會誤判。
KNOWN_FOLDER_IDS = {
    "桌面": "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}",
    "「文件」資料夾": "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}",
    "「下載」資料夾": "{374DE290-123F-4565-9164-39C4925E467B}",
}


def _known_folder_paths():
    """回傳 {名稱: 實際路徑}；查不到時退回使用者資料夾底下的預設位置。"""
    home = Path(os.path.expanduser("~"))
    out = {"桌面": home / "Desktop", "「文件」資料夾": home / "Documents",
           "「下載」資料夾": home / "Downloads"}
    if sys.platform != "win32":
        return out
    try:
        import ctypes
        import uuid
        from ctypes import wintypes

        class GUID(ctypes.Structure):
            _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD),
                        ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]

        get_path = ctypes.windll.shell32.SHGetKnownFolderPath
        free = ctypes.windll.ole32.CoTaskMemFree
        for label, fid in KNOWN_FOLDER_IDS.items():
            guid = GUID.from_buffer_copy(uuid.UUID(fid).bytes_le)
            p = ctypes.c_wchar_p()
            if get_path(ctypes.byref(guid), 0, None, ctypes.byref(p)) == 0:
                out[label] = Path(p.value)
                free(p)
    except Exception:
        pass
    return out


def _is_network_path(path: Path) -> bool:
    s = str(path)
    if s.startswith("\\\\"):
        return True
    if sys.platform == "win32" and len(s) >= 2 and s[1] == ":":
        try:
            import ctypes
            return ctypes.windll.kernel32.GetDriveTypeW(s[:2] + "\\") == 4   # DRIVE_REMOTE
        except Exception:
            return False
    return False


def _under(path: Path, folder: Path) -> bool:
    p, f = str(path).lower().rstrip("\\/"), str(folder).lower().rstrip("\\/")
    return bool(f) and (p == f or p.startswith(f + "\\") or p.startswith(f + "/"))


def _registered_sync_roots():
    """Windows 登記的雲端同步根目錄（OneDrive 商務版、SharePoint／Teams 文件庫、Box、
    Google Drive 等用檔案總管整合的服務都會登記在這裡）。只讀取，不修改登錄檔。"""
    roots = []
    if sys.platform != "win32":
        return roots
    try:
        import winreg
        base = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SyncRootManager"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base) as k:
            for i in range(winreg.QueryInfoKey(k)[0]):
                sub = winreg.EnumKey(k, i)
                try:
                    with winreg.OpenKey(k, sub + r"\UserSyncRoots") as u:
                        for j in range(winreg.QueryInfoKey(u)[1]):
                            _n, val, _t = winreg.EnumValue(u, j)
                            if isinstance(val, str) and val:
                                roots.append((sub.split("!")[0], Path(val)))
                except OSError:
                    continue
    except OSError:
        pass
    return roots


def detect_cloud(path: Path, sync_roots=None):
    """回傳雲端服務名稱；不在雲端資料夾內則回傳 None。"""
    roots = _registered_sync_roots() if sync_roots is None else sync_roots
    for provider, root in roots:
        if _under(path, root):
            return f"雲端同步資料夾（{provider}）"
    for part in Path(path).parts[1:]:
        low = part.lower().rstrip("\\/")
        name = CLOUD_FOLDER_NAMES.get(low)
        if name:
            return name
        # 品牌名稱是整段名稱，或開頭後面接空白、連字號、括號（「OneDrive - 政大」）；
        # 「OneDrive使用說明」「dropboxes」不算
        for key, name in CLOUD_KEYS.items():
            if low == key or re.match(re.escape(key) + r"[\s\-_(（]", low):
                return name
    return None


def detect_local_risk(path: Path, known_folders=None):
    """回傳不適合存放的本機位置名稱（桌面／文件／下載／暫存）；否則 None。"""
    if _is_network_path(path):
        return None                 # 處內共用磁碟不是這台電腦的桌面或暫存區
    folders = known_folders if known_folders is not None else _known_folder_paths()
    for label, folder in folders.items():
        if _under(path, folder):
            return label
    temps = {os.environ.get("TEMP"), os.environ.get("TMP"),
             os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp")}
    if any(t and len(t) > 3 and _under(path, Path(t)) for t in temps):
        return "暫存資料夾"
    parts = [p.lower() for p in Path(path).parts]
    if any(p in ("temp", "tmp") for p in parts[1:2]):      # C:\Temp、D:\tmp
        return "暫存資料夾"
    return None


# ====================== 相依套件檢查 ======================

def check_dependencies():
    """回傳缺少的套件清單。打包版已內含，不會缺。"""
    missing = []
    for mod in ("pandas", "openpyxl", "xlrd"):
        try:
            __import__(mod)
        except ImportError:
            missing.append(mod)
    return missing


# ====================== 專案管理 ======================

BAD_CHARS = r'\/:*?"<>|'


def list_projects():
    PROJECTS_ROOT.mkdir(exist_ok=True)
    return sorted(p.name for p in PROJECTS_ROOT.iterdir()
                  if p.is_dir() and not p.name.startswith("."))


_RESERVED = {"con", "prn", "aux", "nul", *(f"com{i}" for i in range(1, 10)), *(f"lpt{i}" for i in range(1, 10))}


def sanitize(name: str):
    """把不能當資料夾名稱的字元換掉，並回傳 (合法名稱, 是否被改過)。"""
    cleaned = "".join("_" if (c in BAD_CHARS or ord(c) < 32) else c for c in name).strip(" .")
    cleaned = re.sub(r"\s+", " ", cleaned)
    # 超過 60 字會被截斷。舊版把「是否被改過」拿截斷前的字串去比，
    # 結果截斷不會跳確認視窗——兩個只差尾綴的長業務名會靜默撞成同一個
    # 專案，共用同一個 salt 與 input。這裡改成用截斷後的結果比。
    # 截斷後結尾若剛好是空白或句點，Windows 會自動去掉，一定要再修一次。
    truncated = cleaned[:60].rstrip(" .")
    if truncated.split(".")[0].lower() in _RESERVED:
        truncated = "_" + truncated                  # NUL、CON 這類是 Windows 保留名稱
    return truncated, truncated != name.strip()


def use_project(name: str):
    """切換目前專案，並同步更新 anonymize 模組的工作路徑。資料夾建立失敗時丟出 OSError，目前專案不變。"""
    global PROJ_DIR, IN_DIR, OUT_DIR, PRIV_DIR
    proj = PROJECTS_ROOT / name
    (proj / "input").mkdir(parents=True, exist_ok=True)
    PROJ_DIR = proj
    IN_DIR = PROJ_DIR / "input"
    OUT_DIR = PROJ_DIR / "output"
    PRIV_DIR = PROJ_DIR / "_private"
    try:
        LAST_FILE.write_text(name, encoding="utf-8")
    except Exception:
        pass
    return PROJ_DIR


def _input_files():
    """和引擎相同的規則：略過 ~$ 暫存鎖檔與 desktop.ini、Thumbs.db 這類系統檔。"""
    sync_anonymize()
    import anonymize
    return [p for p in sorted(IN_DIR.rglob("*")) if p.is_file() and not p.name.startswith("~$")
            and p.name.lower() not in anonymize.SYSTEM_FILES]


def sync_anonymize():
    """執行前把 anonymize 的路徑指到目前專案。"""
    import anonymize
    anonymize.set_workdir(PROJ_DIR)


# ====================== stdout 轉接 ======================

class QueueWriter:
    """把 anonymize.py 的 print 輸出導到視窗訊息區。"""

    def __init__(self, q):
        self.q = q

    def write(self, s):
        if s:
            self.q.put(s)

    def flush(self):
        pass


# ====================== 主視窗 ======================

class App:

    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.running = False
        self._private_warned = False   # _private 外流警告本次只跳一次
        self._migration_broken = False  # 啟動時搬移舊資料失敗（salt 沒搬過去）→ 不准執行

        root.title("上傳前假名化工具　v2.9　政大研發處")
        root.configure(bg=BG)
        try:                                  # 視窗與工作列圖示（Windows）
            if ICON_FILE.exists():
                root.iconbitmap(default=str(ICON_FILE))
        except Exception:
            pass
        # 依實際螢幕高度決定視窗大小，避免在 1080p 或有工作列縮放的螢幕上被截掉
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        w = min(940, max(820, sw - 120))
        h = min(940, max(640, sh - 100))      # 多了步驟零，預設開高一點
        root.geometry(f"{w}x{h}+{(sw - w) // 2}+16")
        root.minsize(800, 600)

        self._build_ui()
        self.root.after(100, self._drain_queue)
        self.root.after(200, self._startup)

    # ---------- 啟動流程 ----------

    def _startup(self):
        self._migration_broken = False
        self._migrate_legacy()
        projects = list_projects()

        if not projects:
            messagebox.showinfo(
                "先建立一個專案",
                "每一件工作都放在自己的專案資料夾裡，彼此不會互相影響。\n\n"
                "接下來請輸入一個專案名稱，例如：\n"
                "　　國科會新進教師分析\n"
                "　　產學合作經費分析")
            self.new_project()
        else:
            last = None
            if LAST_FILE.exists():
                try:
                    last = LAST_FILE.read_text(encoding="utf-8").strip()
                except Exception:
                    last = None
            match = [p for p in projects if last and p.casefold() == last.casefold()]
            self._set_project(match[0] if match else projects[0])

        if self._cloud:
            self.root.after(300, self._warn_cloud)
        elif self._local_risk:
            self.root.after(300, self._warn_local_risk)

    def _migrate_legacy(self):
        """舊版（v1.x）把 input/output/_private 放在工具根目錄。
        偵測到就整批搬進「專案」底下，避免同仁手動搬錯。"""
        legacy_in = ROOT / "input"
        legacy_priv = ROOT / "_private"
        has_files = legacy_in.exists() and any(legacy_in.iterdir())
        if not (has_files or legacy_priv.exists()):
            return

        if not messagebox.askyesno(
                "偵測到舊版資料",
                "這個資料夾裡有舊版（單一專案）的 input／output／_private。\n\n"
                "新版把每件工作分成獨立專案。\n"
                "要把舊資料整批搬進一個專案資料夾嗎？\n\n"
                "（選「否」則保持原樣不動，但新版不會讀取它們。）"):
            return

        # 舊版會直接 rmtree 掉目標專案的 _private（salt.txt 與 mapping.csv），
        # 完全沒有確認，而且成功訊息還寫「代碼不受影響」。這兩個檔是整個
        # 工具裡唯一無法重建的東西，一旦覆蓋就永久失去跨年對應。
        # 名稱衝突時直接再問一次（這個搬移只在啟動時跑，叫人「再試一次」卻沒有地方可以試）。
        while True:
            name = simpledialog.askstring(
                "搬移到哪個專案", "請給這批舊資料一個專案名稱：",
                initialvalue="原有資料", parent=self.root)
            if not name or not name.strip():
                return
            name, _ = sanitize(name)
            if not name:
                messagebox.showerror("名稱不合法", "請換一個名稱。")
                continue
            target = PROJECTS_ROOT / name
            # 舊的 _private 上次已經搬走、工具資料夾只剩 input／output：output 裡的代碼要靠那份 salt.txt，
            # 只能搬進有 salt.txt 的專案（搬到新專案會和鹽值分家，舊代碼查不回）
            old_out = ROOT / "output"
            if not legacy_priv.exists() and old_out.exists() and any(old_out.iterdir()) \
                    and not (target / "_private" / "salt.txt").exists():
                messagebox.showerror(
                    "這個專案沒有舊資料的 salt.txt",
                    "工具資料夾裡的舊 _private（salt.txt、mapping.csv）已經不在了，可能上次搬移時已經搬進某個專案。\n\n"
                    "舊的 output 裡是代碼，必須和那份 salt.txt 放在同一個專案，否則查不回姓名。\n"
                    f"專案「{name}」沒有 salt.txt，不會搬進去。\n\n"
                    "請輸入上次搬進去的專案名稱；若那個專案的 input／output 已經有新檔，"
                    "請先把新檔移到別處，再重新開啟工具（按「取消」則保持原樣不動）。")
                continue
            # 只有「舊資料還有這個資料夾、目標專案也已經有東西」才算衝突；
            # 上次搬到一半（例如 _private 已搬過去、input 因檔案開著沒搬）可以用同一個名稱續搬
            clash = [sub for sub in ("input", "output", "_private")
                     if (ROOT / sub).exists() and (target / sub).exists() and any((target / sub).iterdir())]
            if not clash:
                break
            messagebox.showerror(
                "這個專案已經有資料了",
                f"專案「{name}」底下已經有：{'、'.join(clash)}\n\n"
                "為了避免覆蓋掉既有的 salt.txt 與 mapping.csv（這兩個檔\n"
                "一旦遺失，今年與去年的代碼就永遠對不起來），不會搬進這個專案。\n\n"
                "按「確定」後請換一個全新的專案名稱。")

        had_salt = (legacy_priv / "salt.txt").exists()
        target.mkdir(parents=True, exist_ok=True)
        moved, failed = [], []
        for sub in ("input", "output", "_private"):
            src = ROOT / sub
            if not src.exists():
                continue
            dst = target / sub
            try:
                # 目標資料夾已存在但是空的（選專案時會自動建立空的 input）：先移除。
                # 用 os.replace 整個資料夾改名：同一個磁碟上是一次完成，裡面有檔案被 Excel 開著時
                # 會整個失敗、原地不動，不會像 shutil.move 那樣複製到一半、兩邊各留一份。
                if dst.exists():
                    dst.rmdir()
                os.replace(src, dst)
                moved.append(sub)
            except OSError:
                failed.append(sub)
        if had_salt and not (target / "_private" / "salt.txt").exists():
            self._migration_broken = True
            messagebox.showerror(
                "搬移後找不到 salt.txt",
                f"專案\\{name}\\_private\\salt.txt 不存在。\n\n"
                "在查明之前請不要執行假名化（會產生一套新的代碼，跟舊資料對不起來），\n"
                "請先關閉開著的 Excel 檔，重新開啟工具再搬一次；仍有問題請聯絡維護人員。")
            return
        if failed:
            messagebox.showwarning(
                "有些資料夾沒有搬過去",
                f"已搬到　專案\\{name}\\：{'、'.join(moved) or '（無）'}\n"
                f"沒有搬動：{'、'.join(failed)}（可能有檔案正被開著）\n\n"
                "請關閉開著的檔案後重新開啟工具，搬移會再詢問一次；\n"
                f"屆時請輸入同一個專案名稱「{name}」。")
            return
        if moved:
            messagebox.showinfo(
                "搬移完成",
                f"已搬到　專案\\{name}\\\n（{'、'.join(moved)}）\n\n"
                + ("salt.txt 一併搬過去了，代碼不受影響。" if had_salt else ""))

    # ---------- 介面 ----------

    def _build_ui(self):
        pad = dict(padx=16, pady=(0, 10))

        head = tk.Frame(self.root, bg=BG)
        head.pack(fill="x", padx=16, pady=(14, 6))

        # 右側：CopyNoright 標章（載入失敗也不影響任何功能）
        badge = tk.Frame(head, bg=BG)
        badge.pack(side="right", padx=(12, 0))
        try:
            self._badge_img = tk.PhotoImage(data=BADGE_B64)
            tk.Label(badge, image=self._badge_img, bg=BG).pack()
            tk.Label(badge, text="自由取用．歡迎轉發", font=("Microsoft JhengHei UI", 8),
                     bg=BG, fg=MUTED).pack(pady=(2, 0))
        except Exception:
            pass

        txt_head = tk.Frame(head, bg=BG)
        txt_head.pack(side="left", fill="x", expand=True)
        tk.Label(txt_head, text="上傳前假名化工具", font=FONT_H,
                 bg=BG, fg=INK).pack(anchor="w")
        tk.Label(txt_head,
                 text="把原始檔中的姓名換成固定代碼，再交給外部 AI 分析。"
                      "對照表只留在這台電腦。",
                 font=FONT, bg=BG, fg=MUTED, justify="left").pack(anchor="w", pady=(2, 0))

        self._cloud = detect_cloud(ROOT)
        self._local_risk = None if self._cloud else detect_local_risk(ROOT)
        if self._cloud:
            tk.Label(txt_head,
                     text=f"⚠　這個資料夾在 {self._cloud} 裡面。"
                          f"姓名對照表會被同步到雲端，請把整個資料夾移到本機再作業。",
                     font=FONT_B, bg="#fee2e2", fg=RED,
                     wraplength=840, justify="left", padx=8, pady=5
                     ).pack(anchor="w", fill="x", pady=(6, 0))
        elif self._local_risk:
            tk.Label(txt_head,
                     text=f"⚠　這個資料夾在{self._local_risk}裡面。"
                          f"這類位置容易被雲端備份接管或被系統清掉，"
                          f"請移到 D:\\假名化作業\\（沒有 D 槽就用 C:\\假名化作業\\）再作業。",
                     font=FONT_B, bg="#fef3c7", fg=AMBER,
                     wraplength=840, justify="left", padx=8, pady=5
                     ).pack(anchor="w", fill="x", pady=(6, 0))

        # ---------- 步驟零 ----------
        f0 = tk.LabelFrame(self.root, text=" 步驟零　選擇專案 ",
                           font=FONT_B, bg=BG, fg=TEAL, bd=1, relief="solid")
        f0.pack(fill="x", **pad)

        b0 = tk.Frame(f0, bg=BG)
        b0.pack(fill="x", padx=12, pady=(10, 4))
        tk.Label(b0, text="目前專案：", font=FONT, bg=BG, fg=INK).pack(side="left")
        self.cmb = ttk.Combobox(b0, font=FONT, state="readonly", width=30)
        self.cmb.pack(side="left", padx=(4, 8))
        self.cmb.bind("<<ComboboxSelected>>",
                      lambda e: self._set_project(self.cmb.get()))
        tk.Button(b0, text="新增專案…", font=FONT, width=12,
                  command=self.new_project).pack(side="left")
        tk.Button(b0, text="開啟專案資料夾", font=FONT, width=14,
                  command=lambda: self._open(PROJ_DIR)).pack(side="left", padx=(8, 0))

        self.lbl_proj = tk.Label(f0, text="", font=("Consolas", 8),
                                 bg=BG, fg=MUTED, anchor="w", justify="left")
        self.lbl_proj.pack(anchor="w", fill="x", padx=12, pady=(0, 4))
        tk.Label(f0,
                 text="每個專案有自己的代碼系統，互不相干。"
                      "需要跨年追蹤的同一件業務，請固定使用同一個專案。",
                 font=FONT, bg=BG, fg=MUTED, wraplength=840, justify="left"
                 ).pack(anchor="w", padx=12, pady=(0, 10))

        # ---------- 步驟一 ----------
        f1 = tk.LabelFrame(self.root, text=" 步驟一　放入要處理的檔案 ",
                           font=FONT_B, bg=BG, fg=TEAL, bd=1, relief="solid")
        f1.pack(fill="x", **pad)

        b1 = tk.Frame(f1, bg=BG)
        b1.pack(fill="x", padx=12, pady=(10, 6))
        tk.Button(b1, text="開啟 input 資料夾", font=FONT, width=18,
                  command=self.open_input).pack(side="left")
        tk.Button(b1, text="選擇檔案加入…", font=FONT, width=16,
                  command=self.add_files).pack(side="left", padx=(8, 0))
        tk.Button(b1, text="重新整理", font=FONT, width=10,
                  command=self.refresh_files).pack(side="left", padx=(8, 0))
        tk.Button(b1, text="清空 input", font=FONT, width=10,
                  command=self.clear_input).pack(side="left", padx=(8, 0))

        self.lbl_drop = tk.Label(
            f1, text="也可以直接把檔案拖曳到這個視窗上",
            font=FONT, bg=BG, fg=MUTED)
        self.lbl_drop.pack(anchor="w", padx=12)

        self.lst = tk.Listbox(f1, height=4, font=FONT_LOG,
                              bd=1, relief="solid", highlightthickness=0)
        self.lst.pack(fill="x", padx=12, pady=(6, 12))

        # ---------- 步驟二 ----------
        f2 = tk.LabelFrame(self.root, text=" 步驟二　執行 ",
                           font=FONT_B, bg=BG, fg=TEAL, bd=1, relief="solid")
        # 先不 pack；等步驟三與底部工具列固定在下方後，再讓它吃掉剩餘空間

        b2 = tk.Frame(f2, bg=BG)
        b2.pack(fill="x", padx=12, pady=(10, 8))
        self.btn_dry = tk.Button(b2, text="① 先檢視（不會改任何檔案）", font=FONT,
                                 width=26, command=lambda: self.run(True))
        self.btn_dry.pack(side="left")
        self.btn_go = tk.Button(b2, text="② 進行假名化", font=FONT_B,
                                width=20, bg="#0f766e", fg="white",
                                activebackground="#115e59", activeforeground="white",
                                command=lambda: self.run(False))
        self.btn_go.pack(side="left", padx=(10, 0))
        self.pb = ttk.Progressbar(b2, mode="indeterminate", length=140)
        self.pb.pack(side="left", padx=(12, 0))

        self.txt = tk.Text(f2, font=FONT_LOG, bd=1, relief="solid",
                           wrap="word", highlightthickness=0, bg="white",
                           height=11, width=1)  # height 必須明確指定
        sb = tk.Scrollbar(f2, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set, state="disabled")
        sb.pack(side="right", fill="y", padx=(0, 12), pady=(0, 12))
        self.txt.pack(fill="both", expand=True, padx=(12, 0), pady=(0, 12))

        self.txt.tag_config("file", foreground=TEAL, font=("Consolas", 10, "bold"))
        self.txt.tag_config("warn", foreground=AMBER)
        self.txt.tag_config("err", foreground=RED, font=("Consolas", 10, "bold"))
        self.txt.tag_config("ok", foreground=INK, font=("Consolas", 10, "bold"))
        self.txt.tag_config("muted", foreground=MUTED)

        # ---------- 步驟三 ----------
        f3 = tk.LabelFrame(self.root, text=" 步驟三　檢視結果 ",
                           font=FONT_B, bg=BG, fg=TEAL, bd=1, relief="solid")

        b3 = tk.Frame(f3, bg=BG)
        b3.pack(fill="x", padx=12, pady=(10, 6))
        self.btn_view = tk.Button(
            b3, text="③ 檢視（同時開啟 output 與 _private）", font=FONT_B,
            width=34, bg="#0f766e", fg="white",
            activebackground="#115e59", activeforeground="white",
            command=self.view_results)
        self.btn_view.pack(side="left")
        tk.Button(b3, text="只開 output", font=FONT, width=12,
                  command=self.open_output).pack(side="left", padx=(10, 0))
        tk.Button(b3, text="只開 _private", font=FONT, width=12,
                  command=self.open_private).pack(side="left", padx=(8, 0))
        self.btn_merge = tk.Button(b3, text="姓名整併…", font=FONT, width=12,
                                   command=self.merge_names)
        self.btn_merge.pack(side="left", padx=(8, 0))

        tk.Label(f3,
                 text="output ＝ 可以上傳的檔案　│　"
                      "_private ＝ 鹽值與姓名對照表，絕對不可外流、不可刪除",
                 font=FONT, bg=BG, fg=RED, wraplength=840, justify="left"
                 ).pack(anchor="w", padx=12, pady=(0, 10))

        # ---------- 底部工具列 ----------
        f4 = tk.Frame(self.root, bg=BG)

        tk.Label(f4, text="代碼還原：", font=FONT, bg=BG, fg=INK).pack(side="left")
        self.ent_code = tk.Entry(f4, font=FONT_LOG, width=14, bd=1, relief="solid")
        self.ent_code.pack(side="left", padx=(4, 6))
        self.ent_code.bind("<Return>", lambda e: self.decode())
        tk.Button(f4, text="查詢姓名", font=FONT, command=self.decode).pack(side="left")
        self.lbl_decode = tk.Label(f4, text="", font=FONT_B, bg=BG, fg=TEAL)
        self.lbl_decode.pack(side="left", padx=(10, 0))

        tk.Button(f4, text="清除本專案代碼（僅示範階段使用）", font=FONT,
                  fg=RED, command=self.reset_all).pack(side="right")

        # ---- 版面固定順序 ----
        # 由下往上 pack：底部工具列 → 步驟三 → 步驟二。
        # 這樣不論訊息區有多少內容、視窗被縮到多小，
        # 步驟三與底部工具列都保證看得見（被壓縮的一定是訊息區）。
        f4.pack(side="bottom", fill="x", padx=16, pady=(0, 12))
        f3.pack(side="bottom", fill="x", padx=16, pady=(0, 10))
        f2.pack(side="top", fill="both", expand=True, padx=16, pady=(0, 10))

        self._enable_dnd()

    def _warn_cloud(self):
        messagebox.showwarning(
            "存放位置需要調整",
            f"這個工具目前放在 {self._cloud} 的同步資料夾裡：\n\n"
            f"{ROOT}\n\n"
            "執行後產生的 _private\\mapping.csv 內含真實姓名，\n"
            f"會被自動同步到 {self._cloud}，等於離開了這台電腦——\n"
            "假名化的防護就失效了。\n\n"
            "請先關閉本工具，把整個資料夾「移動」到本機路徑，例如\n\n"
            "　　D:\\假名化作業\\\n"
            "　（若沒有 D 槽，改用 C:\\假名化作業\\）\n\n"
            "移好之後從新位置重新啟動即可，設定不會遺失。\n\n"
            "※ 用示範檔練習階段可以先不管，因為那是虛構資料；\n"
            "　 但換成真實名單之前，一定要先移好。")

    def _warn_local_risk(self):
        messagebox.showwarning(
            "存放位置需要調整",
            f"這個工具目前放在{self._local_risk}：\n\n"
            f"{ROOT}\n\n"
            "這類位置有兩個風險：\n"
            "　1. Windows 的資料夾備份可能把它同步到雲端，\n"
            "　　 含真實姓名的 _private\\mapping.csv 就離開了這台電腦。\n"
            "　2. 暫存位置可能被系統清掉，salt.txt 一旦遺失，\n"
            "　　 明年的代碼就跟今年對不起來。\n\n"
            "請把整個資料夾「移動」到本機固定位置，例如\n\n"
            "　　D:\\假名化作業\\\n"
            "　（若沒有 D 槽，改用 C:\\假名化作業\\）\n\n"
            "移好之後從新位置重新啟動即可，設定不會遺失。")

    # ---------- 拖放（有裝 tkinterdnd2 才啟用）----------

    def _enable_dnd(self):
        """tkinterdnd2 有裝才啟用；沒裝也完全不影響其他功能。"""
        try:
            self.root.drop_target_register("DND_Files")     # 由 tkinterdnd2 提供
            self.root.dnd_bind("<<Drop>>", self._on_drop)
            self.lbl_drop.config(text="也可以直接把檔案拖曳到這個視窗上　（拖放已啟用）",
                                 fg=TEAL)
        except Exception:
            self.lbl_drop.config(
                text="（此電腦未啟用拖放功能，請用上面的「開啟 input 資料夾」"
                     "或「選擇檔案加入…」）", fg=MUTED)

    def _on_drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        self._copy_in([Path(p) for p in paths])

    # ---------- 專案 ----------

    def _set_project(self, name: str):
        if self.running:
            # 執行中不可換專案，並把下拉選單轉回目前實際在跑的專案
            self._busy()
            if PROJ_DIR is not None:
                self.cmb.set(PROJ_DIR.name)
            return
        try:
            use_project(name)
        except OSError as e:
            messagebox.showerror("無法開啟專案", f"專案資料夾建立失敗：{name}\n\n{e}\n\n請換一個名稱。")
            if PROJ_DIR is not None:
                self.cmb.set(PROJ_DIR.name)
            return
        vals = list_projects()
        self.cmb.config(values=vals)
        self.cmb.set(name)
        self.lbl_proj.config(text=f"資料夾：{PROJ_DIR}")
        self.lbl_decode.config(text="")
        self._clear_log()
        self.log(f"目前專案：{name}\n", "ok")
        self.log(f"　{PROJ_DIR}\n\n", "muted")
        if PRIV_DIR.exists():
            n = self._mapping_count()
            self.log(f"本專案已有代碼 {n} 筆，"
                     "繼續執行會沿用同一套代碼。\n", "muted")
        else:
            self.log("本專案尚未執行過，第一次執行會產生專屬的鹽值與代碼。\n",
                     "muted")
        self.refresh_files()

    def new_project(self):
        if self._busy():
            return
        name = simpledialog.askstring(
            "新增專案",
            "請輸入專案名稱（建議用工作的名字，例如「國科會新進教師分析」）：",
            parent=self.root)
        if not name or not name.strip():
            return
        clean, changed = sanitize(name)
        if not clean:
            messagebox.showerror("名稱不合法", "請換一個名稱。")
            return
        same = [p for p in list_projects() if p.casefold() == clean.casefold()]
        if same:
            clean = same[0]
        exists = (PROJECTS_ROOT / clean).exists()
        if changed and exists:
            # 兩個只差尾綴的長名稱，截斷後會變成同一個專案；
            # 使用者輸入的差異恰好就是被砍掉的那段，一定要先講清楚再切換。
            if not messagebox.askyesno(
                    "名稱調整後與既有專案相同",
                    "資料夾名稱不能包含 \\ / : * ? \" < > | 等字元，且最長 60 個字。\n\n"
                    f"你輸入的名稱調整後變成：\n{clean}\n\n"
                    "這個名稱的專案「已經存在」。\n"
                    "要切換到這個既有專案嗎？\n\n"
                    "（如果這是另一件工作，請按「否」，換一個較短、不會重複的名稱）"):
                return
            self._set_project(clean)
            return
        if exists:
            messagebox.showinfo("已存在", f"專案「{clean}」已經存在，直接切換過去。")
            self._set_project(clean)
            return
        if changed:
            if not messagebox.askyesno(
                    "名稱已調整",
                    "資料夾名稱不能包含 \\ / : * ? \" < > | 等字元，且最長 60 個字。\n\n"
                    f"將建立為：{clean}\n\n可以嗎？"):
                return
        try:
            (PROJECTS_ROOT / clean / "input").mkdir(parents=True, exist_ok=True)
        except OSError as e:
            messagebox.showerror("無法建立專案", f"{clean}\n\n{e}\n\n請換一個名稱。")
            return
        self._set_project(clean)
        self.log(f"\n已建立專案「{clean}」。"
                 "請按「開啟 input 資料夾」放入要處理的檔案。\n", "ok")

    def _mapping_count(self):
        """回傳筆數；讀不到時回傳「無法讀取」字樣（不要顯示成 0 筆，同仁會以為對照表是空的）。"""
        if not (PRIV_DIR / "mapping.csv").exists():
            return 0
        try:
            sync_anonymize()
            import anonymize
            return len(anonymize.load_mapping())
        except Exception:
            return "無法讀取"

    def _known_names(self):
        try:
            sync_anonymize()
            import anonymize
            if not (PRIV_DIR / "mapping.csv").exists():
                return None
            return anonymize.KnownNames({v: k for k, v in anonymize.load_mapping().items()})
        except Exception:
            return None

    # ---------- 檔案操作 ----------

    def _need_project(self):
        if PROJ_DIR is None:
            messagebox.showinfo("請先選擇專案", "請在步驟零選擇或新增一個專案。")
            return False
        return True

    def _busy(self):
        """執行期間擋掉會改變工作路徑的操作。

        anonymize 模組用全域變數保存路徑；背景執行緒還在跑時若切換專案、
        或按「查詢姓名」「姓名整併」，路徑會被改指到另一個專案，
        背景執行緒接著就把這個專案的對照表寫進「別人」的資料夾，
        把那個專案的 mapping.csv 整個蓋掉，而畫面照樣顯示完成。
        """
        if self.running:
            messagebox.showinfo(
                "正在執行中",
                "假名化正在進行，請等畫面顯示完成之後再操作。\n\n"
                "（執行期間切換專案或查詢代碼，可能會影響另一個專案的對照表。）")
            return True
        return False

    def _copy_in(self, paths):
        if self._busy() or not self._need_project():
            return
        IN_DIR.mkdir(parents=True, exist_ok=True)
        n = 0
        for p in paths:
            try:
                if p.is_dir():
                    shutil.copytree(p, IN_DIR / p.name, dirs_exist_ok=True)
                    n += 1
                elif p.is_file():
                    shutil.copy2(p, IN_DIR / p.name)
                    n += 1
            except Exception as e:
                messagebox.showerror("複製失敗", f"{p.name}\n{e}")
        if n:
            self.refresh_files()
            self.log(f"\n已加入 {n} 個項目到「{PROJ_DIR.name}」的 input。\n", "ok")

    def add_files(self):
        if self._busy() or not self._need_project():
            return
        paths = filedialog.askopenfilenames(
            title=f"選擇要加入「{PROJ_DIR.name}」的檔案",
            filetypes=[("試算表", "*.xlsx *.xls *.xlsm"), ("所有檔案", "*.*")])
        if paths:
            self._copy_in([Path(p) for p in paths])

    def clear_input(self):
        if self._busy() or not self._need_project():
            return
        if not messagebox.askyesno(
                "確認", f"要清空「{PROJ_DIR.name}」的 input 資料夾嗎？\n\n"
                        "（只會刪掉待處理的原始檔，不影響 output 與 _private）"):
            return
        for p in IN_DIR.glob("*"):
            shutil.rmtree(p, ignore_errors=True) if p.is_dir() else p.unlink(missing_ok=True)
        self.refresh_files()
        self.log("\ninput 已清空。\n", "muted")

    def refresh_files(self):
        self.lst.delete(0, "end")
        if PROJ_DIR is None:
            self.lst.insert("end", "　（請先在步驟零選擇專案）")
            self.btn_go.config(state="disabled")
            self.btn_dry.config(state="disabled")
            return
        IN_DIR.mkdir(parents=True, exist_ok=True)
        files = _input_files()
        if not files:
            self.lst.insert("end", "　（input 目前是空的——請先放入檔案）")
            self.btn_go.config(state="disabled")
            self.btn_dry.config(state="disabled")
        else:
            known = self._known_names()
            import anonymize                    # 檔名裡的人名遮起來（同仁常截整個視窗求助）
            for p in files:
                self.lst.insert("end", f"　{anonymize.mask_text(str(p.relative_to(IN_DIR)), known)}")
            if not self.running:
                self.btn_go.config(state="normal")
                self.btn_dry.config(state="normal")

    def _open(self, path, create=True):
        if path is None:
            self._need_project()
            return
        path = Path(path)
        if not create and not path.exists():
            return
        path.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.run(["open", str(path)])
            else:
                subprocess.run(["xdg-open", str(path)])
        except Exception as e:
            messagebox.showinfo("資料夾位置", f"{path}\n\n（無法自動開啟：{e}）")

    def open_input(self):
        if self._need_project():
            self._open(IN_DIR)

    def open_output(self):
        if not self._need_project():
            return
        if not OUT_DIR.exists():
            messagebox.showinfo("尚未產生", "這個專案還沒有 output，請先執行步驟二。")
            return
        self._open(OUT_DIR)

    def open_private(self):
        if not self._need_project():
            return
        if not PRIV_DIR.exists():
            messagebox.showinfo("尚未產生", "這個專案還沒有 _private，請先執行步驟二。")
            return
        if self._private_warned or messagebox.askyesno(
                "提醒",
                "_private 內是鹽值與姓名對照表。\n\n"
                "這兩個檔案絕對不可以上傳、不可以寄出、不可以放雲端。\n\n"
                "確定要開啟嗎？"):
            self._private_warned = True
            self._open(PRIV_DIR)

    def view_results(self):
        """同時開啟 output 與 _private 兩個視窗，並在訊息區列出內容清單。"""
        # 執行中讀 mapping.csv 會和引擎存檔撞在一起，導致整批中止
        if self._busy() or not self._need_project():
            return
        if not OUT_DIR.exists() and not PRIV_DIR.exists():
            messagebox.showinfo("尚未產生",
                                "這個專案還沒有結果可以檢視。\n"
                                "請先執行步驟二的「② 進行假名化」。")
            return

        if not self._private_warned:
            if not messagebox.askyesno(
                    "即將開啟兩個資料夾",
                    "　output　　＝ 假名化後的檔案，這些才可以上傳\n"
                    "　_private　＝ 鹽值 salt.txt 與姓名對照表 mapping.csv\n\n"
                    "_private 裡的兩個檔案絕對不可以上傳、寄出或放雲端，\n"
                    "也不可以刪除（刪掉明年的代碼會跟今年對不起來）。\n\n"
                    "確定要開啟嗎？"):
                return
            self._private_warned = True

        self._inventory()

        # 兩個視窗依序開啟；稍微間隔，避免檔案總管把它們併成同一個視窗
        if OUT_DIR.exists():
            self._open(OUT_DIR)
        if PRIV_DIR.exists():
            p = PRIV_DIR
            self.root.after(400, lambda p=p: self._open(p, create=False))

    def _inventory(self):
        """把 output 與 _private 的內容摘要寫進訊息區，方便同仁核對。"""
        self.log("\n" + "─" * 58 + f"\n檢視結果　專案：{PROJ_DIR.name}\n", "ok")

        if OUT_DIR.exists():
            outs = sorted(p for p in OUT_DIR.rglob("*") if p.is_file())
            self.log(f"\n▸ output（可上傳）　共 {len(outs)} 個檔案\n", "file")
            import anonymize                    # 檔名裡看起來像人名的片段遮起來，同仁常截圖求助
            known = self._known_names()
            for p in outs[:20]:
                self.log(f"  · {anonymize.mask_text(str(p.relative_to(OUT_DIR)), known)}\n", "muted")
            if len(outs) > 20:
                self.log(f"  · …其餘 {len(outs) - 20} 個\n", "muted")
        else:
            self.log("\n▸ output 尚未產生（可能只跑過檢視模式）\n", "warn")

        if PRIV_DIR.exists():
            self.log("\n▸ _private（不可外流）\n", "file")
            salt = PRIV_DIR / "salt.txt"
            warn = PRIV_DIR / "name_warnings.csv"
            self.log(f"  · salt.txt　　　{'已存在' if salt.exists() else '缺少'}"
                     "　←　這個檔案決定本專案的代碼，遺失或更換會讓跨年代碼對不起來\n",
                     "muted" if salt.exists() else "err")
            if (PRIV_DIR / "mapping.csv").exists():
                self.log(f"  · mapping.csv　{self._mapping_count()} 筆姓名對照\n",
                         "muted")
            if warn.exists():
                self.log("  · name_warnings.csv　有姓名寫法異常紀錄，建議打開看看\n",
                         "warn")

        self.log("\n即將開啟兩個檔案總管視窗：output 與 _private。\n"
                 "只有 output 裡的檔案可以上傳。\n" + "─" * 58 + "\n", "ok")

    def reset_all(self):
        # 執行中刪掉 _private，背景執行緒會接著重建 mapping.csv，但 salt.txt 已經不見了
        if self._busy() or not self._need_project():
            return
        if not messagebox.askyesno(
                "確認：清除本專案的代碼",
                f"這個動作只影響專案「{PROJ_DIR.name}」，\n"
                "會刪除它的 output 與 _private 兩個資料夾，\n"
                "已產生的代碼與對照表都會消失，下次執行時重新產生。\n\n"
                "【只有在下面這個情況才該按】\n"
                "　這個專案誤跑了測試或示範資料，要從頭來過。\n\n"
                "【處理真實資料的專案，絕對不要按】\n"
                "　會導致今年與去年的代碼對不起來。\n\n"
                "確定要清除嗎？"):
            return
        if not messagebox.askyesno(
                "再確認一次",
                f"真的要刪除「{PROJ_DIR.name}」的 output 與 _private 嗎？"):
            return
        shutil.rmtree(OUT_DIR, ignore_errors=True)
        shutil.rmtree(PRIV_DIR, ignore_errors=True)
        left = [p for p in (OUT_DIR, PRIV_DIR) if p.exists() and any(p.rglob("*"))]
        if left:
            # 有檔案開著時只會清掉一半；不能說「已清除」，否則下次會混著舊對照表與舊 output 執行
            remain = [str(q.relative_to(PROJ_DIR)) for p in left for q in p.rglob("*") if q.is_file()][:10]
            messagebox.showwarning(
                "沒有完全清除",
                "下列檔案刪不掉（可能正被 Excel 開著）：\n\n　" + "\n　".join(remain) +
                "\n\n請關閉這些檔案後再按一次「清除本專案代碼」。")
            self.log("\n沒有完全清除，請關閉開著的檔案後再試一次。\n", "err")
            return
        self.log(f"\n已清除「{PROJ_DIR.name}」的 output 與 _private，"
                 "下次執行會重新產生鹽值與代碼。\n", "warn")

    def decode(self):
        if self._busy() or not self._need_project():
            return
        code = self.ent_code.get().strip().upper()
        if not code:
            return
        try:
            sync_anonymize()
            import anonymize
            name = anonymize.load_mapping().get(code)
            if name is None and anonymize.load_retired().get(code):
                name = f"{anonymize.load_retired()[code]}（這個代碼已整併）"
        except Exception as e:
            messagebox.showerror("錯誤", str(e))
            return
        self.lbl_decode.config(
            text=(name or "（本專案查無此代碼）"),
            fg=TEAL if name else RED)


    # ---------- 姓名整併 ----------

    def _read_warning_groups(self):
        """從 name_warnings.csv 讀出疑似同一人的群組。"""
        f = PRIV_DIR / "name_warnings.csv"
        if not f.exists():
            return {}
        groups = {}
        try:
            with f.open(encoding="utf-8-sig", newline="") as fh:
                for r in csv.DictReader(fh):
                    key = (r.get("正規化姓名") or "").strip()
                    name = (r.get("原始寫法") or "").strip()
                    code = (r.get("代碼") or "").strip()
                    if key and name:
                        groups.setdefault(key, []).append((name, code))
        except Exception as e:
            messagebox.showerror("讀取失敗", f"name_warnings.csv\n{e}")
            return {}
        return groups

    def merge_names(self):
        """開啟姓名整併視窗：告訴工具哪些不同寫法其實是同一個人。"""
        if self._busy() or not self._need_project():
            return
        sync_anonymize()
        import anonymize

        groups = self._read_warning_groups()
        try:
            rules = anonymize.load_aliases_raw()    # 原本寫的規則（不攤平串接，循環規則也列出來才能刪）
        except Exception as e:
            messagebox.showerror("姓名整併表讀不到", f"{e}\n\n檔案位置：{PRIV_DIR / 'aliases.csv'}")
            return

        win = tk.Toplevel(self.root)
        win.title(f"姓名整併　－　{PROJ_DIR.name}")
        win.configure(bg=BG)
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        w, h = min(820, sw - 120), min(760, sh - 120)
        win.geometry(f"{w}x{h}+{(sw - w) // 2}+24")
        win.transient(self.root)
        win.grab_set()

        tk.Label(win, text="姓名整併", font=FONT_H, bg=BG, fg=INK
                 ).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(win,
                 text="同一個人被打成兩種寫法（多空格、錯字）時，會拿到兩個代碼。"
                      "在這裡指定哪些是同一人，整併後他們會共用同一個代碼。",
                 font=FONT, bg=BG, fg=MUTED, wraplength=w - 60, justify="left"
                 ).pack(anchor="w", padx=16)

        # ---- 底部按鈕先固定（避免內容太長時被擠出畫面）----
        bar = tk.Frame(win, bg=BG)
        bar.pack(side="bottom", fill="x", padx=16, pady=(6, 14))
        tk.Label(win,
                 text="⚠　整併會改變代碼。存檔後請重新執行，"
                      "並用新的 output 取代先前已上傳的檔案。",
                 font=FONT, bg=BG, fg=RED, wraplength=w - 60, justify="left"
                 ).pack(side="bottom", anchor="w", padx=16, pady=(0, 4))

        # ---- 可捲動的內容區 ----
        body_wrap = tk.Frame(win, bg=BG)
        body_wrap.pack(fill="both", expand=True, padx=16, pady=(10, 0))
        canvas = tk.Canvas(body_wrap, bg=BG, highlightthickness=0, bd=0)
        vs = tk.Scrollbar(body_wrap, orient="vertical", command=canvas.yview)
        body = tk.Frame(canvas, bg=BG)
        body.bind("<Configure>",
                  lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win_id = canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win_id, width=e.width))
        canvas.configure(yscrollcommand=vs.set)
        canvas.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        # bind_all 是「全域」綁定：視窗關掉之後若沒解綁，之後在主視窗每滾一次
        # 滑鼠滾輪就丟一個 TclError，執行期間還會把 traceback 灌進訊息區。
        def _on_wheel(e):
            try:
                canvas.yview_scroll(int(-e.delta / 120), "units")
            except tk.TclError:
                pass

        canvas.bind_all("<MouseWheel>", _on_wheel)

        def _release_wheel():
            try:
                win.unbind_all("<MouseWheel>")
            except Exception:
                pass

        win.bind("<Destroy>", lambda e: _release_wheel() if e.widget is win else None)

        # ---- 區塊一：自動偵測到的疑似同一人 ----
        picks = {}
        fa = tk.LabelFrame(body, text=" 自動偵測：疑似同一人 ",
                           font=FONT_B, bg=BG, fg=TEAL, bd=1, relief="solid")
        fa.pack(fill="x", pady=(0, 12))

        if not groups:
            tk.Label(fa, text="目前沒有偵測到疑似同一人的寫法差異。",
                     font=FONT, bg=BG, fg=MUTED).pack(anchor="w", padx=12, pady=10)
        else:
            tk.Label(fa,
                     text="請在每一組中，點選「要保留的正確寫法」；"
                          "其餘寫法會併到它身上。\n"
                          "提示：中文姓名通常不含空格；英文姓名通常「要」空格。"
                          "不確定的就選「先不處理」。",
                     font=FONT, bg=BG, fg=MUTED, justify="left",
                     wraplength=w - 90).pack(anchor="w", padx=12, pady=(8, 4))
            for key, variants in groups.items():
                g = tk.Frame(fa, bg=BG, bd=1, relief="groove")
                g.pack(fill="x", padx=12, pady=5)
                var = tk.StringVar(value="")      # 預設「先不處理」
                picks[key] = (var, variants)
                for name, code in variants:
                    tk.Radiobutton(
                        g, text=f"保留「{name}」　（目前代碼 {code}）",
                        variable=var, value=name, font=FONT, bg=BG,
                        anchor="w", justify="left", selectcolor="white"
                    ).pack(anchor="w", padx=8, pady=1)
                tk.Radiobutton(g, text="先不處理（不是同一人，或還要再確認）",
                               variable=var, value="", font=FONT, bg=BG,
                               fg=MUTED, anchor="w", selectcolor="white"
                               ).pack(anchor="w", padx=8, pady=(1, 4))

        # ---- 區塊二：手動整併 ----
        fm = tk.LabelFrame(body, text=" 手動整併（錯字等自動偵測不到的情況） ",
                           font=FONT_B, bg=BG, fg=TEAL, bd=1, relief="solid")
        fm.pack(fill="x", pady=(0, 12))
        tk.Label(fm,
                 text="例如「陳大文」被打成「陳大聞」，兩個寫法都要照原樣輸入。",
                 font=FONT, bg=BG, fg=MUTED).pack(anchor="w", padx=12, pady=(8, 4))
        mrow = tk.Frame(fm, bg=BG)
        mrow.pack(fill="x", padx=12, pady=(0, 10))
        tk.Label(mrow, text="把", font=FONT, bg=BG).pack(side="left")
        e_from = tk.Entry(mrow, font=FONT, width=20, bd=1, relief="solid")
        e_from.pack(side="left", padx=4)
        tk.Label(mrow, text="併入", font=FONT, bg=BG).pack(side="left")
        e_to = tk.Entry(mrow, font=FONT, width=20, bd=1, relief="solid")
        e_to.pack(side="left", padx=4)

        # ---- 區塊三：目前規則 ----
        fr = tk.LabelFrame(body, text=" 目前的整併規則 ",
                           font=FONT_B, bg=BG, fg=TEAL, bd=1, relief="solid")
        fr.pack(fill="x", pady=(0, 12))
        lst = tk.Listbox(fr, height=6, font=FONT_LOG, bd=1, relief="solid",
                         highlightthickness=0)
        lst.pack(fill="x", padx=12, pady=(10, 4))

        def refresh_rules():
            lst.delete(0, "end")
            if not rules:
                lst.insert("end", "　（尚無規則）")
            for a, b in sorted(rules.items(), key=lambda x: (x[1], x[0])):
                lst.insert("end", f"　「{a}」　→　「{b}」")

        def add_manual():
            a, b = e_from.get().strip(), e_to.get().strip()
            if not a or not b:
                messagebox.showinfo("請填寫", "兩邊都要填。", parent=win)
                return
            if a == b:
                messagebox.showinfo("不需整併", "兩個寫法一樣。", parent=win)
                return
            rules[a] = b
            e_from.delete(0, "end"), e_to.delete(0, "end")
            refresh_rules()

        def del_rule():
            sel = lst.curselection()
            if not sel:
                return
            line = lst.get(sel[0])
            if "→" not in line:
                return
            a = line.split("「")[1].split("」")[0]
            rules.pop(a, None)
            refresh_rules()

        tk.Button(mrow, text="加入", font=FONT, width=8,
                  command=add_manual).pack(side="left", padx=(8, 0))
        tk.Button(fr, text="刪除選取的規則", font=FONT,
                  command=del_rule).pack(anchor="w", padx=12, pady=(0, 10))
        refresh_rules()

        # ---- 存檔 ----
        def collect():
            merged = dict(rules)
            for key, (var, variants) in picks.items():
                keep = var.get()
                if not keep:
                    continue
                for name, _code in variants:
                    if name != keep:
                        merged[name] = keep
            return merged

        def save(rerun: bool):
            merged = collect()
            try:
                sync_anonymize()
                import anonymize
                anonymize.save_aliases(merged)
            except Exception as e:
                messagebox.showerror("儲存失敗", str(e), parent=win)
                return
            win.grab_release()
            win.destroy()
            self.log(f"\n已儲存 {len(merged)} 條姓名整併規則"
                     f"（{PRIV_DIR / 'aliases.csv'}）。\n", "ok")
            if rerun:
                self.run(False)
            else:
                self.log("下次執行時才會生效。\n", "muted")

        tk.Button(bar, text="儲存並立即重新執行", font=FONT_B, width=20,
                  bg="#0f766e", fg="white", activebackground="#115e59",
                  activeforeground="white",
                  command=lambda: save(True)).pack(side="left")
        tk.Button(bar, text="只儲存", font=FONT, width=10,
                  command=lambda: save(False)).pack(side="left", padx=(10, 0))
        tk.Button(bar, text="取消", font=FONT, width=10,
                  command=lambda: (win.grab_release(), win.destroy())
                  ).pack(side="right")

    # ---------- 執行 ----------

    def run(self, dry: bool):
        if self.running or not self._need_project():
            return
        files = _input_files()
        if not files:
            messagebox.showinfo("input 是空的", "請先把要處理的檔案放進 input 資料夾。")
            return
        if self._migration_broken:
            messagebox.showerror(
                "舊版資料還沒搬完",
                "啟動時搬移舊版資料失敗（salt.txt 沒有搬過去）。\n\n"
                "現在執行會產生一套新的代碼，跟舊資料對不起來。\n"
                "請關閉開著的 Excel 檔，重新開啟工具再搬一次。")
            return
        if not dry and not messagebox.askyesno(
                "確認執行",
                f"專案：{PROJ_DIR.name}\n\n"
                f"將處理 input 內的 {len(files)} 個檔案，\n"
                "結果寫入這個專案的 output 資料夾（原始檔不會被改動）。\n\n"
                "確定開始嗎？"):
            return

        self.running = True
        self.btn_go.config(state="disabled")
        self.btn_dry.config(state="disabled")
        self.pb.start(12)
        self._clear_log()
        self.log(f"專案：{PROJ_DIR.name}\n", "ok")
        self.log("【檢視模式】不會寫出任何檔案\n\n" if dry else "【開始執行】\n\n", "ok")

        threading.Thread(target=self._worker, args=(dry,), daemon=True).start()

    def _worker(self, dry: bool):
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = sys.stderr = QueueWriter(self.q)
        result, crashed = None, False
        try:
            sync_anonymize()          # 把工作路徑指到目前專案
            import anonymize
            result = anonymize.main(["--dry-run"] if dry else [])
        except SystemExit as e:
            crashed = bool(e.code)
            if e.code:
                print(f"\n!! [中止] {e}")
        except Exception:
            crashed = True
            print("\n" + "=" * 60)
            print("!! 發生錯誤，處理並未完成，請勿上傳 output 裡的檔案。")
            print("請把以下內容整段複製給資訊窗口或原分析人員：")
            try:
                import anonymize                # 錯誤訊息裡偶爾會帶到儲存格內容，遮掉像人名的片段
                print(anonymize.mask_text(traceback.format_exc()))
            except Exception:
                print(traceback.format_exc())
        finally:
            sys.stdout, sys.stderr = old_out, old_err
            self.q.put(("__DONE__", dry, result, crashed))

    def _drain_queue(self):
        try:
            while True:
                item = self.q.get_nowait()
                if isinstance(item, tuple) and item[0] == "__DONE__":
                    self._finish(*item[1:])
                else:
                    self.log(item)
        except queue.Empty:
            pass
        self.root.after(100, self._drain_queue)

    @staticmethod
    def _problems(dry: bool, result, crashed: bool):
        """把引擎回傳的結果整理成要讓同仁確認的事項。空清單＝可以說「完成」。"""
        problems = []
        if crashed:
            problems.append("執行過程發生錯誤，" + ("檢視沒有完成" if dry else "處理並不完整"))
        if not isinstance(result, dict):
            return problems
        r = result.get
        if r("stop_reason"):
            problems.append(f"處理沒有開始：{str(r('stop_reason')).rstrip('。')}")
        if r("mapping_locked"):
            problems.append("對照表 mapping.csv 無法存檔（可能正被 Excel 開著），處理已中止")
        if r("not_processed"):
            problems.append(f"{r('not_processed')} 個檔案沒有處理")
        if r("mapping_missing"):
            problems.append("mapping.csv 不見了（有鹽值、output 也有檔案），請先從備份還原")
        if r("failed"):
            problems.append(f"{r('failed')} 個檔案處理失敗")
        if r("mapping_rewritten"):
            problems.append("mapping.csv 有多餘欄位或不是 UTF-8，存檔會改寫格式（原檔已先備份，見訊息區）")
        if r("skipped"):
            problems.append(f"{r('skipped')} 個檔案工具無法處理（沒有放進 output，原檔請勿上傳）")
        if r("no_name_files"):
            problems.append(f"{r('no_name_files')} 個檔案（或其中的工作表）沒有偵測到姓名欄位")
        if r("zero_replaced_files"):
            problems.append(f"{r('zero_replaced_files')} 個檔案偵測到姓名欄卻沒有換成代碼")
        if r("name_hits_elsewhere"):
            problems.append(f"姓名欄以外找到 {r('name_hits_elsewhere')} 處已知姓名"
                            + ("（執行時會換成代碼）" if dry else "（已換成代碼，附近可能還有漏網的名字）"))
        if r("short_name_files"):
            problems.append(f"{r('short_name_files')} 個檔案的句子裡有兩個字的已知姓名或疑似姓名沒有自動替換")
        if r("structure_warn_files"):
            problems.append(f"{r('structure_warn_files')} 個檔案的表格結構特殊（直式表單、欄名清單或疑似第二段表頭）")
        if r("person_names_in_titles"):
            problems.append(f"{r('person_names_in_titles')} 個檔案的工作表名稱或檔名看起來含人名"
                            + ("（執行時會改名）" if dry else "（output 已改名）"))
        if r("already_coded_files"):
            problems.append(f"{r('already_coded_files')} 個檔案的姓名欄已經是代碼（可能放錯檔案）")
        if r("suspect_files"):
            problems.append(f"合計 {r('suspect_files')} 個檔案有需要人工確認的地方（見訊息區 ! 開頭的說明）")
        if r("config_problems"):
            problems.append(f"欄位設定.txt 或姓名整併表有 {r('config_problems')} 項沒有生效")
        if r("quarantine_failed"):
            problems.append(f"output 裡有 {r('quarantine_failed')} 個應該移出的舊檔搬不走，仍在 output 裡")
        if r("quarantined") or r("stale_quarantined"):
            problems.append(f"已把 output 裡 {(r('quarantined') or 0) + (r('stale_quarantined') or 0)} 個舊檔移到 "
                            "_private\\output_隔離（若先前上傳過 output，請確認這些檔案的內容）")
        if r("quarantine_pending"):
            problems.append(f"output 裡有 {r('quarantine_pending')} 個舊檔，執行時會先移出")
        if r("legacy_unverified"):
            problems.append(f"output 裡有 {r('legacy_unverified')} 個沒有來源紀錄的舊檔，無法確認內容")
        if r("sources_problem"):
            problems.append("output_sources.csv 讀寫失敗或不存在，分批處理時的同名檔案可能互相覆蓋")
        if r("collisions"):
            problems.append(f"{r('collisions')} 個姓名的代碼和別人（或被整併掉的舊代碼）相同，已改配；"
                            "若是取消了姓名整併，請確認同一人前後代碼是否一致")
        if dry and r("name_dups"):
            problems.append(f"偵測到 {r('name_dups')} 組疑似同一人的不同寫法（執行後可以用「姓名整併」處理）")
        if r("warnings_locked"):
            problems.append("name_warnings.csv 正被開啟，這次沒有更新")
        if r("earlier_hits"):
            problems.append(f"先前批次的 {r('earlier_hits')} 個 output 含這次才認得的姓名，請把原始檔放回 input 重跑")
        return problems

    def _finish(self, dry: bool, result=None, crashed: bool = False):
        self.running = False
        self.pb.stop()
        self.refresh_files()
        # 舊版不論結果如何都印「完成」並響提示音；檢視模式則完全不提示。
        # 檢視是同仁「正式跑之前先確認」的那一步，出錯時一樣要攔下來。
        problems = self._problems(dry, result, crashed)
        if dry:
            if problems:
                self.log("\n檢視發現需要確認的狀況：" + "；".join(problems) + "。\n", "err")
                self.root.bell()
                messagebox.showwarning(
                    "檢視發現需要確認的狀況",
                    "檢視時發現：\n\n　・" + "\n　・".join(problems)
                    + "\n\n請看訊息區的說明，確認後再按「② 進行假名化」。")
            elif result is None:
                self.log("\n沒有檢視任何檔案（input 沒有可處理的檔案）。\n", "warn")
            else:
                self.log("\n檢視完成，沒有發現需要確認的狀況。\n", "ok")
            return
        if problems:
            self.log("\n未完全成功：" + "；".join(problems) + "。\n", "err")
            self.log("請先看上面的訊息確認清楚，再決定要不要上傳 output。\n", "err")
            self.root.bell()
            messagebox.showwarning(
                "處理未完全成功",
                "這次執行有狀況需要你確認：\n\n　・"
                + "\n　・".join(problems)
                + "\n\n請看訊息區的說明。在確認之前，請不要上傳 output 裡的檔案。")
        elif result is None:
            self.log("\n沒有處理任何檔案（input 沒有可處理的檔案）。\n", "warn")
        else:
            self.log("\n完成。請按下方步驟三的「③ 檢視」核對結果。\n", "ok")
            self.root.bell()
        if PRIV_DIR and (PRIV_DIR / "name_warnings.csv").exists():
            if messagebox.askyesno(
                    "偵測到疑似同一人",
                    "有幾組姓名寫法不同、但很可能是同一個人\n"
                    "（例如多了空格）。\n\n"
                    "這會讓同一個人拿到兩個代碼，跨檔比對就對不起來。\n\n"
                    "要現在開啟「姓名整併」，指定哪些是同一人嗎？"):
                self.merge_names()

    # ---------- 訊息區 ----------

    def _clear_log(self):
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.config(state="disabled")

    def log(self, s: str, tag=None):
        if tag is None:
            t = s.strip()
            if t.startswith("▸"):
                tag = "file"
            # 引擎的訊息有固定前綴：「!!」＝失敗、這些檔案不能用（紅）；「!」＝需要確認（橘）。
            # 不用中文字眼猜：專案名稱或欄名含「失敗」「中止」時，舊寫法會把一般訊息染紅。
            elif t.startswith("!!") and not t.startswith("!!!") or t.startswith("Traceback"):
                tag = "err"
            elif t.startswith("!") or t.startswith("注意：") or t.startswith("・"):
                tag = "warn"
            elif t.startswith("·") or t.startswith("="):
                tag = "muted"
            elif t.startswith(("處理檔案", "替換儲存格", "本次新增姓名")):
                tag = "ok"
        # Tk 的 Text 排版「單一超長行」的成本會暴增（實測 16,000 字一行要 30 秒，
        # 整個視窗凍住）；同樣字數拆成多行只要零點零幾秒。
        if len(s) > 400:
            s = "\n".join(
                "\n".join(line[k:k + 200] for k in range(0, len(line), 200)) if len(line) > 400 else line
                for line in s.split("\n"))
        self.txt.config(state="normal")
        self.txt.insert("end", s, tag or ())
        self.txt.see("end")
        self.txt.config(state="disabled")


# ====================== 進入點 ======================

def main():
    missing = check_dependencies()
    if missing:
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror(
            "尚未安裝必要套件",
            "缺少：" + "、".join(missing) + "\n\n"
            "請先關閉本視窗，回到資料夾雙擊執行\n"
            "「1_第一次執行_安裝環境.bat」\n"
            "安裝完成後再重新啟動本工具。")
        return

    # 未打包時，確保能 import 同資料夾的 anonymize.py
    # （打包成 exe 時不加，否則 exe 旁邊若留有舊版 anonymize.py 會被誤用）
    if not getattr(sys, "frozen", False):
        sys.path.insert(0, str(ROOT))

    # 有裝 tkinterdnd2 就用它建立視窗（才支援把檔案拖進來）；沒裝則退回標準視窗
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except Exception:
        root = tk.Tk()

    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
