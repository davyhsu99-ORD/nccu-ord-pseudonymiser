#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
====================================================================
 政大研發處　上傳前假名化工具　解除安裝  uninstall.py  v3.0
====================================================================

設計原則（要改這支程式之前請先看完）

1. 預設什麼都不刪。直接按 Enter 一律是「取消」。
2. 要刪資料夾，一定要先備份，而且備份要**逐檔比對 SHA-256**；
   驗證不過就什麼都不刪。
3. `_private\\mapping.csv` 是「代碼 → 真實姓名」的唯一一份對照表。
   刪掉之後，已經發出去的代碼永遠還原不回姓名，明年的年度作業也接不上。
   所以任何會碰到它的動作都要多繞好幾道。
4. 備份範圍一律「整個工具資料夾」（只排除 __pycache__），
   確保**備份範圍必定涵蓋刪除範圍**——不去猜哪些檔案重要。
   舊版（v1.x）會把 input／output／_private 直接放在工具根目錄
   （見 anonymize_gui.py 的 _migrate_legacy：使用者當初若選「否」就會留在那裡），
   整包備份可以連這種情況一起涵蓋，不必特別判斷。
5. Python 本體與 pandas／openpyxl／xlrd 一律不自動移除——
   這台電腦的其他程式很可能也在用。只顯示指令，由人自己決定。
6. 全部邏輯寫在 Python 而不是批次檔：路徑與輸入不會被當成指令執行，
   雜湊比對用 hashlib（不依賴外部指令是否存在）。
"""

import ctypes
import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path

# ── 這個工具一定會有的檔案。缺任何一個就不是工具資料夾，直接拒絕執行 ──
SIGNATURE = ("anonymize.py", "anonymize_gui.py", "2_啟動假名化工具.bat")

# ── 屬於「程式」的檔案：選項 3 只刪這些，使用者資料一概不動 ──
PROGRAM_FILES = (
    "anonymize.py", "anonymize_gui.py", "requirements.txt", "使用說明.txt",
    "假名化工具.ico", "欄位設定_範例.txt", "field_settings_example_EN.txt", "LICENSE.txt",
    "1_第一次執行_安裝環境.bat", "2_啟動假名化工具.bat", "9_解除安裝.bat",
    "uninstall.py", "last_project.txt",
)

PACKAGES = ("pandas", "openpyxl", "xlrd", "tkinterdnd2")

CLOUD_WORDS = ("onedrive", "dropbox", "google drive", "googledrive", "雲端硬碟",
               "icloud", "sharepoint", "creative cloud", "nextcloud", "seafile",
               "堅果雲", "mega", "pcloud", "box sync")

SEP = "=" * 66
ROOT = Path(__file__).resolve().parent


# ══════════════════════════════════════════════════════════════
#  畫面
# ══════════════════════════════════════════════════════════════

def say(msg=""):
    print(msg, flush=True)


def title(text):
    say()
    say(SEP)
    say("　" + text)
    say(SEP)
    say()


def ask(prompt):
    """讀一行。使用者直接按 Enter 或關視窗，一律回空字串。"""
    try:
        return input(prompt).strip()
    except (EOFError, KeyboardInterrupt):
        say()
        return ""


def confirm(prompt):
    """只有明確輸入 y／Y 才算同意；Enter＝否。"""
    return ask(prompt + "（輸入 y 表示是，直接按 Enter 表示否）：").lower() == "y"


def pause():
    ask("按 Enter 回到選單…")


# ══════════════════════════════════════════════════════════════
#  安全鎖
# ══════════════════════════════════════════════════════════════

def refuse_if_not_tool_folder():
    missing = [n for n in SIGNATURE if not (ROOT / n).exists()]
    if missing:
        title("停下來了，沒有做任何事")
        say("這個資料夾看起來不是「假名化工具」的資料夾，因為找不到：")
        for n in missing:
            say("　　・" + n)
        say()
        say("目前位置：" + str(ROOT))
        say()
        say("請把這支程式放回工具資料夾再執行。")
        say("（這道檢查是為了避免它被複製到別處後誤刪其他資料。）")
        say()
        pause()
        sys.exit(1)


def refuse_if_dangerous_location():
    """不准在桌面、文件、家目錄、系統資料夾或磁碟根目錄執行。"""
    p = ROOT
    if p.parent == p:                      # 磁碟根目錄
        bad = "磁碟根目錄"
    else:
        names = {}
        for env in ("USERPROFILE", "ProgramFiles", "ProgramFiles(x86)", "windir", "PUBLIC"):
            v = os.environ.get(env)
            if v:
                names[Path(v).resolve()] = env
        home = Path.home().resolve()
        for extra in ("Desktop", "Documents", "桌面", "文件"):
            names[home / extra] = extra
        # 開了 OneDrive 資料夾備份時，真正的桌面／文件在 OneDrive 底下，以登錄檔記錄的位置為準
        names.setdefault(desktop_dir().resolve(), "桌面")
        names.setdefault(documents_dir().resolve(), "文件")
        names[home] = "使用者家目錄"
        bad = names.get(p)
    if bad:
        title("停下來了，沒有做任何事")
        say(f"這支程式不能在「{bad}」直接執行：{p}")
        say("工具資料夾應該是一個獨立的資料夾，例如 D:\\假名化作業\\假名化工具\\。")
        say()
        pause()
        sys.exit(1)


# ══════════════════════════════════════════════════════════════
#  盤點
# ══════════════════════════════════════════════════════════════

def iter_files(base: Path):
    """工具資料夾底下所有檔案，排除 __pycache__。"""
    for p in base.rglob("*"):
        if p.is_file() and "__pycache__" not in p.parts:
            yield p


def count_names(csv_path: Path):
    """數 mapping.csv 有幾筆姓名（扣掉標題列）。讀不到就回 None。"""
    try:
        with csv_path.open("r", encoding="utf-8", errors="replace") as fh:
            n = sum(1 for _ in fh)
        return max(0, n - 1)
    except OSError:
        return None


def scan():
    """回傳現況。mapping.csv 用 rglob 找，巢狀專案結構也找得到。"""
    info = {"files": 0, "bytes": 0, "maps": [], "names": 0, "legacy": False,
            "projects": set(), "unknown_root": []}
    for p in iter_files(ROOT):
        info["files"] += 1
        try:
            info["bytes"] += p.stat().st_size
        except OSError:
            pass
        if p.name == "mapping.csv":
            n = count_names(p)
            info["maps"].append((p, n))
            if n:
                info["names"] += n
    # 專案（不限第一層）
    proj_root = ROOT / "專案"
    if proj_root.is_dir():
        for p in proj_root.rglob("_private"):
            if p.is_dir():
                info["projects"].add(p.parent)
    # 舊版資料直接放在工具根目錄
    if (ROOT / "_private").is_dir() or (ROOT / "input").is_dir() or (ROOT / "output").is_dir():
        info["legacy"] = True
    # 根目錄下不屬於產品的檔案（使用者自己放的）
    for p in ROOT.iterdir():
        if p.is_file() and p.name not in PROGRAM_FILES and p.name != "欄位設定.txt":
            info["unknown_root"].append(p.name)
    return info


def human(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:,.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024


def show_status(info):
    say("　工具資料夾：")
    say("　　" + str(ROOT))
    say()
    say("　目前狀況：")
    say(f"　　檔案總數　：{info['files']:,} 個（{human(info['bytes'])}）")
    say(f"　　專案　　　：{len(info['projects'])} 個")
    if info["maps"]:
        total = sum(n for _, n in info["maps"] if n)
        say(f"　　對照表　　：{len(info['maps'])} 份，合計約 {total:,} 筆姓名")
    else:
        say("　　對照表　　：沒有找到 mapping.csv（可能還沒處理過真實資料）")
    if info["legacy"]:
        say("　　※ 偵測到舊版（v1.x）直接放在工具根目錄的 input／output／_private")
    if info["unknown_root"]:
        say(f"　　※ 根目錄有 {len(info['unknown_root'])} 個不屬於本工具的檔案，"
            "備份時會一併複製")
    say()


def warn_mapping(info):
    if not info["maps"]:
        return
    total = sum(n for _, n in info["maps"] if n)
    say("　" + "─" * 62)
    say("　⚠　這台電腦上有 %d 份姓名對照表，合計約 %s 筆姓名。" % (len(info["maps"]), f"{total:,}"))
    say("　　　mapping.csv 是「代碼 → 真實姓名」的唯一一份紀錄。")
    say("　　　刪掉之後，已經發出去的代碼永遠還原不回姓名，")
    say("　　　明年的年度作業也會接不上（代碼會全部重新洗牌）。")
    say("　" + "─" * 62)
    say()


# ══════════════════════════════════════════════════════════════
#  備份與驗證
# ══════════════════════════════════════════════════════════════

def sha256(path: Path):
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def is_inside(child: Path, parent: Path):
    try:
        child.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


def choose_backup_dir():
    """問備份位置。回傳 Path 或 None（放棄）。"""
    # 預設放在工具資料夾旁邊（第零步要求的本機硬碟）。家目錄\Desktop 在 OneDrive 資料夾備份下
    # 不是真正的桌面（看不到、使用者找不回），真正的桌面又會同步上雲，都不適合放含真實姓名的備份
    default = ROOT.parent / ("假名化工具_備份_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    say("　請問要把備份放在哪裡？")
    say("　　直接按 Enter＝用預設位置：")
    say("　　" + str(default))
    say("　　（也可以貼上一個資料夾路徑，例如 D:\\備份）")
    say()
    raw = ask("　備份位置：")
    if not raw:
        target = default
    else:
        raw = raw.strip().strip('"').strip("'")
        try:
            base = Path(raw).expanduser()
        except (OSError, ValueError):
            say("　這個路徑看不懂，取消。")
            return None
        if not base.is_absolute():
            say("　請輸入完整路徑（例如 D:\\備份），取消。")
            return None
        if not base.exists():
            say("　找不到這個資料夾：" + str(base))
            return None
        target = base / ("假名化工具_備份_" + datetime.now().strftime("%Y%m%d_%H%M%S"))

    if is_inside(target, ROOT):
        say("　不行：備份不能放在工具資料夾裡面，否則刪除時會跟著一起消失。")
        return None
    low = str(target).lower()
    if any(w in low for w in CLOUD_WORDS):
        say()
        say("　⚠　這個位置看起來會自動同步到雲端。")
        say("　　　備份裡有完整的真實姓名對照表，同步上去等於姓名離開這台電腦。")
        if not confirm("　　　仍然要用這個位置嗎？"):
            return None
    if target.exists():
        say("　這個備份資料夾已經存在，為了不覆蓋舊備份，取消。")
        return None
    return target


def do_backup(target: Path):
    """整包複製並逐檔 SHA-256 驗證。回傳 True/False。"""
    src_files = sorted(iter_files(ROOT))
    say()
    say(f"　開始備份 {len(src_files):,} 個檔案 → {target}")
    say("　（大的專案可能要等一下，請不要關視窗）")
    try:
        target.mkdir(parents=True, exist_ok=False)
    except OSError as e:
        say("　★ 建立備份資料夾失敗：" + str(e))
        return False

    copied = []
    for i, p in enumerate(src_files, 1):
        rel = p.relative_to(ROOT)
        dst = target / rel
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(p, dst)
            copied.append(rel)
        except OSError as e:
            say(f"　★ 複製失敗：{rel}")
            say(f"　　{e}")
            say("　　（若檔案正被 Excel 開著，請關掉後重新執行）")
            return False
        if i % 200 == 0:
            say(f"　　已複製 {i:,}/{len(src_files):,}")

    say(f"　複製完成 {len(copied):,} 個檔案，開始逐檔比對 SHA-256……")
    ok, bad = verify_backup(target, quiet=False)
    if not ok:
        say("　★ 備份驗證沒有通過，這份備份不可信任。")
        for b in bad[:10]:
            say("　　" + b)
        return False

    # 附一張說明，讓一年後撿到這個資料夾的人知道它是什麼
    try:
        (target / "這份備份是什麼_請先讀我.txt").write_text(
            "\r\n".join([
                "這是「上傳前假名化工具」的完整備份。",
                "",
                "備份時間：" + datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "備份來源：" + str(ROOT),
                "",
                "【裡面最重要的東西】",
                "　各專案的 _private 資料夾，裡面有：",
                "　　salt.txt　　　產生代碼用的鹽值",
                "　　mapping.csv　 代碼與真實姓名的對照表",
                "　這兩個檔一旦遺失，已經發出去的代碼就永遠還原不回姓名，",
                "　明年的年度作業也會接不上。",
                "",
                "【這份備份含真實姓名】",
                "　請比照公務個資保管：不要放進共用雲端硬碟、不要隨意轉寄。",
                "　這台電腦若要交回或報廢，這份備份必須先移到安全的地方。",
                "",
                "【以後要怎麼用回來】",
                "　1. 重新安裝假名化工具（解壓縮後放到本機硬碟）。",
                "　2. 把這份備份裡的「專案」整個資料夾，覆蓋回新的工具資料夾。",
                "　3. 開啟工具，選同一個專案，代碼就會和以前完全一致。",
                "",
                "【不確定的時候】",
                "　不要自己猜，先問資訊窗口或原分析人員。",
                "",
            ]) + "\r\n", encoding="utf-8")
    except OSError:
        pass
    return True


def verify_backup(target: Path, quiet=True):
    """來源與備份逐檔比對 SHA-256。任何一項對不上或算不出來，一律視為失敗。"""
    bad = []
    n = 0
    for p in sorted(iter_files(ROOT)):
        rel = p.relative_to(ROOT)
        dst = target / rel
        if not dst.exists():
            bad.append(f"備份裡找不到：{rel}")
            continue
        try:
            a, b = sha256(p), sha256(dst)
        except OSError as e:
            bad.append(f"讀不到（視為失敗）：{rel}　{e}")
            continue
        if len(a) != 64 or a != b:
            bad.append(f"內容不一致：{rel}")
            continue
        n += 1
    if not quiet:
        say(f"　　比對完成：{n:,} 個檔案一致，{len(bad)} 個有問題。")
    return (len(bad) == 0 and n > 0), bad


# ══════════════════════════════════════════════════════════════
#  桌面捷徑
# ══════════════════════════════════════════════════════════════

def desktop_dir():
    try:
        import winreg
        key = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key) as k:
            return Path(winreg.QueryValueEx(k, "Desktop")[0])
    except Exception:
        return Path.home() / "Desktop"


def documents_dir():
    try:
        import winreg
        key = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key) as k:
            return Path(winreg.QueryValueEx(k, "Personal")[0])
    except Exception:
        return Path.home() / "Documents"


def shortcut_target(lnk: Path):
    """讀 .lnk 指向哪裡。讀不到回 None。"""
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command",
             "$s=(New-Object -COM WScript.Shell).CreateShortcut($env:LNK);"
             "Write-Output $s.TargetPath; Write-Output $s.WorkingDirectory"],
            capture_output=True, text=True, timeout=30,
            env={**os.environ, "LNK": str(lnk)})
        return r.stdout.strip().splitlines()
    except Exception:
        return None


def remove_shortcut():
    lnk = desktop_dir() / "假名化工具.lnk"
    if not lnk.exists():
        say("　桌面上沒有「假名化工具」捷徑，不用移除。")
        return
    # 只認 TargetPath。工作目錄不能當判準——別的程式的捷徑也可能把
    # 工作目錄設在這個資料夾，那樣會誤刪別人的捷徑（實測踩過）。
    tgt = shortcut_target(lnk)
    target_path = tgt[0].strip() if tgt and tgt[0].strip() else ""
    if not target_path:
        say("　讀不到桌面捷徑指向哪裡，為了安全不動它：")
        say("　　" + str(lnk))
        say("　確認它是這個工具的捷徑之後，可以自己刪掉。")
        return
    try:
        inside = is_inside(Path(target_path), ROOT)
    except (OSError, ValueError):
        inside = False
    if not inside:
        say("　桌面上的「假名化工具」捷徑指向另外一個位置：")
        say("　　" + target_path)
        say("　這次不動它（可能是另一份工具的捷徑）。")
        return
    try:
        lnk.unlink()
        say("　已移除桌面捷徑：" + str(lnk))
    except OSError as e:
        say("　移除桌面捷徑失敗：" + str(e))


# ══════════════════════════════════════════════════════════════
#  刪除
# ══════════════════════════════════════════════════════════════

def remove_program_files():
    """只刪程式檔，使用者資料一概不動。"""
    gone, kept_err = [], []
    for name in PROGRAM_FILES:
        p = ROOT / name
        if p.exists():
            try:
                p.unlink()
                gone.append(name)
            except OSError as e:
                kept_err.append(f"{name}：{e}")
    for d in ROOT.rglob("__pycache__"):
        shutil.rmtree(d, ignore_errors=True)
    return gone, kept_err


CLEANUP = r'''# -*- coding: utf-8 -*-
"""解除安裝的收尾程式：先刪掉自己，再等主程式結束後刪掉整個工具資料夾。"""
import shutil, sys, time
from pathlib import Path

# 程式已經讀進記憶體，現在就能刪掉這個檔；放到最後的話，使用者直接按 X 關視窗就會留在暫存資料夾
try:
    Path(__file__).unlink()
except Exception:
    pass
target = Path(sys.argv[1])
backup = sys.argv[2] if len(sys.argv) > 2 else ""
print("正在移除：" + str(target))
ok = False
for attempt in range(8):
    time.sleep(2)
    shutil.rmtree(target, ignore_errors=True)
    if not target.exists():
        ok = True
        break
print()
print("=" * 66)
if ok:
    print("　假名化工具已經完整移除。")
else:
    print("　有部分檔案刪不掉（可能正被其他程式開著）。")
    print("　剩下的內容在：" + str(target))
    print("　請關掉相關程式後，自己把這個資料夾刪掉。")
if backup:
    print()
    print("　你的資料備份保留在：")
    print("　" + backup)
    print("　這份備份含真實姓名，請比照公務個資保管。")
print("=" * 66)
print()
try:
    input("按 Enter 關閉這個視窗…")
except Exception:
    pass
'''


def remove_everything(backup_dir: Path):
    """把整個工具資料夾刪掉。由另一個行程接手，因為不能刪掉自己正在用的資料夾。"""
    helper = Path(tempfile.gettempdir()) / f"假名化工具_收尾_{os.getpid()}.py"
    helper.write_text(CLEANUP, encoding="utf-8")
    # 先離開工具資料夾，否則它會被自己的行程佔住
    os.chdir(tempfile.gettempdir())
    exe = sys.executable
    flags = 0
    if os.name == "nt":
        flags = subprocess.CREATE_NEW_CONSOLE
    subprocess.Popen([exe, "-X", "utf8", str(helper), str(ROOT), str(backup_dir)],
                     creationflags=flags, close_fds=True)
    say()
    say("　已交給收尾程式處理，會另外開一個視窗。")
    say("　這個視窗現在要關閉了。")
    time.sleep(2)


# ══════════════════════════════════════════════════════════════
#  各選項
# ══════════════════════════════════════════════════════════════

def opt_backup_only(state):
    title("只備份，不移除任何東西")
    target = choose_backup_dir()
    if target is None:
        say("　已取消，沒有做任何事。")
        pause()
        return
    if do_backup(target):
        say()
        say(SEP)
        say("　備份完成，沒有移除任何東西。")
        say("　備份位置：" + str(target))
        say("　⚠　這份備份含真實姓名，請比照公務個資保管。")
        say(SEP)
        state["backup"] = target
    else:
        say()
        say("　備份沒有成功，什麼都沒有移除。")
    pause()


def opt_shortcut(state):
    title("只移除桌面捷徑")
    remove_shortcut()
    say()
    say("　程式與資料都沒有動。")
    pause()


def opt_program(state):
    info = state["info"]
    title("移除程式，保留資料")
    say("　會刪掉：這個資料夾裡的程式檔（.py、.bat、說明、圖示、設定範例）")
    say("　會留著：「專案」資料夾（input／output／_private）、欄位設定.txt，")
    say("　　　　　以及你自己放在這裡的其他檔案。")
    say()
    say("　也就是說 mapping.csv 與 salt.txt 都會留在原地。")
    say()
    if info["maps"]:
        say("　（這個選項不會碰到對照表，所以不需要備份。想保險的話可以先選 1。）")
        say()
    if not confirm("　確定要移除程式嗎？"):
        say("　已取消，沒有移除任何東西。")
        pause()
        return
    remove_shortcut()
    gone, err = remove_program_files()
    say()
    say(SEP)
    say(f"　已移除 {len(gone)} 個程式檔。")
    for e in err:
        say("　★ 刪不掉：" + e)
    say("　你的資料留在：" + str(ROOT))
    say(SEP)
    say()
    say("　（這支解除安裝程式本身也已經移除，所以不會再執行第二次。")
    say("　　之後若要連資料一起清掉，請用檔案總管自己刪——")
    say("　　那樣會丟進資源回收桶，比直接刪除安全。）")
    say()
    ask("按 Enter 結束…")
    sys.exit(0)


def opt_everything(state):
    info = state["info"]
    title("完整移除（連資料一起刪）")
    warn_mapping(info)
    say("　這個選項會把整個工具資料夾刪掉：")
    say("　　" + str(ROOT))
    say(f"　　共 {info['files']:,} 個檔案（{human(info['bytes'])}）")
    say()
    say("　規則：沒有備份、或備份驗證沒過，就不會刪。這一步不能跳過。")
    say()
    if not confirm("　要繼續嗎？"):
        say("　已取消，沒有刪除任何東西。")
        pause()
        return

    # 第一道：備份
    target = state.get("backup")
    if target and target.exists():
        say()
        say("　這次已經備份過了：" + str(target))
        if not confirm("　要沿用這份備份嗎？（選否會重做一份）"):
            target = None
    if not target:
        target = choose_backup_dir()
        if target is None or not do_backup(target):
            say()
            say(SEP)
            say("　沒有刪除任何東西。")
            say("　因為備份沒有做成功。在確定資料有另一份完整複本之前，")
            say("　這支程式不會刪你的資料夾。")
            say(SEP)
            pause()
            return
        state["backup"] = target

    # 第二道：刪之前再驗一次（備份完成到現在中間隔了一段時間）
    say()
    say("　刪除前再檢查一次備份……")
    ok, bad = verify_backup(target, quiet=False)
    if not ok:
        say()
        say("　★ 備份和現在的檔案對不起來，沒有刪除任何東西。")
        for b in bad[:10]:
            say("　　" + b)
        pause()
        return

    # 第三道：打字確認
    say()
    say("　" + "─" * 62)
    say("　你的資料已經完整備份在：")
    say("　　" + str(target))
    say("　請先親眼打開這個資料夾確認裡面有東西，再繼續。")
    say("　" + "─" * 62)
    say()
    say("　確認過了，請輸入六個大寫英文字母：DELETE")
    say("　（不想刪就直接按 Enter）")
    say()
    if ask("　請輸入：") != "DELETE":
        say()
        say("　你輸入的不是 DELETE，所以沒有刪除任何東西。")
        say("　剛才的備份仍然保留在：" + str(target))
        pause()
        return

    # 第四道：最後一次
    say()
    if not confirm("　最後一次確認，真的要刪掉整個工具資料夾嗎？"):
        say("　已取消，沒有刪除任何東西。備份仍保留在：" + str(target))
        pause()
        return

    remove_shortcut()
    remove_everything(target)
    sys.exit(0)


def opt_python_info(state):
    title("Python 與套件要怎麼移除")
    say("　這支程式「不會」自動移除它們，因為這台電腦的其他程式很可能也在用。")
    say("　如果你確定沒有別的程式需要，可以自己在「命令提示字元」執行：")
    say()
    say("　　移除套件（四個都是常見套件，移除前請先確認）：")
    say("　　　py -3 -m pip uninstall " + " ".join(PACKAGES))
    say()
    say("　　移除 Python 本體（只有當初是用安裝腳本自動裝的才適用）：")
    say("　　　winget uninstall --id Python.Python.3.12")
    say()
    exe = shutil.which("python") or shutil.which("py")
    if exe:
        say("　這台電腦目前的 Python：" + exe)
    else:
        say("　這台電腦目前找不到 Python。")
    say()
    say("　⚠　不確定的話就不要移除。留著 Python 不會有任何壞處。")
    pause()


# ══════════════════════════════════════════════════════════════
#  主選單
# ══════════════════════════════════════════════════════════════

MENU = (
    ("1", "只備份資料，不移除任何東西　　〔建議先做這個〕", opt_backup_only),
    ("2", "只移除桌面捷徑　　　　　　　　程式和資料都不動", opt_shortcut),
    ("3", "移除程式，保留資料　　　　　　「專案」資料夾原封不動留著", opt_program),
    ("4", "完整移除，連資料一起刪　　　　〔危險，必須先備份並驗證成功〕", opt_everything),
    ("5", "Python 和套件要怎麼移除　　　　只顯示說明，不會動到電腦", opt_python_info),
)


def main():
    if os.name == "nt":
        try:
            ctypes.windll.kernel32.SetConsoleTitleW("假名化工具 - 解除安裝")
        except Exception:
            pass
    refuse_if_not_tool_folder()
    refuse_if_dangerous_location()

    state = {"backup": None}
    while True:
        state["info"] = info = scan()
        title("上傳前假名化工具　解除安裝")
        say("　這支程式預設「什麼都不刪」。看不懂就直接把視窗關掉，不會有事。")
        say()
        show_status(info)
        warn_mapping(info)
        say("　請選擇：")
        say()
        for k, label, _ in MENU:
            say(f"　　{k}　{label}")
        say("　　0　離開，什麼都不做")
        say()
        choice = ask("　請輸入數字後按 Enter（直接按 Enter＝離開）：")
        if choice in ("", "0"):
            say()
            say("　結束，沒有做任何事。")
            return
        for k, _, fn in MENU:
            if choice == k:
                fn(state)
                break
        else:
            say("　沒有這個選項。")
            pause()


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except BaseException as e:            # noqa: BLE001  任何意外都要讓使用者看得到
        say()
        say(SEP)
        say("　程式發生預期外的錯誤，沒有繼續執行。")
        say(f"　{type(e).__name__}: {e}")
        say("　請把整個畫面截圖給資訊窗口。")
        say(SEP)
        try:
            input("按 Enter 關閉…")
        except Exception:
            pass
        sys.exit(1)
