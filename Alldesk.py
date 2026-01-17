import tkinter as tk
import shutil
import subprocess
import threading
import os
import stat
import glob
import platform
import sys
import tempfile
import winreg

# third-party / local imports
from openpyxl import load_workbook
from pathlib import Path
from tkinter import font as tkfont
from tkinter import ttk
from tkinter import messagebox

 
# 輕量 DES 實作（支援單一 8-byte 區塊的 ECB 加密）
# 提供兼容介面：DES.new(key, DES.MODE_ECB).encrypt(data)
class _DES:
    """
    簡易 DES 實作（支援單一 8-byte 區塊的 ECB 加密）。

    提供內部 API，模擬 Crypto 庫的行為，使呼叫端可以使用
    `DES.new(key, DES.MODE_ECB).encrypt(data)` 的介面。
    此實作僅用於相容性與小型工具，不建議用於生產環境。
    """

    def __init__(self, key: bytes):
        """初始化 DES 實例。

        參數:
        - key: 8 位元組的金鑰 (bytes)

        例外:
        - TypeError: key 非 bytes-like
        - ValueError: key 長度非 8
        """
        if not isinstance(key, (bytes, bytearray)):
            raise TypeError('key must be bytes')
        if len(key) != 8:
            raise ValueError('DES key must be 8 bytes')
        self.key = bytes(key)
        self.subkeys = self._generate_subkeys(self.key)

    @staticmethod
    def _bytes_to_bits(b: bytes):
        """將 bytes 轉為位元列表（MSB first）。

        輸入例: b"\x01" -> [0,0,0,0,0,0,0,1]
        """
        bits = []
        for byte in b:
            for i in range(8)[::-1]:
                bits.append((byte >> i) & 1)
        return bits

    @staticmethod
    def _bits_to_bytes(bits):
        """將位元列表 (MSB first) 轉回 bytes。

        只支援位元數為 8 的整數倍。
        """
        out = bytearray()
        for i in range(0, len(bits), 8):
            byte = 0
            for bit in bits[i:i+8]:
                byte = (byte << 1) | bit
            out.append(byte)
        return bytes(out)

    def _permute(self, table, bits):
        """依照 permutation 表對位元列表重新排列並回傳新列表。

        table 為 1-based 的索引表。
        """
        return [bits[i-1] for i in table]

    def _left_rotate(self, lst, n):
        """將序列向左旋轉 n 位元。

        用於 DES subkey 的 C/D bits 旋轉。
        """
        return lst[n:]+lst[:n]

    def _generate_subkeys(self, key8: bytes):
        """從 8-byte 原始金鑰產生 16 個 48-bit 子金鑰。

        回傳值為 list[list[int]]，每個子金鑰為位元 (0/1) 列表。
        """
        # PC-1
        pc1 = [57,49,41,33,25,17,9,
               1,58,50,42,34,26,18,
               10,2,59,51,43,35,27,
               19,11,3,60,52,44,36,
               63,55,47,39,31,23,15,
               7,62,54,46,38,30,22,
               14,6,61,53,45,37,29,
               21,13,5,28,20,12,4]

        # PC-2
        pc2 = [14,17,11,24,1,5,
               3,28,15,6,21,10,
               23,19,12,4,26,8,
               16,7,27,20,13,2,
               41,52,31,37,47,55,
               30,40,51,45,33,48,
               44,49,39,56,34,53,
               46,42,50,36,29,32]

        # rotations
        rotations = [1,1,2,2,2,2,2,2,1,2,2,2,2,2,2,1]

        key_bits = self._bytes_to_bits(key8)
        permuted = self._permute(pc1, key_bits)
        c = permuted[:28]
        d = permuted[28:]
        subkeys = []
        for r in rotations:
            c = self._left_rotate(c, r)
            d = self._left_rotate(d, r)
            cd = c + d
            sub = self._permute(pc2, cd)
            subkeys.append(sub)
        return subkeys

    def _feistel(self, r, subkey):
        """DES 的 Feistel 函數 (f 函數)。

        參數:
        - r: 右半部位元列表 (32 bits)
        - subkey: 本輪的子金鑰 (48 bits)

        回傳 32-bit 的位元列表。
        """
        # Expansion table
        e_table = [32,1,2,3,4,5,4,5,6,7,8,9,8,9,10,11,12,13,12,13,14,15,16,17,16,17,18,19,20,21,20,21,22,23,24,25,24,25,26,27,28,29,28,29,30,31,32,1]

        # S-boxes
        s_boxes = [
            # S1
            [
                14,4,13,1,2,15,11,8,3,10,6,12,5,9,0,7,
                0,15,7,4,14,2,13,1,10,6,12,11,9,5,3,8,
                4,1,14,8,13,6,2,11,15,12,9,7,3,10,5,0,
                15,12,8,2,4,9,1,7,5,11,3,14,10,0,6,13
            ],
            # S2
            [
                15,1,8,14,6,11,3,4,9,7,2,13,12,0,5,10,
                3,13,4,7,15,2,8,14,12,0,1,10,6,9,11,5,
                0,14,7,11,10,4,13,1,5,8,12,6,9,3,2,15,
                13,8,10,1,3,15,4,2,11,6,7,12,0,5,14,9
            ],
            # S3
            [
                10,0,9,14,6,3,15,5,1,13,12,7,11,4,2,8,
                13,7,0,9,3,4,6,10,2,8,5,14,12,11,15,1,
                13,6,4,9,8,15,3,0,11,1,2,12,5,10,14,7,
                1,10,13,0,6,9,8,7,4,15,14,3,11,5,2,12
            ],
            # S4
            [
                7,13,14,3,0,6,9,10,1,2,8,5,11,12,4,15,
                13,8,11,5,6,15,0,3,4,7,2,12,1,10,14,9,
                10,6,9,0,12,11,7,13,15,1,3,14,5,2,8,4,
                3,15,0,6,10,1,13,8,9,4,5,11,12,7,2,14
            ],
            # S5
            [
                2,12,4,1,7,10,11,6,8,5,3,15,13,0,14,9,
                14,11,2,12,4,7,13,1,5,0,15,10,3,9,8,6,
                4,2,1,11,10,13,7,8,15,9,12,5,6,3,0,14,
                11,8,12,7,1,14,2,13,6,15,0,9,10,4,5,3
            ],
            # S6
            [
                12,1,10,15,9,2,6,8,0,13,3,4,14,7,5,11,
                10,15,4,2,7,12,9,5,6,1,13,14,0,11,3,8,
                9,14,15,5,2,8,12,3,7,0,4,10,1,13,11,6,
                4,3,2,12,9,5,15,10,11,14,1,7,6,0,8,13
            ],
            # S7
            [
                4,11,2,14,15,0,8,13,3,12,9,7,5,10,6,1,
                13,0,11,7,4,9,1,10,14,3,5,12,2,15,8,6,
                1,4,11,13,12,3,7,14,10,15,6,8,0,5,9,2,
                6,11,13,8,1,4,10,7,9,5,0,15,14,2,3,12
            ],
            # S8
            [
                13,2,8,4,6,15,11,1,10,9,3,14,5,0,12,7,
                1,15,13,8,10,3,7,4,12,5,6,11,0,14,9,2,
                7,11,4,1,9,12,14,2,0,6,10,13,15,3,5,8,
                2,1,14,7,4,10,8,13,15,12,9,0,3,5,6,11
            ]
        ]

        # P permutation
        p_table = [16,7,20,21,29,12,28,17,1,15,23,26,5,18,31,10,2,8,24,14,32,27,3,9,19,13,30,6,22,11,4,25]

        # Expand r
        r_expanded = self._permute(e_table, r)
        # XOR with subkey
        xr = [a ^ b for a, b in zip(r_expanded, subkey)]
        # Split into 8 groups of 6
        out_bits = []
        for i in range(8):
            chunk = xr[i*6:(i+1)*6]
            row = (chunk[0] << 1) | chunk[5]
            col = (chunk[1] << 3) | (chunk[2] << 2) | (chunk[3] << 1) | chunk[4]
            s_val = s_boxes[i][row*16 + col]
            for bit in range(4)[::-1]:
                out_bits.append((s_val >> bit) & 1)
        # P permutation
        p_out = self._permute(p_table, out_bits)
        return p_out

    def encrypt(self, data: bytes) -> bytes:
        """對單一 8-byte 區塊進行 DES 加密（ECB, 單區塊）。

        例外:
        - TypeError: 非 bytes 輸入
        - ValueError: 長度非 8
        回傳: 加密後的 8-byte bytes
        """
        if not isinstance(data, (bytes, bytearray)):
            raise TypeError('data must be bytes')
        if len(data) != 8:
            raise ValueError('DES encrypt expects 8-byte block')

        # Initial permutation
        ip = [58,50,42,34,26,18,10,2,60,52,44,36,28,20,12,4,62,54,46,38,30,22,14,6,64,56,48,40,32,24,16,8,57,49,41,33,25,17,9,1,59,51,43,35,27,19,11,3,61,53,45,37,29,21,13,5,63,55,47,39,31,23,15,7]
        fp = [40,8,48,16,56,24,64,32,39,7,47,15,55,23,63,31,38,6,46,14,54,22,62,30,37,5,45,13,53,21,61,29,36,4,44,12,52,20,60,28,35,3,43,11,51,19,59,27,34,2,42,10,50,18,58,26,33,1,41,9,49,17,57,25]

        bits = self._bytes_to_bits(data)
        permuted = self._permute(ip, bits)
        l = permuted[:32]
        r = permuted[32:]

        for i in range(16):
            sub = self.subkeys[i]
            f_out = self._feistel(r, sub)
            new_r = [a ^ b for a, b in zip(l, f_out)]
            l = r
            r = new_r

        preoutput = r + l
        final_bits = self._permute(fp, preoutput)
        return self._bits_to_bytes(final_bits)


class DES:
    MODE_ECB = 1

    @staticmethod
    def new(key, mode=None):
        """工廠函式，回傳 _DES 實例以相容舊有介面。

        - 若傳入為 str，使用 latin-1 編碼轉為 bytes。
        - mode 目前僅為相容參數，未使用。
        """
        if isinstance(key, str):
            key = key.encode('latin-1')
        return _DES(key)
# VNC helper functions inlined (from former VNCdesk.py)

"""Alldesk GUI 啟動器

功能概述：
- 提供三個分頁：`RustDesk`、`AnyDesk` 與 `TightVNC`，分別對應三種遠端桌面
    連線方式與設定檔準備流程。
- 從 `Alldesk.xlsx` 讀取各分頁的客戶端清單（預期工作表名稱分別為
    'rustdesk'、'anydesk' 與第3張表用於 VNC）。程式以只讀方式讀取 Excel 作
    為單一資料來源，UI 上的快速連線按鈕與手動輸入皆來自該檔案內容。

各分頁主要行為：
- RustDesk：在啟動前於 `%APPDATA%\RustDesk\config` 產生或覆寫
    `RustDesk2.toml` 與 `peers/<id>.toml`，以預載密碼與視窗設定，然後啟動
    `rustdesk.exe --connect <id> --password <pwd>`。
- AnyDesk：在啟動前於 `%APPDATA%\AnyDesk\user.conf` 寫入 viewmode，
    並以命令列（管道）將密碼傳入啟動 AnyDesk。
- TightVNC：由第3張工作表讀取 host/port/password，生成 `vnc.vnc`
    選項檔（輸出到專案內 `./exe/vnc.vnc`），若有密碼則以 TightVNC 相容的
    加密格式轉換（內含簡易 DES 實作），再啟動 TightVNC 並指定該選項檔。

實作細節：
- 程式包含一個輕量的 DES 實作用於 TightVNC 密碼轉換（僅支援單一 8-byte
    區塊的 ECB 加密，為相容用途，不建議用於其他加密需求）。
- 程式會儘量在可用時使用 COM automation 開啟 Excel（以便指定工作表），
    否則會 fallback 至使用系統關聯或 excel.exe 啟動檔案。

安全/相容性說明：
- 此工具以提高便利性為主，檔案 I/O 與執行外部程式的作法會盡量處理常見
    錯誤（例：檔案不存在、唯讀屬性），但使用者應評估在公司環境或生產環境
    的安全性與授權。"""

# 預設值（可用環境變數覆寫）
# 將可執行檔統一放到專案內的 `exe` 資料夾（相對於此檔案），使用環境變數可覆寫
BASE_DIR = Path(__file__).resolve().parent
EXE_DIR = BASE_DIR / 'exe'
# rustdesk 可執行檔路徑（相對或絕對）
RUSTDESK_APP = os.getenv('RUSTDESK_APP', str(EXE_DIR / 'rustdesk.exe'))
# 用於產生 RustDesk2.toml 的 rendezvous server 與 key（固定參數）
RUSTDESK_HOST = 'everdura.ddnsfree.com'
RUSTDESK_KEY = 'kCC8dq5x8uvEI+fpbIsTpYhCMaMbAxpYmGv6XtR7NsY='

# AnyDesk / TightVNC 可執行檔路徑
ANYDESK_APP = os.getenv('ANYDESK_APP', str(EXE_DIR / 'AnyDesk.exe'))
TIGHTVNC_APP = os.getenv('TIGHTVNC_APP', str(EXE_DIR / 'TightVNC.exe'))

# Base dir for resources when bundled with PyInstaller
VNC_BASE_DIR = getattr(sys, '_MEIPASS', None) or str(Path(__file__).resolve().parent)

def resource_path(filename: str) -> str:
    """取得打包後或開發模式下的資源檔案絕對路徑。

    參數:
    - filename: 相對於資源根目錄的檔案名稱

    回傳: 平台相容的絕對路徑字串
    """
    return os.path.join(VNC_BASE_DIR, filename)


def get_app_path(filename: str) -> str:
    """回傳應用程式相對的檔案路徑：
    - 若為 PyInstaller onefile/frozen，使用可執行檔所在資料夾 (sys.executable)
    - 否則使用原始 `BASE_DIR`（原始原始碼所在資料夾）
    """
    try:
        if getattr(sys, 'frozen', False):
            return os.path.join(os.path.dirname(sys.executable), filename)
    except Exception:
        pass
    return os.path.join(str(BASE_DIR), filename)


def _find_excel_exe() -> str | None:
    """嘗試從登錄檢索 excel.exe 的路徑，找不到則回傳 None。"""
    try:
        # 優先查詢 HKLM / HKCU App Paths
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                key = winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe")
                try:
                    exe_path, _ = winreg.QueryValueEx(key, None)
                    if exe_path:
                        return exe_path
                finally:
                    winreg.CloseKey(key)
            except FileNotFoundError:
                # 嘗試 WOW6432Node
                try:
                    key = winreg.OpenKey(hive, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe")
                    try:
                        exe_path, _ = winreg.QueryValueEx(key, None)
                        if exe_path:
                            return exe_path
                    finally:
                        winreg.CloseKey(key)
                except FileNotFoundError:
                    continue
        # 再試 HKEY_CLASSES_ROOT 的 ProgID
        try:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CurVer")
            winreg.CloseKey(key)
            # 若存在就回傳 None 表示安裝，但沒有取得路徑
            return None
        except Exception:
            return None
    except Exception:
        return None


def open_alldesk_excel(sheet_idx: int | None = None):
    """開啟 `Alldesk.xlsx`，若系統可用 COM automation，則嘗試選取指定工作表。

    參數:
    - sheet_idx: 1-based 的工作表索引，若為 None 則不指定。
    """
    xlsx = get_app_path('Alldesk.xlsx')
    if not os.path.exists(xlsx):
        messagebox.showwarning('找不到檔案', '找不到 Alldesk.xlsx')
        return

    # 先嘗試使用 winreg 取得 excel.exe 路徑（純啟動用途的 fallback）
    exe_path = _find_excel_exe()

    # 優先使用 COM automation 控制 Excel（若可用），以便能選擇工作表
    try:
        import win32com.client
        try:
            excel = win32com.client.GetActiveObject('Excel.Application')
        except Exception:
            excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = True

        # 檢查檔案是否已開啟
        wb = None
        try:
            for w in excel.Workbooks:
                try:
                    if os.path.normcase(w.FullName) == os.path.normcase(xlsx):
                        wb = w
                        break
                except Exception:
                    continue
        except Exception:
            pass

        if wb is None:
            wb = excel.Workbooks.Open(xlsx)

        if sheet_idx:
            try:
                # win32com 支援 Worksheets(1-based)
                ws = wb.Worksheets(sheet_idx)
                ws.Activate()
            except Exception:
                pass
        return
    except Exception:
        # 若無法使用 COM（例如未安裝 pywin32），退回至啟動檔案或使用 excel.exe
        pass

    # 若有可執行檔路徑就用它開啟，否則用系統預設關聯
    if exe_path:
        try:
            subprocess.Popen([exe_path, xlsx])
            return
        except Exception:
            pass

    try:
        os.startfile(xlsx)
    except Exception:
        messagebox.showwarning('未安裝 Excel', '此電腦未偵測到 Microsoft Excel，無法以 Excel 開啟 Alldesk.xlsx')


def get_writable_dir() -> str:
    """回傳一個在此環境中可寫入的目錄。

    - 若為封裝後的執行檔（frozen），使用系統暫存目錄。
    - 開發模式則回傳此原始檔所在資料夾。
    """
    if getattr(sys, 'frozen', False):
        return tempfile.gettempdir()
    return os.path.dirname(__file__)


def encrypt_tightvnc_password(password: str) -> str:
    """將 TightVNC 的純文字密碼轉為 vnc 設定檔所使用的加密十六進位字串。

    演算法說明：
    - 取前 8 個 ASCII 字元，不足以 NUL 填充。
    - 使用 TightVNC 固定的 challenge bytes，對每個 byte 做 bit-reverse，
      將結果當作 DES key，使用 ECB 加密密碼區塊後回傳 hex 表示。
    """
    # take up to first 8 ASCII chars, pad with NULs
    pw = (password or '')[:8].encode('ascii', errors='ignore')
    pw = pw.ljust(8, b'\x00')

    # TightVNC uses a fixed challenge bytes; reverse bits in each challenge byte to form DES key
    challenge = [23, 82, 107, 6, 35, 78, 88, 7]

    def rev_bits_byte(b):
        b = ((b & 0xF0) >> 4) | ((b & 0x0F) << 4)
        b = ((b & 0xCC) >> 2) | ((b & 0x33) << 2)
        b = ((b & 0xAA) >> 1) | ((b & 0x55) << 1)
        return b

    key = bytes([rev_bits_byte(b) for b in challenge])
    cipher = DES.new(key, DES.MODE_ECB)
    return cipher.encrypt(pw).hex()


def create_header_row(parent, on_connect, with_port=False, default_port='5900'):
    """建立各遠端分頁共用的標頭區塊（輸入欄位與連接按鈕）。

    參數:
    - parent: 要放置 header 的父容器 (tk widget)
    - on_connect: 當使用者按下「連接」按鈕時的回呼，會傳入 (id, pwd, port)
    - with_port: 是否顯示埠號輸入欄
    - default_port: 埠號欄位的預設值

    回傳: tuple (ent_id, ent_pwd, ent_port) — 若 `with_port` 為 False，則
    ent_port 會是 None。
    """
    header = ttk.Frame(parent)
    header.grid(row=0, column=0, columnspan=10, sticky='w')

    # 連接 ID
    f_id = ttk.Frame(header)
    f_id.pack(side='left', padx=10)
    tk.Label(f_id, text="連接ID:").pack(side='left')
    ent_id = tk.Entry(f_id, width=28)
    ent_id.pack(side='left', padx=6)

    # 密碼
    f_pwd = ttk.Frame(header)
    f_pwd.pack(side='left', padx=10)
    tk.Label(f_pwd, text="密碼:").pack(side='left')
    ent_pwd = tk.Entry(f_pwd, show='*', width=30)
    ent_pwd.pack(side='left', padx=6)

    # 連接按鈕
    # ent_port 可能尚未定義，故預先建立為 None，lambda 內再讀取
    ent_port = None
    def _on_click():
        on_connect(ent_id.get(), ent_pwd.get(), ent_port.get() if with_port and ent_port is not None else None)

    btn = tk.Button(header, text="連接", command=_on_click)
    btn.pack(side='left', padx=6)

    # 埠（可選）
    if with_port:
        f_port = ttk.Frame(header)
        f_port.pack(side='left', padx=10)
        tk.Label(f_port, text="埠:").pack(side='left')
        ent_port = tk.Entry(f_port, width=8)
        ent_port.pack(side='left', padx=6)
        ent_port.insert(0, default_port)

    return ent_id, ent_pwd, ent_port

class RustDesk():
    """RustDesk 分頁：從 Excel 載入 client 並發起 RustDesk 連線。

    主要職責：
    - 從 `Alldesk.xlsx` 的 'rustdesk' 工作表讀取客戶清單。
    - 在啟動連線前於 %AppData%/RustDesk/config 產生或覆寫 `RustDesk2.toml`，
      並在 peers 資料夾中寫入 `<ID>.toml`，以調整視窗/預載密碼等設定。
    """
    def __init__(self, notebook: ttk.Notebook):
        """建立 RustDesk 分頁物件並初始化其 UI 與資料。

        傳入 `notebook` 並呼叫 `init_rustdesk` 進行資料讀取與 UI 容器建立。
        """
        self.init_rustdesk(notebook)

    def init_rustdesk(self, notebook: ttk.Notebook):
                """初始化 RustDesk 分頁：

                - 讀取 `Alldesk.xlsx` 的 'rustdesk' 工作表（或第一張表），
                    解析成 (tag, id, password) 的 client 列表。
                - 正規化 rustdesk 可執行檔路徑並建立 UI 容器。
                """
                app = RUSTDESK_APP
        # 讀取 `Alldesk.xlsx` 的 'rustdesk' 工作表；若不存在則不載入 client（僅支援 Excel）
        excel_path = Path(get_app_path('Alldesk.xlsx'))
        # 使用模組級常數 `RUSTDESK_HOST` / `RUSTDESK_KEY`，不在物件上存放副本
        clients = []
        if excel_path.exists():
            try:
                wb = load_workbook(filename=str(excel_path), read_only=True, data_only=True)
                # 優先嘗試按工作表名稱取表，否則使用第一張
                if 'rustdesk' in wb.sheetnames:
                    ws = wb['rustdesk']
                else:
                    ws = wb[wb.sheetnames[0]] if wb.sheetnames else None
                rows = []
                if ws is not None:
                    for r in ws.iter_rows(values_only=True):
                        rows.append(['' if v is None else str(v) for v in r])
                # 只取前三欄 (tag, id, password)，若不足則以空字串補齊
                clients = [
                    (row[0] if len(row) > 0 else '', row[1] if len(row) > 1 else '', row[2] if len(row) > 2 else '')
                    for row in rows
                ]
            except Exception:
                clients = []
        else:
            clients = []
        # 使用固定 rustdesk 可執行檔，並正規化路徑
        exec_target = os.path.normpath(RUSTDESK_APP)

        self.exec_target = exec_target
        self.clients = clients
        self.frame = ttk.Frame(notebook)
        notebook.add(self.frame, text = 'RustDesk')

    def _prepare_rustdesk_conf(self, client_id: str, password: str):
        r"""在 %AppData%\RustDesk\config 下準備 RustDesk 設定。

        功能：
        - 刪除 peers 資料夾中可能殘留的 ThreadId 檔案。
        - 在 peers 下建立或覆寫 `<client_id>.toml`（包含視窗、自適應等設定）。
        - 產生 `RustDesk2.toml`，寫入 rendezvous server、relay、key 與預載密碼。

        參數：
        - client_id: 目標機器 ID。
        - password: 該機器的密碼（會寫入 RustDesk2.toml 的 peer_settings）。
        """
        appdata = os.getenv('APPDATA')
        if not appdata:
            return
        cfg_dir = os.path.join(appdata, 'RustDesk', 'config')
        peers_dir = os.path.join(cfg_dir, 'peers')
        Path(peers_dir).mkdir(parents=True, exist_ok=True)

        # 刪除 peers 目錄內所有包含 ThreadId 的臨時檔（避免衝突）
        try:
            for f in glob.glob(os.path.join(peers_dir, '*ThreadId*')):
                try:
                    os.remove(f)
                except Exception:
                    pass
        except Exception:
            pass

        # 嘗試從 Alldesk.xlsx 第1張工作表讀取第一列的 ID/密碼（優先），若無則使用傳入的 client_id/password
        excel_path = Path(get_app_path('Alldesk.xlsx'))
        if excel_path.exists():
            try:
                wb0 = load_workbook(filename=str(excel_path), read_only=True, data_only=True)
                ws0 = wb0[wb0.sheetnames[0]] if wb0.sheetnames else None
                sheet_id = ''
                sheet_pwd = ''
                if ws0 is not None:
                    rows0 = [tuple('' if v is None else v for v in r) for r in ws0.iter_rows(values_only=True)]
                    # 需有 header + data 才視為有可取的預設值（模擬 pandas 讀入後 df.empty 檢查）
                    if len(rows0) >= 2:
                        headers = [str(c).lower() if c is not None else '' for c in rows0[0]]
                        data_row = [str(v) if v is not None else '' for v in rows0[1]]
                        id_idx = None
                        pwd_idx = None
                        for i, c in enumerate(headers):
                            if id_idx is None and 'id' in c:
                                id_idx = i
                            if pwd_idx is None and ('pass' in c or 'password' in c):
                                pwd_idx = i
                        cols_count = len(headers)
                        if id_idx is None and cols_count >= 2:
                            id_idx = 1
                        if pwd_idx is None and cols_count >= 3:
                            pwd_idx = 2
                        if id_idx is None:
                            id_idx = 0
                        if pwd_idx is None:
                            pwd_idx = id_idx + 1 if cols_count > id_idx + 1 else id_idx
                        try:
                            sheet_id = str(data_row[id_idx]).strip()
                        except Exception:
                            sheet_id = ''
                        try:
                            sheet_pwd = str(data_row[pwd_idx]).strip()
                        except Exception:
                            sheet_pwd = ''
                    # 只有在傳入的 client_id / password 為空時，才使用 Excel 第1張的預設值
                    try:
                        client_id_empty = not client_id or str(client_id).strip() == ''
                    except Exception:
                        client_id_empty = True
                    try:
                        password_empty = not password or str(password).strip() == ''
                    except Exception:
                        password_empty = True
                    if client_id_empty and sheet_id:
                        client_id = sheet_id
                    if password_empty and sheet_pwd:
                        password = sheet_pwd
            except Exception:
                pass

        peer_file = os.path.join(peers_dir, f"{client_id}.toml")
        # 若檔案存在，確保可寫（移除唯讀屬性，使用 windows attrib -r 更可靠）
        try:
            if os.path.exists(peer_file):
                try:
                    subprocess.run(['attrib', '-r', peer_file], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                except Exception:
                    os.chmod(peer_file, stat.S_IWRITE)
        except Exception:
            pass

        # 準備 peers/<id>.toml 內容（參考提供的 batch 範本）
        try:
            uname = os.getlogin() if hasattr(os, 'getlogin') else ''
        except Exception:
            uname = ''
        host = platform.node() or 'localhost'
        peer_content = (
            "password = []\n"
            "size = [0, 0, 0, 0]\n"
            "size_ft = [0, 0, 0, 0]\n"
            "size_pf = [0, 0, 0, 0]\n"
            "view_style = 'adaptive'\n"
            "scroll_style = 'scrollauto'\n"
            "edge_scroll_edge_thickness = 100\n"
            "image_quality = 'balanced'\n"
            "custom_image_quality = [50]\n"
            "show_remote_cursor = false\n"
            "lock_after_session_end = false\n"
            "terminal-persistent = false\n"
            "privacy_mode = false\n"
            "allow_swap_key = false\n"
            "port_forwards = []\n"
            "direct_failures = 1\n"
            "disable_audio = false\n"
            "disable_clipboard = false\n"
            "enable-file-copy-paste = true\n"
            "show_quality_monitor = false\n"
            "follow_remote_cursor = false\n"
            "follow_remote_window = false\n"
            "keyboard_mode = 'map'\n"
            "view_only = false\n"
            "show_my_cursor = false\n"
            "sync-init-clipboard = false\n"
            "trackpad-speed = 100\n\n"
            "[options]\n"
            "swap-left-right-mouse = ''\n"
            "codec-preference = 'auto'\n"
            "collapse_toolbar = ''\n"
            "zoom-cursor = ''\n"
            "i444 = ''\n"
            "custom-fps = '30'\n\n"
            "[ui_flutter]\n"
            "wm_RemoteDesktop = '{" + '"width":1300.0,"height":740.0,"offsetWidth":1270.0,"offsetHeight":710.0,"isMaximized":true,"isFullscreen":false' + "}'\n\n"
            "[info]\n"
            f"username = '{uname}'\n"
            f"hostname = '{host}'\n"
            "platform = 'Windows'\n"
        )
        try:
            # 使用原生文字模式寫入（預設會在 Windows 翻譯為 CRLF）並確保 flush+fsync
            with open(peer_file, 'w', encoding='utf-8') as fw:
                fw.write(peer_content)
                try:
                    fw.flush()
                    os.fsync(fw.fileno())
                except Exception:
                    pass
        except Exception:
            pass

        # 寫入或覆寫主設定檔 RustDesk2.toml（包含 rendezvous 與 peer 密碼）
        cfg_file = os.path.join(cfg_dir, 'RustDesk2.toml')
        try:
            rendezvous = f"{RUSTDESK_HOST}:21116" if RUSTDESK_HOST else ''
            key = RUSTDESK_KEY
            with open(cfg_file, 'w', encoding='utf-8') as fw:
                fw.write(f"rendezvous_server = '{rendezvous}'\n")
                fw.write("nat_type = 1\n")
                fw.write("[options]\n")
                fw.write(f"custom-rendezvous-server = '{RUSTDESK_HOST}'\n")
                fw.write(f"relay-server = '{RUSTDESK_HOST}'\n")
                fw.write(f"key = '{key}'\n")
                fw.write(f"[peer_settings.{client_id}]\n")
                fw.write(f"password = '{password}'\n")
                try:
                    fw.flush()
                    os.fsync(fw.fileno())
                except Exception:
                    pass
        except Exception:
            pass

  

    def run_rustdesk(self, client_id, password):
        """啟動 RustDesk 連線（RustDesk 專用）。

        步驟：
        1. 呼叫 `_prepare_rustdesk_conf`，在 %AppData% 準備必要設定檔。
        2. 以參數列表方式啟動 `rustdesk.exe --connect <id> --password <pwd>`。
        """
        exec_target = self.exec_target
        cmd = [exec_target, '--connect', client_id, '--password', password]
        # 在啟動 RustDesk 前，先寫入 RustDesk2.toml 以設定 view_style
        self._prepare_rustdesk_conf(client_id, password)
        # 以非同步方式啟動 RustDesk，模擬 batch 的 `start` 行為（立即返回）
        exe_dir = os.path.dirname(exec_target) or None
        try:
            subprocess.Popen(cmd, creationflags=subprocess.CREATE_NEW_CONSOLE, cwd=exe_dir)
        except Exception:
            try:
                subprocess.Popen(cmd, cwd=exe_dir)
            except Exception:
                try:
                    subprocess.Popen(cmd)
                except Exception:
                    pass


    def set_elements_rustdesk(self):
        """建立 RustDesk 分頁的 UI 元件。

        元件包括：
        - 輸入欄位：連接 ID、密碼
        - 連接按鈕：手動輸入後啟動連線
        """
        # 使用共用 header
        create_header_row(
            self.frame,
            on_connect=lambda cid, pwd, _: self.run_rustdesk(cid, pwd),
            with_port=False
        )

        ttk.Separator(self.frame, orient='horizontal').grid(row=1, column=0, columnspan=10, sticky='ew', padx=10, pady=5)

        # 按鈕容器，避免 header 欄位影響按鈕排版
        btn_container = ttk.Frame(self.frame)
        btn_container.grid(row=2, column=0, columnspan=10, sticky='w')
        row, col = 0, 0
        for client in self.clients:
            # client 可能是長度不一的 tuple/list，安全取值
            try:
                tag = client[0]
            except Exception:
                tag = ''
            try:
                client_id = client[1]
            except Exception:
                client_id = ''
            try:
                password = client[2]
            except Exception:
                password = ''
            # 避免把 Excel 的表頭當成按鈕（例如：'設備名稱', 'ID', 'Item'）
            if isinstance(tag, str) and tag.strip().lower() in ('設備名稱', 'id', 'item', 'name'):
                continue
            if isinstance(client_id, str) and client_id.strip().lower() in ('設備名稱', 'id', 'item', 'name'):
                continue
            tk.Button(btn_container, text=f"{tag}\n{client_id}", font=('微軟正黑體',10), width=15, height=4, 
                command = lambda cid = client_id, pwd = password: self.run_rustdesk(cid, pwd)
            ).grid(row=row, column=col, padx=3, pady=3)
            col += 1
            if col >= 10:
                col = 0
                row += 1
    
class AnyDesk():
    """AnyDesk 分頁：從 Excel 載入 client 並啟動 AnyDesk 連線。

    主要職責：
    - 從 `Alldesk.xlsx` 的 'anydesk' 工作表讀取客戶清單。
    - 在啟動 AnyDesk 前於 %AppData%/AnyDesk 寫入 `user.conf`，以控制視圖模式。
    """
    def __init__(self, notebook: ttk.Notebook):
        """建立 AnyDesk 分頁物件並初始化其 UI 與資料。

        傳入 `notebook` 並呼叫 `init_anydesk` 讀取 Excel 並準備按鈕與執行檔路徑。
        """
        self.init_anydesk(notebook)

    def init_anydesk(self, notebook: ttk.Notebook):
                """初始化 AnyDesk 分頁：

                - 讀取 `Alldesk.xlsx` 的 'anydesk' 工作表（或第二張表），
                    解析成 (tag, id, password) 的 client 列表。
                - 正規化 AnyDesk 可執行檔路徑並建立 UI 容器。
                """
                app: str = ANYDESK_APP
        # 讀取 `Alldesk.xlsx` 的 'anydesk' 工作表；若不存在則不載入 client（僅支援 Excel）
        excel_path = Path(get_app_path('Alldesk.xlsx'))
        clients = []
        if excel_path.exists():
            try:
                wb = load_workbook(filename=str(excel_path), read_only=True, data_only=True)
                # 優先嘗試按工作表名稱取表，否則使用第二張（index=1）
                if 'anydesk' in wb.sheetnames:
                    ws = wb['anydesk']
                else:
                    ws = wb[wb.sheetnames[1]] if len(wb.sheetnames) > 1 else (wb[wb.sheetnames[0]] if wb.sheetnames else None)
                rows = []
                if ws is not None:
                    for r in ws.iter_rows(values_only=True):
                        rows.append(['' if v is None else str(v) for v in r])
                clients = [
                    (row[0] if len(row) > 0 else '', row[1] if len(row) > 1 else '', row[2] if len(row) > 2 else '')
                    for row in rows
                ]
            except Exception:
                clients = []
        else:
            clients = []
        exec_target = os.path.normpath(app)

        self.exec_target = exec_target
        self.clients = clients
        self.frame = ttk.Frame(notebook)
        notebook.add(self.frame, text = 'AnyDesk')

    def _prepare_anydesk_conf(self, client_id: str):
        r"""在 %AppData%\AnyDesk 下建立 `user.conf` 並設定 viewmode。

        只寫入最小內容：`ad.session.viewmode=<client_id>:2`，用以在啟動 AnyDesk 時
        影響視窗顯示模式（例如強制開啟為檢視模式或預設尺寸）。
        """
        appdata = os.getenv('APPDATA')
        if not appdata:
            return
        anydesk_dir = os.path.join(appdata, 'AnyDesk')
        Path(anydesk_dir).mkdir(parents=True, exist_ok=True)
        conf_file = os.path.join(anydesk_dir, 'user.conf')
        try:
            with open(conf_file, 'w', encoding='utf-8') as fw:
                fw.write(f"ad.session.viewmode={client_id}:2\n")
        except Exception:
            pass

    def run_anydesk(self, client_id, password):
        r"""啟動 AnyDesk 連線（AnyDesk 專用）。

        步驟：
        1. 呼叫 `_prepare_anydesk_conf`，將 viewmode 寫入 `%APPDATA%\AnyDesk\user.conf`。
        2. 以非同步方式呼叫 AnyDesk，並透過命令列管道傳入密碼。
        """
        exec_target = self.exec_target
        # 在啟動 AnyDesk 前，先寫入 user.conf 以設定 viewmode
        self._prepare_anydesk_conf(client_id)

        # 使用 cmd 管道傳入密碼並以非同步方式啟動 AnyDesk
        command = f'cmd /c echo {password} | {exec_target} "{client_id}" --with-password'
        subprocess.Popen(command, creationflags = subprocess.CREATE_NO_WINDOW)

    def set_elements_anydesk(self):
        """建立 AnyDesk 分頁的 UI 元件。

        元件包括：
        - 手動輸入 ID/密碼 的欄位與連接按鈕
        - 由 Excel 載入的快速連線按鈕
        """
        # 使用共用 header
        create_header_row(
            self.frame,
            on_connect=lambda cid, pwd, _: self.run_anydesk(cid, pwd),
            with_port=False
        )

        ttk.Separator(self.frame, orient='horizontal').grid(row=1, column=0, columnspan=10, sticky='ew', padx=10, pady=5)

        # 按鈕容器，避免 header 欄位影響按鈕排版
        btn_container = ttk.Frame(self.frame)
        btn_container.grid(row=2, column=0, columnspan=10, sticky='w')
        row, col = 0, 0
        for client in self.clients:
            try:
                tag = client[0]
            except Exception:
                tag = ''
            try:
                client_id = client[1]
            except Exception:
                client_id = ''
            try:
                password = client[2]
            except Exception:
                password = ''
            # 避免把 Excel 的表頭當成按鈕（例如：'設備名稱', 'ID', 'Item'）
            if isinstance(tag, str) and tag.strip().lower() in ('設備名稱', 'id', 'item', 'name'):
                continue
            if isinstance(client_id, str) and client_id.strip().lower() in ('設備名稱', 'id', 'item', 'name'):
                continue
            tk.Button(btn_container, text=f"{tag}\n{client_id}", font=('微軟正黑體',10), width=15, height=4, 
                command = lambda cid = client_id, pwd = password: self.run_anydesk(cid, pwd)
            ).grid(row=row, column=col, padx=3, pady=3)
            col += 1
            if col >= 10:
                col = 0
                row += 1
    
class TightVNC():
    """VNC 分頁：從 Alldesk.xlsx 第3張工作表載入項目並啟動 VNC 連線。

    欄位對應：
    - Item: 顯示在按鈕上的名稱
    - URL: 目標主機（按鈕上顯示）
    - Password: 連線密碼（按鈕上不顯示）
    - Port: 連接埠（按鈕上不顯示）
    """
    def __init__(self, notebook: ttk.Notebook):
        """建立 TightVNC 分頁物件並從第3張工作表讀取 VNC 連線項目。

        會將讀取結果存在 `self.clients`，並在 `set_elements_tightvnc` 中
        產生 UI 按鈕用於快速連線。
        """
        app = 'vnc'
        excel_path = Path('./Alldesk.xlsx')
        clients = []
        if excel_path.exists():
            try:
                # 讀取第3張工作表（index=2）並轉為字典列（保留欄名以供對應）
                wb = load_workbook(filename=str(excel_path), read_only=True, data_only=True)
                if len(wb.sheetnames) > 2:
                    ws = wb[wb.sheetnames[2]]
                elif wb.sheetnames:
                    ws = wb[wb.sheetnames[0]]
                else:
                    ws = None
                clients = []
                if ws is not None:
                    rows = [tuple('' if v is None else v for v in r) for r in ws.iter_rows(values_only=True)]
                    if rows:
                        headers = [str(h) if h is not None else '' for h in rows[0]]
                        for r in rows[1:]:
                            rec = {headers[i]: ('' if r[i] is None else str(r[i])) if i < len(r) else '' for i in range(len(headers))}
                            clients.append(rec)
            except Exception:
                clients = []
        else:
            clients = []

        self.exec_target = TIGHTVNC_APP
        self.clients = clients
        self.frame = ttk.Frame(notebook)
        notebook.add(self.frame, text = 'TightVNC')

    def _prepare_and_launch_tightvnc(self, host: str, port: str, password: str):
        """準備 TightVNC 的 `vnc.vnc` 選項檔並啟動 TightVNC。

        功能：
        - 讀取預設範本 `vnc.vnc`（資源），在 [connection] 區段替換 host/port/password
        - 若缺少 [connection] 或 [options] 則補上合理的預設值
        - 將處理後的設定寫入專案內 `./exe/vnc.vnc`，並傳給 TightVNC 執行檔
        """
        vnc_source = resource_path('vnc.vnc')
        if os.path.exists(vnc_source):
            try:
                with open(vnc_source, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
            except Exception:
                lines = []
        else:
            lines = []

        out = []
        in_conn = False
        replaced = {'host': False, 'port': False, 'password': False}
        for i, line in enumerate(lines):
            s = line.strip()
            if s.lower() == '[connection]':
                in_conn = True
                out.append(line)
                continue
            if in_conn:
                if s.startswith('[') and s.endswith(']'):
                    in_conn = False
                    out.append(line)
                    continue
                if s.lower().startswith('host='):
                    out.append(f'host={host}\n')
                    replaced['host'] = True
                    continue
                if s.lower().startswith('port='):
                    out.append(f'port={port}\n')
                    replaced['port'] = True
                    continue
                if s.lower().startswith('password='):
                    if password:
                        enc_pw = encrypt_tightvnc_password(password)
                        out.append(f'password={enc_pw}\n')
                        replaced['password'] = True
                    else:
                        out.append(line)
                    continue
            out.append(line)

        if not any(l.strip().lower() == '[connection]' for l in out):
            conn_block = ["[connection]\n", f"host={host}\n", f"port={port}\n"]
            if password:
                enc_pw = encrypt_tightvnc_password(password)
                conn_block.append(f"password={enc_pw}\n")
            out = conn_block + ['\n'] + out
        else:
            if not (replaced['host'] and replaced['port'] and replaced['password']):
                new_out = []
                i = 0
                while i < len(out):
                    new_out.append(out[i])
                    if out[i].strip().lower() == '[connection]':
                        j = i + 1
                        consume = []
                        while j < len(out) and not out[j].strip().startswith('['):
                            consume.append(out[j])
                            j += 1
                        conn_lines = [f'host={host}\n', f'port={port}\n']
                        if password:
                                        enc_pw = encrypt_tightvnc_password(password)
                                        conn_lines.append(f'password={enc_pw}\n')
                        else:
                            for c in consume:
                                if c.strip().lower().startswith('password='):
                                    conn_lines.append(c)
                                    break
                        new_out.extend(conn_lines)
                        i = j
                        continue
                    i += 1
                out = new_out

                # 確保 [options] 區段存在，且強制設定允許遠端控制（非唯讀）
                def ensure_options(lines):
                    has_options = False
                    i = 0
                    while i < len(lines):
                        if lines[i].strip().lower() == '[options]':
                            has_options = True
                            # 收集直到下一個區段
                            j = i + 1
                            opts = {}
                            while j < len(lines) and not lines[j].strip().startswith('['):
                                s = lines[j].strip()
                                if '=' in s:
                                    k, v = s.split('=', 1)
                                    opts[k.strip().lower()] = v.strip()
                                j += 1
                            # 強制指定必要的鍵值
                            opts['viewonly'] = '0'
                            opts['shared'] = '1'
                            opts['swapmouse'] = opts.get('swapmouse', '0')
                            # 重建 options 區塊
                            new_block = ['[options]\n']
                            for k, v in opts.items():
                                new_block.append(f'{k}={v}\n')
                            # 取代第 i..j-1 行
                            lines[i:j] = new_block
                            break
                        i += 1
                    if not has_options:
                        # 在 connection 後附加 options 區塊
                        opts_block = ['[options]\n', 'viewonly=0\n', 'shared=1\n', 'swapmouse=0\n', '\n']
                        lines.extend(opts_block)
                    return lines

                out = ensure_options(out)

        # 將輸出改為專案內的 ./exe 資料夾（相對路徑），並確保該資料夾存在
        try:
            Path(EXE_DIR).mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        out_path = os.path.join(str(EXE_DIR), 'vnc.vnc')
        try:
            with open(out_path, 'w', encoding='utf-8') as f:
                f.writelines(out)
        except Exception:
            return

        exe_path = resource_path('TightVNC.exe')
        if not os.path.exists(exe_path):
            exe_path = TIGHTVNC_APP
        if not os.path.exists(exe_path):
            exe_path = 'TightVNC.exe'
        args = [exe_path, f'-optionsfile={out_path}', '-showcontrols=no']
        try:
            subprocess.Popen(args, cwd=get_writable_dir())
        except Exception:
            pass

    def run_tightvnc(self, item, url, password, port):
        """啟動 TightVNC 連線的高階介面。

        參數:
        - item: 顯示用的項目名稱（不影響連線）
        - url: 目標主機
        - password: 連線密碼（可空）
        - port: 埠號（預設 5900）
        """
        host = url or ''
        prt = port or '5900'
        self._prepare_and_launch_tightvnc(host, prt, password)

    def set_elements_tightvnc(self):
        """建立 TightVNC 分頁的 UI 元件。使用共用 header（含埠）。"""
        create_header_row(
            self.frame,
            on_connect=lambda cid, pwd, port: self.run_tightvnc('', cid, pwd, port),
            with_port=True,
            default_port='5900'
        )

        ttk.Separator(self.frame, orient='horizontal').grid(row=1, column=0, columnspan=10, sticky='ew', padx=10, pady=5)

        # 按鈕容器，避免 header 欄位影響按鈕排版
        btn_container = ttk.Frame(self.frame)
        btn_container.grid(row=2, column=0, columnspan=10, sticky='w')
        row, col = 0, 0
        # 支援多種表頭名稱（中/英）對應到 item / url / password / 埠
        def get_field(rec, candidates):
            for key in rec.keys():
                k = str(key).strip().lower()
                for c in candidates:
                    if c == k:
                        return str(rec[key]).strip()
            return ''
        for rec in self.clients:
            tag = get_field(rec, ['item', '設備名稱', 'name'])
            url = get_field(rec, ['url', 'id', 'address'])
            pwd = get_field(rec, ['password', '密碼', 'pass'])
            prt = get_field(rec, ['port', '埠', '埠號'])
            tk.Button(btn_container, text=f"{tag}\n{url}", font=('微軟正黑體',10), width=15, height=4,
                command = (lambda t=tag, u=url, p=pwd, pt=prt: self.run_tightvnc(t, u, p, pt))
            ).grid(row=row, column=col, padx=3, pady=3)
            col += 1
            if col >= 10:
                col = 0
                row += 1
    

gui = tk.Tk()
gui.title('Alldesk')

# 調整 Notebook 標籤字型：加大並改為粗體以便與 UI 一致
style = ttk.Style()
# 為了讓 tab 的背景/前景 mapping 生效，嘗試使用 'clam' 主題（較支援 element 顏色客製化）
try:
    if 'clam' in style.theme_names():
        style.theme_use('clam')
except Exception:
    pass
tab_font = tkfont.Font(family='微軟正黑體', size=11, weight='bold')
style.configure('Big.TNotebook.Tab', font=tab_font, padding=[12, 6], background='#f0f0f0', foreground='black')
# 確保 Notebook 本體與 tab 的預設背景一致
try:
    style.configure('TNotebook', background='#f0f0f0')
    style.configure('TNotebook.Tab', background='#f0f0f0')
except Exception:
    pass
# 當 tab 被選取時顯示黑底白字；未選取則為淺灰底黑字
style.map('Big.TNotebook.Tab',
    background=[('selected', 'black'), ('!selected', '#f0f0f0')],
    foreground=[('selected', 'white'), ('!selected', 'black')]
)

# 使用一個容器，將 `Notebook` 放左邊，右邊放一個 `EXCEL` 按鈕
container = ttk.Frame(gui)
container.pack(fill='both', expand=True)
notebook = ttk.Notebook(container, style='Big.TNotebook')
notebook.pack(side='left', fill='both', expand=True)

rustdesk = RustDesk(notebook)
rustdesk.set_elements_rustdesk()

anydesk = AnyDesk(notebook)
anydesk.set_elements_anydesk()

tightvnc = TightVNC(notebook)
tightvnc.set_elements_tightvnc()

# 新增一個與其他分頁相同風格的 `EXCEL` 分頁（放在最右側）
excel_frame = ttk.Frame(notebook)
notebook.add(excel_frame, text='EXCEL')

# 追蹤上一個選取的 tab id，點選 EXCEL 分頁時開啟檔案並還原為上一分頁
_last_tab = {'id': notebook.select()}

def _on_tab_changed(event):
    """Notebook tab 變更事件處理。

    若使用者切換到 'EXCEL' 分頁，則開啟 `Alldesk.xlsx`（嘗試指定
    對應的工作表），並恢復到上一個選取的分頁。
    """
    sel = event.widget.select()
    try:
        txt = event.widget.tab(sel, 'text') or ''
    except Exception:
        txt = ''
    if txt.upper() == 'EXCEL':
        # 由上一個被選取的 tab 判斷要開啟的 sheet（1-based index）
        sheet_idx = None
        try:
            last_id = _last_tab.get('id')
            if last_id:
                last_text = event.widget.tab(last_id, 'text') or ''
                lt = str(last_text).strip().lower()
                if 'rust' in lt:
                    sheet_idx = 1
                elif 'any' in lt:
                    sheet_idx = 2
                elif 'vnc' in lt or 'tight' in lt:
                    sheet_idx = 3
        except Exception:
            sheet_idx = None

        open_alldesk_excel(sheet_idx)
        try:
            # 還原到上一個 tab
            event.widget.select(_last_tab['id'])
        except Exception:
            pass
    else:
        _last_tab['id'] = sel

notebook.bind('<<NotebookTabChanged>>', _on_tab_changed)

# 將主視窗置中於螢幕
try:
    gui.update_idletasks()
    w = gui.winfo_width() or gui.winfo_reqwidth()
    h = gui.winfo_height() or gui.winfo_reqheight()
    sw = gui.winfo_screenwidth()
    sh = gui.winfo_screenheight()
    x = max((sw - w) // 2, 0)
    y = max((sh - h) // 2, 0)
    gui.geometry(f'+{x}+{y}')
except Exception:
    pass

gui.mainloop()