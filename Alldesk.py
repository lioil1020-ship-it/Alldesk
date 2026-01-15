import tkinter as tk
import shutil, subprocess, threading
import os, stat, glob, platform
import pandas as pd
from pathlib import Path
from tkinter import font as tkfont
from tkinter import ttk
from VNCdesk import encrypt_vnc_password, resource_path, get_writable_dir

"""Alldesk GUI 啟動器

提供兩個分頁：RustDesk 與 AnyDesk。程式會從 `Alldesk.xlsx` 中載入
client 列表（sheet 名稱分別為 'rustdesk' 與 'anydesk'），並在啟動
遠端連線前，於使用者的 %AppData% 下準備必要的設定檔 (RustDesk2.toml / user.conf)。
所有功能以只讀 Excel 為來源，已移除 CSV 回退邏輯。
"""

# 預設值（可用環境變數覆寫）
# 將可執行檔統一放到專案內的 `exe` 資料夾（相對於此檔案），使用環境變數可覆寫
BASE_DIR = Path(__file__).resolve().parent
EXE_DIR = BASE_DIR / 'exe'
# rustdesk 可執行檔路徑（相對或絕對）
RUSTDESK_APP = os.getenv('RUSTDESK_APP', str(EXE_DIR / 'RustDesk.exe'))
# 用於產生 RustDesk2.toml 的 rendezvous server 與 key（固定參數）
RUSTDESK_HOST = 'everdura.ddnsfree.com'
RUSTDESK_KEY = 'kCC8dq5x8uvEI+fpbIsTpYhCMaMbAxpYmGv6XtR7NsY='

# AnyDesk / TightVNC 可執行檔路徑
ANYDESK_APP = os.getenv('ANYDESK_APP', str(EXE_DIR / 'AnyDesk.exe'))
TIGHTVNC_APP = os.getenv('TIGHTVNC_APP', str(EXE_DIR / 'TightVNC.exe'))

class RustDesk():
    """RustDesk 分頁：從 Excel 載入 client 並發起 RustDesk 連線。

    主要職責：
    - 從 `Alldesk.xlsx` 的 'rustdesk' 工作表讀取客戶清單。
    - 在啟動連線前於 %AppData%/RustDesk/config 產生或覆寫 `RustDesk2.toml`，
      並在 peers 資料夾中寫入 `<ID>.toml`，以調整視窗/預載密碼等設定。
    """
    def __init__(self, notebook: ttk.Notebook):
        self.init_rustdesk(notebook)

    def init_rustdesk(self, notebook: ttk.Notebook):
        app = RUSTDESK_APP
        # 讀取 `Alldesk.xlsx` 的 'rustdesk' 工作表；若不存在則不載入 client（僅支援 Excel）
        excel_path = Path('./Alldesk.xlsx')
        # 使用模組級常數 `RUSTDESK_HOST` / `RUSTDESK_KEY`，不在物件上存放副本
        clients = []
        if excel_path.exists():
            try:
                try:
                    df = pd.read_excel(excel_path, sheet_name='rustdesk', engine='openpyxl')
                except Exception:
                    # 若找不到命名 sheet，回退到第 0 張
                    df = pd.read_excel(excel_path, sheet_name=0, engine='openpyxl')
                rows = df.astype(str).fillna('').values.tolist()
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
        excel_path = Path('./Alldesk.xlsx')
        if excel_path.exists():
            try:
                df0 = pd.read_excel(excel_path, sheet_name=0, engine='openpyxl').astype(str).fillna('')
                if not df0.empty:
                    cols = [c.lower() for c in df0.columns]
                    id_idx = None
                    pwd_idx = None
                    for i, c in enumerate(cols):
                        if id_idx is None and 'id' in c:
                            id_idx = i
                        if pwd_idx is None and ('pass' in c or 'password' in c):
                            pwd_idx = i
                    if id_idx is None and df0.shape[1] >= 2:
                        id_idx = 1
                    if pwd_idx is None and df0.shape[1] >= 3:
                        pwd_idx = 2
                    if id_idx is None:
                        id_idx = 0
                    if pwd_idx is None:
                        pwd_idx = id_idx + 1 if df0.shape[1] > id_idx + 1 else id_idx
                    sheet_id = str(df0.iat[0, id_idx]).strip()
                    sheet_pwd = str(df0.iat[0, pwd_idx]).strip()
                    if sheet_id:
                        client_id = sheet_id
                    if sheet_pwd:
                        password = sheet_pwd
            except Exception:
                pass

        peer_file = os.path.join(peers_dir, f"{client_id}.toml")
        # 若檔案存在，確保可寫（移除唯讀屬性）
        try:
            if os.path.exists(peer_file):
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
            "wm_RemoteDesktop = '{" + '"width":1300.0,"height":740.0,"offsetWidth":1270.0,"offsetHeight":710.0,"isMaximized":true,"isFullscreen":true' + "}'\n\n"
            "[info]\n"
            f"username = '{uname}'\n"
            f"hostname = '{host}'\n"
            "platform = 'Windows'\n"
        )
        try:
            with open(peer_file, 'w', encoding='utf-8') as fw:
                fw.write(peer_content)
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
        except Exception:
            pass

  

    def run_rustdesk(self, client_id, password):
        """啟動 RustDesk 連線（RustDesk 專用）。

        步驟：
        1. 呼叫 `_prepare_rustdesk_conf`，在 %AppData% 準備必要設定檔。
        2. 以參數列表方式啟動 `rustdesk.exe --connect <id> --password <pwd>`。
        """
        exec_target = self.exec_target
        command = [exec_target, '--connect', client_id, '--password', password]
        # 在啟動 RustDesk 前，先寫入 RustDesk2.toml 以設定 view_style
        self._prepare_rustdesk_conf(client_id, password)
        # 呼叫可執行檔（使用參數列表較安全）
        subprocess.run(command)
        # 可改為非同步啟動：subprocess.Popen(command, creationflags=subprocess.CREATE_NO_WINDOW)


    def set_elements_rustdesk(self):
        """建立 RustDesk 分頁的 UI 元件。

        元件包括：
        - 輸入欄位：連接 ID、密碼
        - 連接按鈕：手動輸入後啟動連線
        - 以 Excel 讀取到的 client 列表動態建立按鈕（快速連線）
        """
        tk.Label(self.frame, text="連接ID:").grid(row=0, column=0, columnspan=2, padx=10, sticky='w')
        ent_id = tk.Entry(self.frame, width=28)
        ent_id.grid(row=0, column=0, columnspan=2, padx= 10, sticky='e')
        tk.Label(self.frame, text="密碼:").grid(row=0, column=2, sticky='w', columnspan=2, padx=10)
        ent_pass = tk.Entry(self.frame, show='*', width=30)
        ent_pass.grid(row=0, column=2, sticky='e', columnspan=2, padx=10)
        tk.Button(self.frame, text="連接", command=lambda:self.run_rustdesk(ent_id.get(), ent_pass.get())).grid(row=0, column=4, sticky='w', padx=10)
        ttk.Separator(self.frame, orient='horizontal').grid(row=1, column=0, columnspan=10, sticky='ew', padx=10, pady=5)

        row, col = 2, 0
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
            tk.Button(self.frame, text=f"{tag}\n{client_id}", font=('微軟正黑體',10), width=15, height=4, 
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
        self.init_anydesk(notebook)

    def init_anydesk(self, notebook: ttk.Notebook):
        app: str = ANYDESK_APP
        # 讀取 `Alldesk.xlsx` 的 'anydesk' 工作表；若不存在則不載入 client（僅支援 Excel）
        excel_path = Path('./Alldesk.xlsx')
        clients = []
        if excel_path.exists():
            try:
                try:
                    df = pd.read_excel(excel_path, sheet_name='anydesk', engine='openpyxl')
                except Exception:
                    # 若找不到命名 sheet，回退到第 1 張
                    df = pd.read_excel(excel_path, sheet_name=1, engine='openpyxl')
                rows = df.astype(str).fillna('').values.tolist()
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
        tk.Label(self.frame, text="連接ID:").grid(row=0, column=0, columnspan=2, padx=10, sticky='w')
        ent_id = tk.Entry(self.frame, width=28)
        ent_id.grid(row=0, column=0, columnspan=2, padx= 10, sticky='e')
        tk.Label(self.frame, text="密碼:").grid(row=0, column=2, sticky='w', columnspan=2, padx=10)
        ent_pass = tk.Entry(self.frame, show='*', width=30)
        ent_pass.grid(row=0, column=2, sticky='e', columnspan=2, padx=10)
        tk.Button(self.frame, text="連接", command=lambda:self.run_anydesk(ent_id.get(), ent_pass.get())).grid(row=0, column=4, sticky='w', padx=10)
        ttk.Separator(self.frame, orient='horizontal').grid(row=1, column=0, columnspan=10, sticky='ew', padx=10, pady=5)

        row, col = 2, 0
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
            tk.Button(self.frame, text=f"{tag}\n{client_id}", font=('微軟正黑體',10), width=15, height=4,
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
        app = 'vnc'
        excel_path = Path('./Alldesk.xlsx')
        clients = []
        if excel_path.exists():
            try:
                # 讀取第3張工作表（index=2）並轉為字典列（保留欄名以供對應）
                df = pd.read_excel(excel_path, sheet_name=2, engine='openpyxl').astype(str).fillna('')
                clients = df.to_dict(orient='records')
            except Exception:
                clients = []
        else:
            clients = []

        self.exec_target = TIGHTVNC_APP
        self.clients = clients
        self.frame = ttk.Frame(notebook)
        notebook.add(self.frame, text = 'TightVNC')

    def _prepare_and_launch_tightvnc(self, host: str, port: str, password: str):
        # 以 VNCdesk 的資源路徑與工具產生可寫入的 vnc.vnc 並啟動 TightVNC
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
                        enc_pw = encrypt_vnc_password(password)
                        out.append(f'password={enc_pw}\n')
                        replaced['password'] = True
                    else:
                        out.append(line)
                    continue
            out.append(line)

        if not any(l.strip().lower() == '[connection]' for l in out):
            conn_block = ["[connection]\n", f"host={host}\n", f"port={port}\n"]
            if password:
                enc_pw = encrypt_vnc_password(password)
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
                                        enc_pw = encrypt_vnc_password(password)
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

        out_path = os.path.join(get_writable_dir(), 'vnc.vnc')
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
        # 呼叫核心邏輯
        host = url or ''
        prt = port or '5900'
        self._prepare_and_launch_tightvnc(host, prt, password)

    def set_elements_tightvnc(self):
        # 不顯示 Item 的頂部輸入欄；將 連接ID、密碼、連接按鈕與其他分頁對齊，
        # 並將埠放在最右側（預設為 5900）
        tk.Label(self.frame, text="連接ID:").grid(row=0, column=0, columnspan=2, padx=10, sticky='w')
        ent_url = tk.Entry(self.frame, width=28)
        ent_url.grid(row=0, column=0, columnspan=2, padx=10, sticky='e')

        tk.Label(self.frame, text="密碼:").grid(row=0, column=2, columnspan=2, padx=10, sticky='w')
        ent_pass = tk.Entry(self.frame, show='*', width=30)
        ent_pass.grid(row=0, column=2, columnspan=2, padx=10, sticky='e')

        # 連接按鈕放在與其他分頁相同的位置
        tk.Button(self.frame, text="連接", command=lambda: self.run_tightvnc('', ent_url.get(), ent_pass.get(), ent_port.get())).grid(row=0, column=4, sticky='w', padx=10)

        # 埠 放在最右側並預設為 5900
        tk.Label(self.frame, text="埠:").grid(row=0, column=5, sticky='w', padx=6)
        ent_port = tk.Entry(self.frame, width=8)
        ent_port.grid(row=0, column=6, sticky='e', padx=6)
        ent_port.insert(0, '5900')
        ttk.Separator(self.frame, orient='horizontal').grid(row=1, column=0, columnspan=10, sticky='ew', padx=10, pady=5)

        row, col = 2, 0
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
            tk.Button(self.frame, text=f"{tag}\n{url}", font=('微軟正黑體',10), width=15, height=4,
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
tab_font = tkfont.Font(family='微軟正黑體', size=11, weight='bold')
style.configure('Big.TNotebook.Tab', font=tab_font, padding=[12, 6])
notebook = ttk.Notebook(gui, style='Big.TNotebook')
notebook.pack(fill = 'both', expand = True)

rustdesk = RustDesk(notebook)
rustdesk.set_elements_rustdesk()

anydesk = AnyDesk(notebook)
anydesk.set_elements_anydesk()

tightvnc = TightVNC(notebook)
tightvnc.set_elements_tightvnc()

gui.mainloop()