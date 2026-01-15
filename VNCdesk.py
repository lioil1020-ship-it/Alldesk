import os
import sys
import tempfile
from pathlib import Path

try:
    from Crypto.Cipher import DES
    _HAS_CRYPTO = True
except Exception:
    _HAS_CRYPTO = False


def resource_path(relative_path: str) -> str:
    """回傳資源檔的實際路徑（支援 PyInstaller 打包）。"""
    if getattr(sys, 'frozen', False):
        base = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    else:
        base = os.path.dirname(__file__)
    return os.path.join(base, relative_path)


def get_writable_dir() -> str:
    """回傳一個可寫入的目錄（打包時回傳 exe 所在目錄，開發時回傳原始碼目錄）。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(__file__)


def encrypt_vnc_password(password: str) -> str:
    """簡單的 VNC 密碼處理。

    注意：原程式會將此結果寫入 `vnc.vnc` 作為 password 欄位內容。若系統需要
    與 TightVNC 相容的 DES 加密，請安裝 `pycryptodome` 並替換此實作以符合目標 viewer。

    此處實作：
    - 若安裝了 Crypto，會用 DES (ECB) 與固定 key 做簡單加密，並回傳 hex；
    - 否則回傳原始密碼（最多取前 8 字元）。
    """
    if not password:
        return ''
    pw = str(password)[:8]
    if not _HAS_CRYPTO:
        return pw
    # 注意：這不是正式 TightVNC 格式的保證實作，但在有 crypto 時提供一個加密結果。
    key = b'VeryKey!'
    cipher = DES.new(key, DES.MODE_ECB)
    data = pw.encode('utf-8')
    if len(data) < 8:
        data = data.ljust(8, b'\x00')
    else:
        data = data[:8]
    enc = cipher.encrypt(data)
    return enc.hex()
import csv
import os
import sys
import tempfile
import tkinter as tk
from tkinter import font, messagebox
from Crypto.Cipher import DES


WIN_WIDTH = 1200
WIN_HEIGHT = 600
MARGIN = 10
SPACING = 5
COLUMNS = 6
BUTTON_HEIGHT = 80


def encrypt_vnc_password(password):
    # 取前 8 個 ASCII 字元，不足則以 NUL 填充
    pw = (password or '')[:8].encode('ascii', errors='ignore')
    pw = pw.ljust(8, b'\x00')

    # TightVNC 使用固定的 challenge bytes；對每個位元反轉以生成 DES 金鑰
    challenge = [23, 82, 107, 6, 35, 78, 88, 7]

    def rev_bits_byte(b):
        b = ((b & 0xF0) >> 4) | ((b & 0x0F) << 4)
        b = ((b & 0xCC) >> 2) | ((b & 0x33) << 2)
        b = ((b & 0xAA) >> 1) | ((b & 0x55) << 1)
        return b

    key = bytes([rev_bits_byte(b) for b in challenge])
    cipher = DES.new(key, DES.MODE_ECB)
    return cipher.encrypt(pw).hex()


def read_csv(path):
    items = []
    if not os.path.exists(path):
        return items
    with open(path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        # 正規化欄名可能的 BOM
        rows = []
        fieldnames = reader.fieldnames
        if fieldnames:
            normalized = [fn.lstrip('\ufeff').strip() for fn in fieldnames]
            if normalized != fieldnames:
                for r in reader:
                    newr = {}
                    for k, v in r.items():
                        nk = k.lstrip('\ufeff')
                        newr[nk] = v
                    rows.append(newr)
            else:
                rows = list(reader)
        for r in rows:
            item = (r.get('Item') or r.get('\ufeffItem') or '').strip()
            url = (r.get('URL') or '').strip()
            port = (r.get('port') or r.get('Port') or '').strip()
            password = (r.get('password') or r.get('Password') or '').strip()
            if item:
                items.append((item, url, port, password))
    return items


# 資源基底目錄。當由 PyInstaller 打包時，sys._MEIPASS 指向暫時解壓資料夾。
# 對於可寫輸出（修改後的 vnc 檔），若 frozen 則使用系統暫存目錄；否則使用腳本資料夾。
BASE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(__file__))

def resource_path(filename):
    return os.path.join(BASE_DIR, filename)

def get_writable_dir():
    if getattr(sys, 'frozen', False):
        return tempfile.gettempdir()
    return os.path.dirname(__file__)





class App(tk.Tk):
    def __init__(self, csv_path):
        super().__init__()
        self.title('VNC by lioil')
        # 先設定基礎視窗大小
        self.geometry(f'{WIN_WIDTH}x{WIN_HEIGHT}')
        self.resizable(False, False)

        # 將視窗置中顯示
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = (screen_w - WIN_WIDTH) // 2
        y = (screen_h - WIN_HEIGHT) // 2
        self.geometry(f'{WIN_WIDTH}x{WIN_HEIGHT}+{x}+{y}')

        self.csv_path = csv_path

        # 頂部控制列（第一列）
        top_frame = tk.Frame(self)
        top_frame.pack(fill='x', padx=8, pady=6)

        tk.Label(top_frame, text='URL').pack(side='left')
        self.entry_url = tk.Entry(top_frame, width=40)
        self.entry_url.pack(side='left', padx=(4, 12))

        tk.Label(top_frame, text='PORT').pack(side='left')
        self.entry_port = tk.Entry(top_frame, width=8)
        self.entry_port.pack(side='left', padx=(4, 12))
        # 預設埠號
        self.entry_port.insert(0, '5900')

        tk.Label(top_frame, text='Password').pack(side='left')
        self.entry_password = tk.Entry(top_frame, width=20, show='*')
        self.entry_password.pack(side='left', padx=(4, 12))

        load_btn = tk.Button(top_frame, text='連線', command=self.connect_entry)
        load_btn.pack(side='left')

        # 按鈕區域（使用 Canvas）及可選捲軸
        container = tk.Frame(self)
        container.pack(fill='both', expand=True)


        # 讓 geometry manager 決定高度；避免使用固定高度以免產生大小問題
        self.canvas = tk.Canvas(container, width=WIN_WIDTH)
        self.canvas.pack(side='left', fill='both', expand=True)

        self.scrollbar = tk.Scrollbar(container, orient='vertical', command=self.canvas.yview, width=20)

        # 直接設定 scrollbar 的回呼；視需要再 pack/unpack 捲軸
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.buttons_frame = tk.Frame(self.canvas)
        self.window_id = self.canvas.create_window((0, 0), window=self.buttons_frame, anchor='nw')

        self.items = []
        self.buttons = []

        self.load_and_build()

        # 當 Canvas 大小變化時更新 scrollregion
        self.canvas.bind('<Configure>', lambda e: self._on_canvas_configure())

    def _on_canvas_configure(self):
        self.canvas.itemconfig(self.window_id, width=self.canvas.winfo_width())
        self.update_scrollbar_visibility()

    def update_scrollbar_visibility(self):
        self.update_idletasks()
        content_h = self.buttons_frame.winfo_height()
        canvas_h = self.canvas.winfo_height()
        # 增加容差，以避免微小像素差異導致捲軸顯示
        tolerance = 20
        if content_h > canvas_h + tolerance:
            if not self.scrollbar.winfo_ismapped():
                self.scrollbar.pack(side='right', fill='y')
        else:
            if self.scrollbar.winfo_ismapped():
                self.scrollbar.pack_forget()


    def reload(self):
        self.load_and_build()

    def connect_entry(self):
        url = self.entry_url.get().strip()
        port = self.entry_port.get().strip()
        # read password from entry (empty means keep existing password in vnc.vnc)
        password = self.entry_password.get()
        self.write_and_launch(url, port, password)

    def load_and_build(self):
        self.items = read_csv(self.csv_path)
        
        for w in self.buttons_frame.winfo_children():
            w.destroy()
        self.buttons.clear()

        # 計算按鈕寬度（先計算好按鈕大小）
        total_h_spacing = (COLUMNS - 1) * SPACING
        available_w = WIN_WIDTH - 2 * MARGIN - total_h_spacing
        btn_w = available_w // COLUMNS
        btn_h = BUTTON_HEIGHT

        # 準備按鈕文字並找出最長
        texts = []
        for item, url, port, password in self.items:
            t1 = item
            t2 = f"{url}:{port}" if url or port else ''
            texts.append((t1, t2, password))

        # 決定可在兩行內放入 btn_w 與 btn_h 的最大字型大小
        chosen_size = 12
        # 從大到小嘗試
        for size in range(30, 5, -1):
            f = font.Font(size=size)
            max_w = 0
            max_h = 0
            for t1, t2, _ in texts:
                w1 = f.measure(t1)
                w2 = f.measure(t2)
                max_w = max(max_w, w1, w2)
                max_h = max(max_h, f.metrics('linespace'))
            if max_w <= btn_w - 8 and (max_h * 2) <= btn_h - 8:
                chosen_size = size
                break

        btn_font = font.Font(size=chosen_size)

        # 在格線中建立按鈕
        row = 0
        col = 0
        for idx, (t1, t2, password) in enumerate(texts):
            b = tk.Button(self.buttons_frame, text=f"{t1}\n{t2}", width=1, height=1, font=btn_font, anchor='center', justify='center')
            # 使用 place 設定絕對尺寸以確保等間距
            x = MARGIN + col * (btn_w + SPACING)
            y = row * (btn_h + SPACING)
            b.place(x=x, y=y, width=btn_w, height=btn_h)
            self.buttons.append(b)
            col += 1
            if col >= COLUMNS:
                col = 0
                row += 1

        # 設定 frame 大小以包含所有按鈕
        total_rows = (len(texts) + COLUMNS - 1) // COLUMNS
        frame_h = total_rows * btn_h + max(0, total_rows - 1) * SPACING
        frame_w = WIN_WIDTH
        self.buttons_frame.configure(width=frame_w, height=frame_h)
        self.canvas.configure(scrollregion=(0, 0, frame_w, frame_h))
        self.update_scrollbar_visibility()

        # 綁定滑鼠滾輪以在內容溢出時捲動
        def _on_mousewheel(event):
            # Windows: event.delta 通常為 120 的倍數
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        self.canvas.bind_all('<MouseWheel>', _on_mousewheel)

        # 現在綁定正確的 command，使用預設參數捕獲項目值
        for i, btn in enumerate(self.buttons):
            if i < len(self.items):
                item, url, port, password = self.items[i]
                btn.configure(command=lambda u=url, p=port, pw=password: self.write_and_launch(u, p, pw))

    def write_and_launch(self, url, port, password):
        vnc_source = resource_path('vnc.vnc')
        # 讀取原始綁定的 vnc 檔（若存在）
        if os.path.exists(vnc_source):
            with open(vnc_source, 'r', encoding='utf-8') as f:
                lines = f.readlines()
        else:
            lines = []

        # 找出 [connection] 區段並替換 host/port/password 行
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
                    # connection 區段結束
                    in_conn = False
                    out.append(line)
                    continue
                if s.lower().startswith('host='):
                    out.append(f'host={url}\n')
                    replaced['host'] = True
                    continue
                if s.lower().startswith('port='):
                    out.append(f'port={port}\n')
                    replaced['port'] = True
                    continue
                if s.lower().startswith('password='):
                    # 若呼叫者提供密碼則替換；否則保留原有密碼行
                    if password:
                        enc_pw = encrypt_vnc_password(password)
                        out.append(f'password={enc_pw}\n')
                        replaced['password'] = True
                    else:
                        out.append(line)
                    continue
            out.append(line)

        # 若缺少 connection 區段或某些鍵未被替換，重建該區段
        if not any(l.strip().lower() == '[connection]' for l in out):
            # 在開頭插入 connection 區塊；只有提供密碼時才包含 password
            conn_block = ["[connection]\n", f"host={url}\n", f"port={port}\n"]
            if password:
                enc_pw = encrypt_vnc_password(password)
                conn_block.append(f"password={enc_pw}\n")
            out = conn_block + ['\n'] + out
        else:
            # 確保缺少的鍵在 [connection] 後插入
            if not (replaced['host'] and replaced['port'] and replaced['password']):
                new_out = []
                i = 0
                while i < len(out):
                    new_out.append(out[i])
                    if out[i].strip().lower() == '[connection]':
                        # 插入/替換後續行
                        j = i + 1
                        # consume existing until next section
                        consume = []
                        while j < len(out) and not out[j].strip().startswith('['):
                            consume.append(out[j])
                            j += 1
                        # 建立 connection 行
                        conn_lines = [f'host={url}\n', f'port={port}\n']
                        if password:
                            enc_pw = encrypt_vnc_password(password)
                            conn_lines.append(f'password={enc_pw}\n')
                        else:
                            # 從已消耗的區塊保留現有的 password 行（若有）
                            for c in consume:
                                if c.strip().lower().startswith('password='):
                                    conn_lines.append(c)
                                    break
                        new_out.extend(conn_lines)
                        i = j
                        continue
                    i += 1
                out = new_out

        # 寫回可寫入的位置（非 frozen 時為腳本資料夾，否則為暫存目錄）。
        # 使用該路徑作為啟動參數的 optionsfile。
        out_path = os.path.join(get_writable_dir(), 'vnc.vnc')
        with open(out_path, 'w', encoding='utf-8') as f:
            f.writelines(out)

        # 啟動 TightVNC.exe（若有綁定則使用該二進位檔）
        import subprocess
        exe_path = resource_path('TightVNC.exe')
        # 如果未綁定二進位檔，回退為相對名稱以讓系統 PATH 搜尋
        if not os.path.exists(exe_path):
            exe_path = 'TightVNC.exe'

        args = [exe_path, f'-optionsfile={out_path}', '-showcontrols=no']
        try:
            subprocess.Popen(args, cwd=get_writable_dir())
            
        except Exception as e:
            messagebox.showerror('啟動失敗', f'無法啟動 {exe_path}: {e}')


if __name__ == '__main__':
    if getattr(sys, 'frozen', False):
        here = os.path.dirname(sys.executable)
    else:
        here = os.path.dirname(__file__)
    csv_path = os.path.join(here, 'vnc.csv')
    try:
        app = App(csv_path)
        app.mainloop()
    except Exception as e:
        messagebox.showerror('Unhandled exception', str(e))
        raise
