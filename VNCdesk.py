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


def read_csv(path):
    items = []
    if not os.path.exists(path):
        return items
    with open(path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        # normalize potential BOM in fieldnames
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


# Base directory for resources. When bundled by PyInstaller, sys._MEIPASS points to
# the temporary extraction folder. For writable output (modified vnc file), when
# frozen use the system temp dir; otherwise use the script folder.
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
        # set base size first
        self.geometry(f'{WIN_WIDTH}x{WIN_HEIGHT}')
        self.resizable(False, False)

        # center window on screen
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = (screen_w - WIN_WIDTH) // 2
        y = (screen_h - WIN_HEIGHT) // 2
        self.geometry(f'{WIN_WIDTH}x{WIN_HEIGHT}+{x}+{y}')

        self.csv_path = csv_path

        # Top controls (first row)
        top_frame = tk.Frame(self)
        top_frame.pack(fill='x', padx=8, pady=6)

        tk.Label(top_frame, text='URL').pack(side='left')
        self.entry_url = tk.Entry(top_frame, width=40)
        self.entry_url.pack(side='left', padx=(4, 12))

        tk.Label(top_frame, text='PORT').pack(side='left')
        self.entry_port = tk.Entry(top_frame, width=8)
        self.entry_port.pack(side='left', padx=(4, 12))
        # default port
        self.entry_port.insert(0, '5900')

        tk.Label(top_frame, text='Password').pack(side='left')
        self.entry_password = tk.Entry(top_frame, width=20, show='*')
        self.entry_password.pack(side='left', padx=(4, 12))

        load_btn = tk.Button(top_frame, text='連線', command=self.connect_entry)
        load_btn.pack(side='left')

        # Canvas area for buttons + optional scrollbar
        container = tk.Frame(self)
        container.pack(fill='both', expand=True)


        # let geometry manager decide height; avoid fixed height to prevent early sizing issues
        self.canvas = tk.Canvas(container, width=WIN_WIDTH)
        self.canvas.pack(side='left', fill='both', expand=True)

        self.scrollbar = tk.Scrollbar(container, orient='vertical', command=self.canvas.yview, width=20)

        # connect scrollbar set directly; we will pack/unpack scrollbar when needed
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.buttons_frame = tk.Frame(self.canvas)
        self.window_id = self.canvas.create_window((0, 0), window=self.buttons_frame, anchor='nw')

        self.items = []
        self.buttons = []

        self.load_and_build()

        # Update scrollregion when canvas size changes
        self.canvas.bind('<Configure>', lambda e: self._on_canvas_configure())

    def _on_canvas_configure(self):
        self.canvas.itemconfig(self.window_id, width=self.canvas.winfo_width())
        self.update_scrollbar_visibility()

    def update_scrollbar_visibility(self):
        self.update_idletasks()
        content_h = self.buttons_frame.winfo_height()
        canvas_h = self.canvas.winfo_height()
        # add a larger tolerance so tiny pixel differences don't force scrollbar
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

        # Calculate button width (先計算好按鈕大小)
        total_h_spacing = (COLUMNS - 1) * SPACING
        available_w = WIN_WIDTH - 2 * MARGIN - total_h_spacing
        btn_w = available_w // COLUMNS
        btn_h = BUTTON_HEIGHT

        # Prepare texts and find longest
        texts = []
        for item, url, port, password in self.items:
            t1 = item
            t2 = f"{url}:{port}" if url or port else ''
            texts.append((t1, t2, password))

        # Determine max font size that fits into btn_w and btn_h for two lines
        chosen_size = 12
        # try from large to small
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

        # Create buttons in grid
        row = 0
        col = 0
        for idx, (t1, t2, password) in enumerate(texts):
            b = tk.Button(self.buttons_frame, text=f"{t1}\n{t2}", width=1, height=1, font=btn_font, anchor='center', justify='center')
            # use place with absolute size to ensure equal spacing
            x = MARGIN + col * (btn_w + SPACING)
            y = row * (btn_h + SPACING)
            b.place(x=x, y=y, width=btn_w, height=btn_h)
            self.buttons.append(b)
            col += 1
            if col >= COLUMNS:
                col = 0
                row += 1

        # set frame size to contain all buttons
        total_rows = (len(texts) + COLUMNS - 1) // COLUMNS
        frame_h = total_rows * btn_h + max(0, total_rows - 1) * SPACING
        frame_w = WIN_WIDTH
        self.buttons_frame.configure(width=frame_w, height=frame_h)
        self.canvas.configure(scrollregion=(0, 0, frame_w, frame_h))
        self.update_scrollbar_visibility()

        # bind mouse wheel for scrolling when content is overflow
        def _on_mousewheel(event):
            # Windows: event.delta is multiple of 120
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        self.canvas.bind_all('<MouseWheel>', _on_mousewheel)

        # now bind correct commands capturing the item values
        for i, btn in enumerate(self.buttons):
            if i < len(self.items):
                item, url, port, password = self.items[i]
                btn.configure(command=lambda u=url, p=port, pw=password: self.write_and_launch(u, p, pw))

    def write_and_launch(self, url, port, password):
        vnc_source = resource_path('vnc.vnc')
        # read original bundled file (if present)
        if os.path.exists(vnc_source):
            with open(vnc_source, 'r', encoding='utf-8') as f:
                lines = f.readlines()
        else:
            lines = []

        # find [connection] section and replace host/port/password lines
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
                    # end of connection section
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
                    # if caller provided a password, replace it; otherwise keep existing password line
                    if password:
                        enc_pw = encrypt_vnc_password(password)
                        out.append(f'password={enc_pw}\n')
                        replaced['password'] = True
                    else:
                        out.append(line)
                    continue
            out.append(line)

        # if connection section missing or some keys not replaced, reconstruct
        if not any(l.strip().lower() == '[connection]' for l in out):
            # prepend connection block; include password only if provided
            conn_block = ["[connection]\n", f"host={url}\n", f"port={port}\n"]
            if password:
                enc_pw = encrypt_vnc_password(password)
                conn_block.append(f"password={enc_pw}\n")
            out = conn_block + ['\n'] + out
        else:
            # ensure missing keys are inserted right after [connection]
            if not (replaced['host'] and replaced['port'] and replaced['password']):
                new_out = []
                i = 0
                while i < len(out):
                    new_out.append(out[i])
                    if out[i].strip().lower() == '[connection]':
                        # insert/replace following lines
                        j = i + 1
                        # consume existing until next section
                        consume = []
                        while j < len(out) and not out[j].strip().startswith('['):
                            consume.append(out[j])
                            j += 1
                        # build connection lines
                        conn_lines = [f'host={url}\n', f'port={port}\n']
                        if password:
                            enc_pw = encrypt_vnc_password(password)
                            conn_lines.append(f'password={enc_pw}\n')
                        else:
                            # preserve any existing password line from the consumed block
                            for c in consume:
                                if c.strip().lower().startswith('password='):
                                    conn_lines.append(c)
                                    break
                        new_out.extend(conn_lines)
                        i = j
                        continue
                    i += 1
                out = new_out

        # write back to a writable location (script folder when not frozen,
        # otherwise temp dir). Use that path as optionsfile for the launched exe.
        out_path = os.path.join(get_writable_dir(), 'vnc.vnc')
        with open(out_path, 'w', encoding='utf-8') as f:
            f.writelines(out)

        # launch TightVNC.exe (use bundled binary when available)
        import subprocess
        exe_path = resource_path('TightVNC.exe')
        # if binary not bundled, fall back to relative name so system PATH can find it
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
