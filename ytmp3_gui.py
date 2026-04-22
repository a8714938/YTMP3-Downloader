import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
import sys
import urllib.request
import zipfile
import json
import subprocess # 用於在背景執行外部的 yt-dlp.exe

def get_app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

class YTMP3Downloader:
    def __init__(self, root):
        self.root = root
        self.root.title("YouTube 轉 MP3 (自帶更新引擎版)")
        self.root.resizable(False, False)
        self.root.config(padx=5, pady=5)
        
        self.app_dir = get_app_dir()
        self.ffmpeg_path = os.path.join(self.app_dir, "ffmpeg.exe")
        self.ytdlp_path = os.path.join(self.app_dir, "yt-dlp.exe") # 新增 yt-dlp 實體路徑
        self.config_path = os.path.join(self.app_dir, "config.json")
        
        self.save_path = self.load_config()
        self.ensure_dir(self.save_path)
        self.is_always_on_top = False

        # === 介面佈局 ===
        tk.Label(root, text="網址:").grid(row=0, column=0, padx=2, pady=5, sticky="e")
        self.url_entry = tk.Entry(root, width=42)
        self.url_entry.grid(row=0, column=1, padx=2, pady=5)
        
        tk.Button(root, text="貼上", command=self.paste_url, width=5).grid(row=0, column=2, padx=2, pady=5)
        self.pin_btn = tk.Button(root, text="📍 固定", command=self.toggle_topmost, width=6, bg="#E0E0E0")
        self.pin_btn.grid(row=0, column=3, padx=2, pady=5)

        tk.Label(root, text="儲存:").grid(row=1, column=0, padx=2, pady=2, sticky="e")
        self.path_var = tk.StringVar(value=self.save_path)
        tk.Entry(root, textvariable=self.path_var, width=42).grid(row=1, column=1, padx=2, pady=2)
        
        tk.Button(root, text="更改", command=self.browse_folder, width=5).grid(row=1, column=2, padx=2, pady=2)
        tk.Button(root, text="打開", command=self.open_folder, width=6).grid(row=1, column=3, padx=2, pady=2)
        
        # === 主要操作與更新按鈕區 ===
        btn_frame = tk.Frame(root)
        btn_frame.grid(row=2, column=0, columnspan=4, pady=8, sticky="we")
        
        self.download_btn = tk.Button(btn_frame, text="一鍵下載 MP3", command=self.start_download, bg="#4CAF50", fg="white", font=("微軟正黑體", 10, "bold"))
        self.download_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # 雙核心更新按鈕
        self.update_yt_btn = tk.Button(btn_frame, text="🔄 更新 yt-dlp", command=self.update_ytdlp, bg="#FF9800", fg="white", font=("微軟正黑體", 9))
        self.update_yt_btn.pack(side=tk.RIGHT, padx=2)

        self.env_btn = tk.Button(btn_frame, text="下載 FFmpeg", command=self.download_ffmpeg, bg="#2196F3", fg="white", font=("微軟正黑體", 9))
        self.env_btn.pack(side=tk.RIGHT, padx=2)

        self.status_label = tk.Label(root, text="準備就緒", fg="gray", font=("微軟正黑體", 9))
        self.status_label.grid(row=3, column=0, columnspan=4)

        self.check_env()

    # --- UI 與系統設定功能 ---
    def toggle_topmost(self):
        self.is_always_on_top = not self.is_always_on_top
        self.root.attributes('-topmost', self.is_always_on_top)
        if self.is_always_on_top:
            self.pin_btn.config(text="📌 取消", bg="#FFEB3B")
        else:
            self.pin_btn.config(text="📍 固定", bg="#E0E0E0")

    def load_config(self):
        default_path = r"D:\YTMP3"
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    return json.load(f).get("last_save_path", default_path)
            except: pass
        return default_path

    def save_config(self, new_path):
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump({"last_save_path": new_path}, f, ensure_ascii=False, indent=4)
        except: pass

    def ensure_dir(self, path):
        try:
            if not os.path.exists(path): os.makedirs(path, exist_ok=True)
            return True
        except:
            fallback = os.path.expanduser("~/Downloads")
            self.save_path = fallback
            if hasattr(self, 'path_var'): self.path_var.set(fallback)
            return False

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            folder = os.path.normpath(folder)
            self.path_var.set(folder)
            self.save_config(folder)
            self.ensure_dir(folder)

    def open_folder(self):
        folder = self.path_var.get()
        if os.path.exists(folder): os.startfile(folder)

    def paste_url(self):
        try:
            self.url_entry.delete(0, tk.END)
            self.url_entry.insert(0, self.root.clipboard_get())
        except: pass

    # --- 環境檢測與更新 (核心修改) ---
    def check_env(self):
        missing = []
        if not os.path.exists(self.ffmpeg_path): missing.append("FFmpeg")
        if not os.path.exists(self.ytdlp_path): missing.append("yt-dlp")
        
        if missing:
            self.status_label.config(text=f"缺少核心引擎: {' 與 '.join(missing)}，請點擊右方按鈕下載", fg="red")
            self.download_btn.config(state=tk.DISABLED)
        else:
            self.status_label.config(text="雙引擎檢測正常，可隨時下載或更新", fg="green")
            self.download_btn.config(state=tk.NORMAL)

    def update_ytdlp(self):
        self.update_yt_btn.config(state=tk.DISABLED, text="下載中...")
        threading.Thread(target=self.process_ytdlp_download, daemon=True).start()

    def process_ytdlp_download(self):
        self.root.after(0, lambda: self.status_label.config(text="正在從官方獲取最新 yt-dlp...", fg="blue"))
        # yt-dlp 官方編譯好的單一執行檔網址
        url = "https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp.exe"
        try:
            urllib.request.urlretrieve(url, self.ytdlp_path)
            self.root.after(0, lambda: self.status_label.config(text="yt-dlp 更新成功！", fg="green"))
            self.root.after(0, self.check_env)
        except Exception:
            self.root.after(0, lambda: self.status_label.config(text="yt-dlp 更新失敗，請檢查網路", fg="red"))
        finally:
            self.root.after(0, lambda: self.update_yt_btn.config(state=tk.NORMAL, text="🔄 更新 yt-dlp"))

    def download_ffmpeg(self):
        self.env_btn.config(state=tk.DISABLED, text="下載中...")
        threading.Thread(target=self.process_ffmpeg_download, daemon=True).start()

    def process_ffmpeg_download(self):
        self.root.after(0, lambda: self.status_label.config(text="正在下載 FFmpeg (檔案較大請稍候)...", fg="blue"))
        url = "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip"
        zip_path = os.path.join(self.app_dir, "ffmpeg_temp.zip")
        try:
            urllib.request.urlretrieve(url, zip_path)
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                for file_info in zip_ref.infolist():
                    if file_info.filename.endswith('ffmpeg.exe'):
                        file_info.filename = "ffmpeg.exe"
                        zip_ref.extract(file_info, self.app_dir)
            if os.path.exists(zip_path): os.remove(zip_path)
            self.root.after(0, lambda: self.status_label.config(text="FFmpeg 下載成功！", fg="green"))
            self.root.after(0, self.check_env)
        except Exception:
            self.root.after(0, lambda: self.status_label.config(text="FFmpeg 下載失敗", fg="red"))
        finally:
            self.root.after(0, lambda: self.env_btn.config(state=tk.NORMAL, text="下載 FFmpeg"))

    # --- 執行下載指令 (核心修改為 subprocess) ---
    def start_download(self):
        url = self.url_entry.get().strip()
        if not url: return messagebox.showwarning("提示", "請貼上網址")
        
        self.ensure_dir(self.path_var.get())
        self.download_btn.config(state=tk.DISABLED, text="處理中...")
        self.status_label.config(text="轉檔中，由於背景隱藏執行，請耐心等待...", fg="blue")
        threading.Thread(target=self.process_download, args=(url,), daemon=True).start()

    def process_download(self, url):
        # 組合命令列參數 (等同於之前的 ydl_opts)
        output_template = os.path.join(self.path_var.get(), '%(title)s.%(ext)s')
        cmd = [
            self.ytdlp_path,            # 呼叫外掛的 yt-dlp.exe
            '-x',                       # 提取音訊
            '--audio-format', 'mp3',    # 轉為 MP3
            '--audio-quality', '192',   # 音質 192k
            '--ffmpeg-location', self.app_dir, # 指定 ffmpeg 位置
            '--no-playlist',            # 不下載播放清單
            '-o', output_template,      # 輸出路徑
            url                         # 影片網址
        ]

        # 隱藏呼叫外部程式時產生的黑色終端機視窗 (Windows 專屬設定)
        startupinfo = None
        if os.name == 'nt':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        try:
            # 執行指令並等待完成
            process = subprocess.run(cmd, startupinfo=startupinfo, creationflags=subprocess.CREATE_NO_WINDOW)
            
            if process.returncode == 0:
                self.root.after(0, lambda: self.update_status("下載與轉檔成功！", "green"))
            else:
                self.root.after(0, lambda: self.update_status("下載失敗：官方演算法可能已更新，請嘗試點擊『更新 yt-dlp』", "red"))
        except Exception as e:
            self.root.after(0, lambda: self.update_status(f"系統錯誤：{str(e)[:20]}", "red"))

    def update_status(self, msg, color):
        self.status_label.config(text=msg, fg=color)
        self.download_btn.config(state=tk.NORMAL, text="一鍵下載 MP3")

if __name__ == "__main__":
    root = tk.Tk()
    app = YTMP3Downloader(root)
    root.mainloop()