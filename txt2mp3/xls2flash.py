"""
aplication for make audio flash cart from excel file
poblished by PENDAR
2026-4
"""

import pandas as pd
from gtts import gTTS
import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import glob
import time

FFMPEG_PATH = "ffmpeg"

def save_tts(text, lang, speed_factor, filename):
    try:
        tts = gTTS(text=text, lang=lang, slow=True)
        temp_raw = f"raw_{filename}"
        tts.save(temp_raw)
        command = [
            FFMPEG_PATH, "-y", "-i", temp_raw,
            "-filter:a", f"atempo={speed_factor}", 
            "-ar", "44100", "-ac", "1", filename
        ]
        subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
        if os.path.exists(temp_raw): os.remove(temp_raw)
    except Exception as e:
        print(f"Error in TTS: {e}")

def create_silence(duration_ms, filename):
    command = [
        FFMPEG_PATH, "-y", "-f", "lavfi", "-i", f"anullsrc=r=44100:cl=mono",
        "-t", str(float(duration_ms)/1000), "-acodec", "libmp3lame", filename
    ]
    subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)

def run_conversion_thread():
    threading.Thread(target=process_logic, daemon=True).start()

def process_logic():
    excel_path = file_entry.get()
    if not excel_path:
        messagebox.showwarning("خطا", "لطفاً فایل اکسل را انتخاب کنید")
        return

    try:
        start_btn.config(state="disabled")
        progress_bar['value'] = 0
        
        sheets = [s.strip() for s in sheet_entry.get().split(',')]
        fr_col = fr_col_entry.get()
        en_col = en_col_entry.get()
        include_en = en_var.get()
        speed = speed_entry.get()
        p_repeat = pause_repeat_entry.get()
        p_word = pause_word_entry.get()
        repeat_count = int(repeat_entry.get())

        output_name = "Learning_Audio.mp3"
        temp_files = []
        
        pause_repeat_file = "p_repeat.mp3"
        pause_word_file = "p_word.mp3"
        create_silence(p_repeat, pause_repeat_file)
        create_silence(p_word, pause_word_file)

        excel_data = pd.ExcelFile(excel_path)
        
        total_rows = 0
        for sheet in sheets:
            if sheet in excel_data.sheet_names:
                total_rows += len(pd.read_excel(excel_path, sheet_name=sheet))
        
        current_step = 0
        file_idx = 0

        for sheet in sheets:
            if sheet not in excel_data.sheet_names: continue
            df = pd.read_excel(excel_path, sheet_name=sheet)
            
            for _, row in df.iterrows():
                fr_text = str(row[fr_col]).strip()
                if fr_text.lower() != 'nan':
                    fr_file = f"temp_{file_idx}_fr.mp3"
                    save_tts(fr_text, 'fr', speed, fr_file)

                    if include_en:
                        en_text = str(row[en_col]).strip()
                        en_file = f"temp_{file_idx}_en.mp3"
                        save_tts(en_text, 'en', speed, en_file)
                        temp_files.append(en_file)
                        temp_files.append(pause_repeat_file)

                    for _ in range(repeat_count):
                        temp_files.append(fr_file)
                        temp_files.append(pause_repeat_file)

                    temp_files.append(pause_word_file)
                    file_idx += 1
                
                current_step += 1
                progress_percent = (current_step / total_rows) * 100
                progress_bar['value'] = progress_percent
                root.update_idletasks()

        if temp_files:
            with open("file_list.txt", "w", encoding="utf-8") as f:
                for file in temp_files:
                    f.write(f"file '{os.path.abspath(file)}'\n")

            final_cmd = [FFMPEG_PATH, "-y", "-f", "concat", "-safe", "0", "-i", "file_list.txt", "-acodec", "libmp3lame", output_name]
            subprocess.run(final_cmd)

        time.sleep(1)
        for f in glob.glob("temp_*.mp3") + ["file_list.txt", "p_repeat.mp3", "p_word.mp3"]:
            try:
                if os.path.exists(f): os.remove(f)
            except: pass

        messagebox.showinfo("موفقیت", "فایل با موفقیت ساخته شد.")
    except Exception as e:
        messagebox.showerror("خطا", f"خطایی رخ داد: {e}")
    finally:
        start_btn.config(state="normal")
        progress_bar['value'] = 0

# --- تنظیمات رابط کاربری ---
root = tk.Tk()
root.title("Excel to MP3 Professional")
root.geometry("480x700")

# --- اضافه شدن عنوان در بالاترین قسمت ---
main_title = tk.Label(root, text="برنامه ای برای تبدیل فایل اکسل به صوت", font=("Tahoma", 14, "bold"), fg="#2c3e50")
main_title.pack(pady=20)

# سایر المان‌ها
tk.Label(root, text="فایل اکسل:").pack(pady=2)
file_entry = tk.Entry(root, width=50)
file_entry.pack()
tk.Button(root, text="انتخاب فایل", command=lambda: file_entry.insert(0, filedialog.askopenfilename())).pack(pady=5)

tk.Label(root, text="نام شیت‌ها (با کاما جدا کنید):").pack()
sheet_entry = tk.Entry(root, width=30); sheet_entry.insert(0, "Anki,Doulingo")
sheet_entry.pack()

tk.Label(root, text="نام ستون فرانسه:").pack()
fr_col_entry = tk.Entry(root, width=20); fr_col_entry.insert(0, "fr")
fr_col_entry.pack()

tk.Label(root, text="نام ستون انگلیسی:").pack()
en_col_entry = tk.Entry(root, width=20); en_col_entry.insert(0, "eng")
en_col_entry.pack()

en_var = tk.BooleanVar(value=True)
tk.Checkbutton(root, text="شامل زبان انگلیسی باشد؟", variable=en_var).pack()

tk.Label(root, text="سرعت خواندن (مثلاً 0.7):").pack()
speed_entry = tk.Entry(root, width=10); speed_entry.insert(0, "0.7")
speed_entry.pack()

tk.Label(root, text="تعداد تکرار فرانسه:").pack()
repeat_entry = tk.Entry(root, width=10); repeat_entry.insert(0, "3")
repeat_entry.pack()

tk.Label(root, text="مکث بین تکرارها (ms):").pack()
pause_repeat_entry = tk.Entry(root, width=10); pause_repeat_entry.insert(0, "2000")
pause_repeat_entry.pack()

tk.Label(root, text="مکث بین کلمات (ms):").pack()
pause_word_entry = tk.Entry(root, width=10); pause_word_entry.insert(0, "3000")
pause_word_entry.pack()

start_btn = tk.Button(root, text="شروع تبدیل", command=run_conversion_thread, bg="#2ecc71", fg="white", height=2, width=20)
start_btn.pack(pady=20)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=350, mode="determinate")
progress_bar.pack(pady=10)

root.mainloop()