import pandas as pd
import pyttsx3
import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import glob
import time

FFMPEG_PATH = "ffmpeg"

def save_tts_offline(text, lang, speed_factor, filename):
    try:
        engine = pyttsx3.init()
        voices = engine.getProperty('voices')
        
        # انتخاب صدا بر اساس زبان
        # برای فرانسه 'FR' و برای انگلیسی 'EN' را جستجو می‌کند
        selected_voice = None
        for voice in voices:
            if lang.upper() in voice.id.upper() or lang.upper() in voice.name.upper():
                selected_voice = voice.id
                break
        
        if selected_voice:
            engine.setProperty('voice', selected_voice)

        # تنظیم سرعت موتور داخلی (پایه)
        engine.setProperty('rate', 150) 
        
        temp_wav = f"temp_{filename}.wav"
        engine.save_to_file(text, temp_wav)
        engine.runAndWait()
        
        # استفاده از FFmpeg برای تبدیل به mp3 و اعمال سرعت دقیق کاربر
        command = [
            FFMPEG_PATH, "-y", "-i", temp_wav,
            "-filter:a", f"atempo={speed_factor}", 
            "-ar", "44100", "-ac", "1", filename
        ]
        subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
        
        if os.path.exists(temp_wav):
            os.remove(temp_wav)
        return True
    except Exception as e:
        print(f"TTS Error: {e}")
        return False

def create_silence(duration_ms, filename):
    command = [
        FFMPEG_PATH, "-y", "-f", "lavfi", "-i", f"anullsrc=r=44100:cl=mono",
        "-t", str(float(duration_ms)/1000), "-acodec", "libmp3lame", filename
    ]
    subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)

def run_process():
    threading.Thread(target=main_logic, daemon=True).start()

def main_logic():
    excel_path = file_entry.get()
    if not excel_path or not os.path.exists(excel_path):
        messagebox.showerror("خطا", "فایل اکسل پیدا نشد!")
        return

    try:
        start_btn.config(state="disabled")
        
        # خواندن ورودی‌های کاربر
        sheets = [s.strip() for s in sheet_entry.get().split(',')]
        fr_col = fr_col_entry.get()
        en_col = en_col_entry.get()
        include_en = en_var.get()
        speed = speed_entry.get()
        rep_count = int(repeat_entry.get())
        p_rep = pause_repeat_entry.get()
        p_word = pause_word_entry.get()

        output_name = "Offline_Learning.mp3"
        temp_files = []
        
        pause_repeat_file = "p_repeat.mp3"
        pause_word_file = "p_word.mp3"
        create_silence(p_rep, pause_repeat_file)
        create_silence(p_word, pause_word_file)

        excel_data = pd.ExcelFile(excel_path)
        
        # محاسبه تعداد کل برای پروگرس بار
        total = 0
        for s in sheets:
            if s in excel_data.sheet_names:
                total += len(pd.read_excel(excel_path, sheet_name=s))

        count = 0
        file_idx = 0
        for sheet in sheets:
            if sheet not in excel_data.sheet_names: continue
            df = pd.read_excel(excel_path, sheet_name=sheet)
            
            for _, row in df.iterrows():
                fr_text = str(row[fr_col]).strip()
                if fr_text.lower() == 'nan': continue

                fr_file = f"temp_{file_idx}_fr.mp3"
                save_tts_offline(fr_text, 'fr', speed, fr_file)

                if include_en:
                    en_text = str(row[en_col]).strip()
                    en_file = f"temp_{file_idx}_en.mp3"
                    save_tts_offline(en_text, 'en', speed, en_file)
                    temp_files.append(en_file)
                    temp_files.append(pause_repeat_file)

                for _ in range(rep_count):
                    temp_files.append(fr_file)
                    temp_files.append(pause_repeat_file)

                temp_files.append(pause_word_file)
                file_idx += 1
                count += 1
                progress_bar['value'] = (count / total) * 100
                root.update_idletasks()

        # ادغام نهایی
        with open("file_list.txt", "w", encoding="utf-8") as f:
            for file in temp_files:
                f.write(f"file '{os.path.abspath(file)}'\n")

        subprocess.run([FFMPEG_PATH, "-y", "-f", "concat", "-safe", "0", "-i", "file_list.txt", "-acodec", "libmp3lame", output_name])

        # پاکسازی
        time.sleep(1)
        for f in glob.glob("temp_*.mp3") + glob.glob("temp_*.wav") + ["file_list.txt", "p_repeat.mp3", "p_word.mp3"]:
            if os.path.exists(f): os.remove(f)

        messagebox.showinfo("موفقیت", f"فایل با موفقیت ساخته شد:\n{output_name}")
    except Exception as e:
        messagebox.showerror("خطا", str(e))
    finally:
        start_btn.config(state="normal")

# --- رابط کاربری ---
root = tk.Tk()
root.title("نسخه آفلاین تبدیل اکسل به صوت")
root.geometry("480x700")

tk.Label(root, text="برنامه ای برای تبدیل فایل اکسل به صوت (آفلاین)", font=("Tahoma", 12, "bold")).pack(pady=15)

file_entry = tk.Entry(root, width=40); file_entry.pack()
tk.Button(root, text="انتخاب فایل اکسل", command=lambda: file_entry.insert(0, filedialog.askopenfilename())).pack(pady=5)

tk.Label(root, text="نام شیت‌ها (مثلاً Anki,Doulingo):").pack()
sheet_entry = tk.Entry(root, width=30); sheet_entry.insert(0, "Anki,Doulingo"); sheet_entry.pack()

tk.Label(root, text="نام ستون فرانسه:").pack()
fr_col_entry = tk.Entry(root, width=15); fr_col_entry.insert(0, "fr"); fr_col_entry.pack()

tk.Label(root, text="نام ستون انگلیسی:").pack()
en_col_entry = tk.Entry(root, width=15); en_col_entry.insert(0, "eng"); en_col_entry.pack()

en_var = tk.BooleanVar(value=True)
tk.Checkbutton(root, text="شامل انگلیسی باشد؟", variable=en_var).pack()

tk.Label(root, text="سرعت (0.7 کند - 1.0 نرمال):").pack()
speed_entry = tk.Entry(root, width=10); speed_entry.insert(0, "0.7"); speed_entry.pack()

tk.Label(root, text="تعداد تکرار فرانسه:").pack()
repeat_entry = tk.Entry(root, width=10); repeat_entry.insert(0, "3"); repeat_entry.pack()

tk.Label(root, text="مکث بین تکرارها (ms):").pack()
pause_repeat_entry = tk.Entry(root, width=10); pause_repeat_entry.insert(0, "2000"); pause_repeat_entry.pack()

tk.Label(root, text="مکث بین کلمات (ms):").pack()
pause_word_entry = tk.Entry(root, width=10); pause_word_entry.insert(0, "3000"); pause_word_entry.pack()

start_btn = tk.Button(root, text="شروع تبدیل آفلاین", command=run_process, bg="blue", fg="white", height=2, width=20)
start_btn.pack(pady=20)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=350, mode="determinate")
progress_bar.pack(pady=10)

root.mainloop()