"""
making audio flashcart from excel file
"""
import pandas as pd
from gtts import gTTS
import os
import subprocess

FFMPEG_PATH = "ffmpeg"

def save_tts(text, lang, filename):
    tts = gTTS(text=text, lang=lang, slow=True)
    temp_file = "temp_raw.mp3"
    tts.save(temp_file)

    # تغییر سرعت و استانداردسازی خروجی صوتی
    command = [
        FFMPEG_PATH, "-y", "-i", temp_file,
        "-filter:a", "atempo=1", 
        "-ar", "44100", "-ac", "1", # استانداردسازی نرخ نمونه‌برداری و کانال
        filename
    ]
    subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
    
    if os.path.exists(temp_file):
        os.remove(temp_file)

def create_silence(duration_ms, filename):
    # ساخت سکوت استاندارد
    command = [
        FFMPEG_PATH, "-y", "-f", "lavfi", "-i", f"anullsrc=r=44100:cl=mono",
        "-t", str(duration_ms/1000), "-acodec", "libmp3lame", filename
    ]
    subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)

def text_to_speech_logic(file_path, sheets, output_name):
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' not found!")
        return

    temp_files = []
    file_index = 0

    pause_2s = "pause_1s.mp3"
    pause_3s = "pause_1_5s.mp3"
    create_silence(1000, pause_2s)
    create_silence(1500, pause_3s)

    excel_data = pd.ExcelFile(file_path)

    for sheet in sheets:
        if sheet not in excel_data.sheet_names:
            continue

        print(f"Processing sheet: {sheet}")
        df = pd.read_excel(file_path, sheet_name=sheet)

        for _, row in df.iterrows():
            try:
                eng = str(row["eng"]).strip()
                fr = str(row["fr"]).strip()

                if 'nan' in (eng.lower(), fr.lower()):
                    continue

                eng_file = f"temp_{file_index}_eng.mp3"
                fr_file = f"temp_{file_index}_fr.mp3"

                save_tts(eng, 'en', eng_file)
                save_tts(fr, 'fr', fr_file)

                # یک بار انگلیسی
                temp_files.append(eng_file)
                temp_files.append(pause_2s)

                # سه بار فرانسوی
                for _ in range(4):
                    temp_files.append(fr_file)
                    temp_files.append(pause_2s)

                temp_files.append(pause_3s)
                file_index += 1

            except Exception as e:
                print(f"Error: {e}")

    # حل مشکل خطا با استفاده از متد concat صریح
    if temp_files:
        with open("file_list.txt", "w", encoding="utf-8") as f:
            for file in temp_files:
                f.write(f"file '{os.path.abspath(file)}'\n")

        # حذف -c copy و استفاده از بازنویسی جریان صوتی برای رفع خطای DTS
        final_command = [
            FFMPEG_PATH, "-y", "-f", "concat", "-safe", "0", "-i", "file_list.txt",
            "-acodec", "libmp3lame", "-q:a", "2", output_name
        ]
        print("Merging files... this might take a moment.")
        subprocess.run(final_command)

    # پاکسازی
    files_to_clean = set(temp_files)
    files_to_clean.update(["file_list.txt", pause_2s, pause_3s])
    for f in files_to_clean:
        if os.path.exists(f):
            os.remove(f)

    print(f"Done! Saved as: {output_name}")

if __name__ == "__main__":
    MY_FILE_PATH = 'cleaned_vocabulary.xlsx'
    MY_SHEETS = ['Anki', 'Doulingo', 'podcast']
    MY_OUTPUT_NAME = 'French_English_Learning.mp3'

    text_to_speech_logic(MY_FILE_PATH, MY_SHEETS, MY_OUTPUT_NAME)