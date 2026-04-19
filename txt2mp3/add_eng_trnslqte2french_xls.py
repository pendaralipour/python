import pandas as pd
from deep_translator import GoogleTranslator
import time

input_file = 'cleaned_vocabulary.xlsx'
output_file = 'cleaned_vocabulary_with_eng.xlsx'

def translate_to_eng(text):
    if pd.isna(text) or str(text).strip() == "":
        return ""
    try:
        # ترجمه از فرانسوی به انگلیسی
        translation = GoogleTranslator(source='fr', target='en').translate(str(text))
        print(f"Translated: {text} -> {translation}")
        return translation
    except Exception as e:
        print(f"Error translating {text}: {e}")
        return "Error"

try:
    excel_data = pd.ExcelFile(input_file)
    with pd.ExcelWriter(output_file) as writer:
        for sheet_name in excel_data.sheet_names:
            print(f"\n--- Processing Sheet: {sheet_name} ---")
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            
            # ترجمه ستون اول (فرانسوی) و ریختن در ستون جدید eng
            df['eng'] = df.iloc[:, 0].apply(translate_to_eng)
            
            # ذخیره شیت
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            # کمی توقف برای جلوگیری از مسدود شدن توسط گوگل
            time.sleep(1)
            
    print(f"\n✅ Done! File saved as: {output_file}")

except Exception as e:
    print(f"❌ An error occurred: {e}")