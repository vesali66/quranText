#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Quran Data Collection Script
Collects Arabic text (with tashkeel and normalized), Persian translations, and Tafsir
Outputs to Excel file
"""

import requests
import openpyxl
import re
import json
from openpyxl import Workbook

# API Endpoints
QURAN_API_BASE = "https://api.quran.com"
ALQURAN_API_BASE = "https://api.alquran.cloud/v1"

def get_quran_verses():
    """Fetch all Quran verses with Arabic text (with and without tashkeel)"""
    print("Fetching Arabic text...")
    
    # Get Uthmani text (with tashkeel)
    uthmani_url = f"{ALQURAN_API_BASE}/quran/uthmani"
    # Get Imlaei simple text (normalized without tashkeel)
    imlaei_url = f"{ALQURAN_API_BASE}/quran/imlaei.simple"
    
    uthmani_response = requests.get(uthmani_url, timeout=60)
    imlaei_response = requests.get(imlaei_url, timeout=60)
    
    if uthmani_response.status_code != 200 or imlaei_response.status_code != 200:
        raise Exception("Failed to fetch Arabic text")
    
    uthmani_data = uthmani_response.json()
    imlaei_data = imlaei_response.json()
    
    verses = {}
    
    for surah in uthmani_data['data']['surahs']:
        surah_num = surah['number']
        for ayah in surah['ayahs']:
            ayah_num = ayah['numberInSurah']
            key = f"{surah_num}:{ayah_num}"
            verses[key] = {
                'uthmani': ayah['text'],
                'imlaei': '',
                'surah': surah_num,
                'ayah': ayah_num
            }
    
    for surah in imlaei_data['data']['surahs']:
        surah_num = surah['number']
        for ayah in surah['ayahs']:
            ayah_num = ayah['numberInSurah']
            key = f"{surah_num}:{ayah_num}"
            if key in verses:
                verses[key]['imlaei'] = ayah['text']
    
    return verses

def get_persian_translations():
    """Fetch multiple Persian translations"""
    print("Fetching Persian translations...")
    
    # Available Persian translation identifiers from alquran.cloud
    persian_translations = [
        'fa.ayati',           # آیتی - AbdolMohammad Ayati
        'fa.fooladvand',      # فولادوند - Mohammad Mahdi Fooladvand
        'fa.ghomshei',        # الهی قمشه‌ای - Mahdi Elahi Ghomshei
        'fa.makarem',         # مکارم شیرازی - Naser Makarem Shirazi
        'fa.ansarian',        # انصاریان - Hussain Ansarian
        'fa.bahrampour',      # بهرام پور - Abolfazl Bahrampour
        'fa.khorramshahi',    # خرمشاهی - Baha'oddin Khorramshahi
        'fa.mojtabavi',       # مجتبوی - Sayyed Jalaloddin Mojtabavi
        'fa.khorramdel',      # خرمدل - Mostafa Khorramdel
        'fa.moezzi',          # معزی - Mohammad Kazem Moezzi
        'fa.gharaati',        # قرائتی - Mohsen Gharaati
        'fa.sadeqi',          # صادقی تهرانی - Mohammad Sadeqi Tehrani
        'fa.safavi',          # صفوی - Mohammad Reza Safavi
    ]
    
    translations = {}
    
    for identifier in persian_translations:
        name = identifier.replace('fa.', '')
        print(f"  Fetching {name}...")
        url = f"{ALQURAN_API_BASE}/quran/{identifier}"
        
        try:
            response = requests.get(url, timeout=60)
            
            if response.status_code == 200:
                data = response.json()
                if 'data' in data and 'surahs' in data['data']:
                    for surah in data['data']['surahs']:
                        surah_num = surah['number']
                        for ayah in surah['ayahs']:
                            ayah_num = ayah['numberInSurah']
                            key = f"{surah_num}:{ayah_num}"
                            
                            if key not in translations:
                                translations[key] = {}
                            
                            translations[key][name] = ayah['text']
            else:
                print(f"  Warning: Could not fetch {name} (status {response.status_code})")
        except Exception as e:
            print(f"  Warning: Error fetching {name}: {e}")
    
    return translations, persian_translations

def get_tafsirs():
    """Fetch multiple Tafsirs (commentaries)"""
    print("Fetching Tafsirs...")
    
    # Available Tafsirs from alquran.cloud
    tafsir_identifiers = [
        'ar.jalalayn',     # تفسير الجلالين - Jalalayn
        'ar.qurtubi',      # تفسير القرطبي - Al-Qurtubi
        'ar.baghawi',      # تفسير البغوي - Al-Baghawi
        'ar.waseet',       # التفسير الوسيط - Al-Waseet
        'ar.muyassar',     # تفسير المیسر - Al-Muyassar
    ]
    
    tafsirs = {}
    
    for identifier in tafsir_identifiers:
        name = identifier.replace('ar.', '')
        print(f"  Fetching {name}...")
        url = f"{ALQURAN_API_BASE}/quran/{identifier}"
        
        try:
            response = requests.get(url, timeout=60)
            
            if response.status_code == 200:
                data = response.json()
                if 'data' in data and 'surahs' in data['data']:
                    for surah in data['data']['surahs']:
                        surah_num = surah['number']
                        for ayah in surah['ayahs']:
                            ayah_num = ayah['numberInSurah']
                            key = f"{surah_num}:{ayah_num}"
                            
                            if key not in tafsirs:
                                tafsirs[key] = {}
                            
                            tafsirs[key][name] = ayah['text']
            else:
                print(f"  Warning: Could not fetch {name} (status {response.status_code})")
        except Exception as e:
            print(f"  Warning: Error fetching {name}: {e}")
    
    return tafsirs, tafsir_identifiers

def normalize_arabic(text):
    """Normalize Arabic text for search purposes"""
    if not text:
        return ""
    
    # Remove tashkeel (diacritics)
    arabic_diacritics = '\u064B\u064C\u064D\u064E\u064F\u0650\u0651\u0652\u0670\u0653\u0654\u0655\u0656\u0657\u0658\u0659\u065A\u065B\u065C\u065D\u065E\u065F'
    normalized = ''.join(char for char in text if char not in arabic_diacritics)
    
    # Normalize Alef variations
    normalized = normalized.replace('آ', 'ا').replace('أ', 'ا').replace('إ', 'ا')
    
    # Normalize Teh Marbuta
    normalized = normalized.replace('ة', 'ه')
    
    # Normalize Yeh variations
    normalized = normalized.replace('ى', 'ي')
    
    # Remove extra spaces
    normalized = ' '.join(normalized.split())
    
    return normalized

def create_excel_file(output_file='quran_data.xlsx'):
    """Create Excel file with all collected data"""
    print("\nStarting data collection...")
    
    # Collect all data
    arabic_verses = get_quran_verses()
    persian_translations, translation_names = get_persian_translations()
    tafsirs, tafsir_names = get_tafsirs()
    
    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Quran Data"
    
    # Create headers
    headers = [
        'Surah',
        'Ayah',
        'Verse Key',
        'Arabic (Uthmani)',
        'Arabic (Normalized)',
        'Arabic (Imlaei Simple)'
    ]
    
    # Add translation headers
    for trans_name in translation_names:
        clean_name = trans_name.replace('fa.', '')
        headers.append(f'Translation_{clean_name}')
    
    # Add tafsir headers
    for tafsir_name in tafsir_names:
        clean_name = tafsir_name.replace('ar.', '')
        headers.append(f'Tafsir_{clean_name}')
    
    # Write headers
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
    
    # Write data
    row = 2
    total_verses = 6236  # Total verses in Quran
    
    print(f"\nWriting {total_verses} verses to Excel...")
    
    for surah in range(1, 115):
        # Get number of verses in this surah (approximate, will be filled)
        for ayah in range(1, 200):  # Max verses in a surah is 286
            verse_key = f"{surah}:{ayah}"
            
            if verse_key not in arabic_verses:
                continue
            
            verse_data = arabic_verses[verse_key]
            
            # Arabic texts
            uthmani = verse_data.get('uthmani', '')
            imlaei = verse_data.get('imlaei', '')
            normalized = normalize_arabic(uthmani)
            
            # Write basic data
            ws.cell(row=row, column=1, value=surah)
            ws.cell(row=row, column=2, value=ayah)
            ws.cell(row=row, column=3, value=verse_key)
            ws.cell(row=row, column=4, value=uthmani)
            ws.cell(row=row, column=5, value=normalized)
            ws.cell(row=row, column=6, value=imlaei)
            
            # Write translations
            col_offset = 6
            translations_data = persian_translations.get(verse_key, {})
            for i, trans_name in enumerate(translation_names, 1):
                trans_text = translations_data.get(trans_name, '')
                ws.cell(row=row, column=col_offset + i, value=trans_text)
            
            # Write tafsirs (if available for this verse)
            tafsir_col_start = col_offset + len(translation_names)
            tafsirs_data = tafsirs.get(verse_key, {})
            for i, tafsir_name in enumerate(tafsir_names, 1):
                tafsir_text = tafsirs_data.get(tafsir_name.replace('ar.', ''), '')
                ws.cell(row=row, column=tafsir_col_start + i, value=tafsir_text)
            
            row += 1
            
            if row % 500 == 0:
                print(f"  Processed {row-2} verses...")
    
    # Save workbook
    print(f"\nSaving to {output_file}...")
    wb.save(output_file)
    print(f"Done! File saved as {output_file}")
    print(f"Total verses: {row-2}")
    print(f"Translations included: {len(translation_names)}")
    print(f"Tafsirs included: {len(tafsir_names)}")
    
    return output_file

if __name__ == "__main__":
    create_excel_file()
