"""
Convert Excel file to JSON for GitHub Pages deployment
"""
import pandas as pd
import json

# Read Excel file
excel_file = 'รวมข้อมูลผู้สูงอายุ_2566.xlsx'

try:
    # Read all sheet names first
    excel_data = pd.ExcelFile(excel_file)
    sheet_names = excel_data.sheet_names
    
    print(f"พบ {len(sheet_names)} sheets: {sheet_names}")
    
    # Read the LAST sheet (Old people)
    last_sheet = sheet_names[-1]
    print(f"\nกำลังอ่านจาก sheet: '{last_sheet}'")
    
    df = pd.read_excel(excel_file, sheet_name=last_sheet)
    
    # Preview data
    print(f"\nตัวอย่างข้อมูล (5 แถวแรก):")
    print(df.head())
    
    # Convert to JSON
    data = df.to_dict('records')
    
    # Save to JSON file
    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print("\n" + "="*50)
    print("SUCCESS: แปลงข้อมูลสำเร็จ!")
    print(f"ไฟล์: data.json")
    print(f"จำนวนแถว: {len(data)}")
    print(f"คอลัมน์: {list(df.columns)}")
    
except Exception as e:
    print(f"ERROR: เกิดข้อผิดพลาด: {e}")
    print("\nกรุณาติดตั้ง pandas และ openpyxl ก่อน:")
    print("pip install pandas openpyxl")

