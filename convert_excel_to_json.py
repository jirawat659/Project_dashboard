"""
Convert Excel file to JSON for GitHub Pages deployment
"""
import pandas as pd
import json

# Read Excel file
excel_file = 'รวมข้อมูลผู้สูงอายุ_2566.xlsx'

try:
    # Read the Excel file
    df = pd.read_excel(excel_file)
    
    # Convert to JSON
    data = df.to_dict('records')
    
    # Save to JSON file
    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print("SUCCESS: แปลงข้อมูลสำเร็จ!")
    print(f"ไฟล์: data.json")
    print(f"จำนวนแถว: {len(data)}")
    print(f"คอลัมน์: {list(df.columns)}")
    
except Exception as e:
    print(f"ERROR: เกิดข้อผิดพลาด: {e}")
    print("\nกรุณาติดตั้ง pandas และ openpyxl ก่อน:")
    print("pip install pandas openpyxl")
