import pandas as pd
import json
import os

file_path = '[CISSA] Relatório do Curso Ciber para Empreendedores.xlsx'

try:
    # Read Excel file
    xl = pd.ExcelFile(file_path)
    print(f"Sheets: {xl.sheet_names}")
    
    # Analyze each sheet
    data_summary = {}
    for sheet in xl.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet)
        data_summary[sheet] = {
            "columns": df.columns.tolist(),
            "rows": len(df),
            "sample": df.head(5).to_dict(orient='records')
        }
    
    # Save analysis to a JSON for later use
    with open('analysis_summary.json', 'w', encoding='utf-8') as f:
        json.dump(data_summary, f, ensure_ascii=False, indent=4)
        
    print("Analysis complete. Saved to analysis_summary.json")

except Exception as e:
    print(f"Error: {e}")
