import pandas as pd
import json

file_path = '[CISSA] Relatório do Curso Ciber para Empreendedores.xlsx'

def clean_percent(val):
    if isinstance(val, str):
        return float(val.replace('%', '').replace(',', '.'))
    if isinstance(val, (int, float)) and val <= 1.0:
        return val * 100
    return val

try:
    df = pd.read_excel(file_path, sheet_name='Sheet1')
    
    # Process data
    df['Progresso'] = df['Progresso do estudante'].apply(clean_percent)
    
    # Force Matheus Villarim as completed if found (user correction)
    df.loc[df['Nome completo'].str.contains('Matheus Vilarim', case=False, na=False), 'Concluído'] = 'Sim'
    df.loc[df['Nome completo'].str.contains('Matheus Vilarim', case=False, na=False), 'Progresso'] = 100

    # Calculations
    total_students = len(df)
    completed_students = len(df[df['Concluído'] == 'Sim'])
    active_students = len(df[df['Status'] == 'Ativo'])
    never_accessed = len(df[df['Último acesso ao curso'] == 'Nunca'])
    never_accessed_rate = f"{(never_accessed/total_students)*100:.1f}%" if total_students > 0 else "0%"
    avg_progress = df['Progresso'].mean()
    
    # Progress distribution
    progress_dist = {
        "0%": len(df[df['Progresso'] == 0]),
        "1-50%": len(df[(df['Progresso'] > 0) & (df['Progresso'] <= 50)]),
        "51-99%": len(df[(df['Progresso'] > 50) & (df['Progresso'] < 100)]),
        "100%": len(df[df['Progresso'] >= 100])
    }
    
    # Students list for a table
    students_list = []
    for _, row in df.iterrows():
        prog_display = f"{row['Progresso']:.1f}%".replace('.', ',')
        students_list.append({
            "Nome completo com link": row['Nome completo'],
            "Progresso do estudante": prog_display,
            "Concluído": row['Concluído'],
            "Último acesso ao curso": row['Último acesso ao curso']
        })

    dashboard_data = {
        "metrics": {
            "total_students": total_students,
            "completed_students": completed_students,
            "active_students": active_students,
            "never_accessed": never_accessed,
            "never_accessed_rate": never_accessed_rate,
            "avg_progress": f"{avg_progress:.1f}%",
            "completion_rate": f"{(completed_students/total_students)*100:.1f}%"
        },
        "progress_dist": progress_dist,
        "students": students_list
    }
    
    with open('dashboard_data.json', 'w', encoding='utf-8') as f:
        json.dump(dashboard_data, f, ensure_ascii=False, indent=4)
        
    import re
    with open('index.html', 'r', encoding='utf-8') as f:
        html = f.read()
    json_str = json.dumps(dashboard_data, ensure_ascii=False, indent=4)
    new_html = re.sub(r'const dashboardData\s*=\s*\{[\s\S]*?\n\s*\};', 'const dashboardData = ' + json_str + ';\n', html)
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(new_html)

    print("Data processed and index.html updated.")

except Exception as e:
    print(f"Error: {e}")
