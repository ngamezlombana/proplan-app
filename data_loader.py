import pandas as pd
import io

def load_excel_data(file_bytes):
    """
    Lee las hojas principales del archivo Excel de ProPlan y extrae los DataFrames limpios.
    """
    try:
        # Usamos pandas para leer el archivo desde la memoria
        xls = pd.ExcelFile(io.BytesIO(file_bytes))
        
        # Cargar hojas
        df_prog = pd.read_excel(xls, sheet_name='PROGRAMADOR')
        df_maq = pd.read_excel(xls, sheet_name='MAQUINAS')
        df_master = pd.read_excel(xls, sheet_name='MAESTRO')

        # Limpiar datos básicos
        df_prog = clean_programador(df_prog)
        machines, blocked_dates = parse_maquinas(df_maq)
        color_map = parse_maestro(df_master)

        return df_prog, machines, blocked_dates, color_map

    except Exception as e:
        print(f"Error cargando Excel: {e}")
        return None, None, None, None

def clean_programador(df):
    """
    Ordena y limpia la hoja del programador (elimina filas vacías, ordena por prioridad).
    """
    # Eliminar filas donde no hay OP
    df = df.dropna(subset=[df.columns[0]])
    
    # Ordenar por PRIORIDAD (asumimos que está en la columna H (índice 7)) y luego OP (índice 0)
    priority_col = df.columns[7]
    op_col = df.columns[0]
    
    # Forzar numérico para prioridad
    df[priority_col] = pd.to_numeric(df[priority_col], errors='coerce').fillna(999)
    df = df.sort_values(by=[priority_col, op_col], ascending=[True, True])
    
    return df

def parse_maquinas(df):
    """
    Extrae la lista de máquinas y su mapa de fechas bloqueadas.
    """
    machines = []
    blocked_dates = {}

    if df.empty:
        return machines, blocked_dates

    # Las fechas están desde la columna 2 en adelante en la cabecera
    # En pandas, df.columns[2:] serían las fechas
    date_headers = df.columns[2:]

    for idx, row in df.iterrows():
        group = str(row.iloc[0]).strip()
        name = str(row.iloc[1]).strip()
        
        if pd.isna(name) or name == 'nan' or name == '':
            continue
            
        machines.append({'group': group, 'name': name, 'priority': idx})
        
        # Buscar fechas bloqueadas ('X')
        blocked_set = set()
        for col_idx, date_val in enumerate(date_headers):
            cell_val = str(row.iloc[col_idx + 2]).strip().upper()
            if cell_val == 'X':
                # Intentamos guardar la fecha como string YYYY-MM-DD
                try:
                    if isinstance(date_val, pd.Timestamp):
                        blocked_set.add(date_val.strftime('%Y-%m-%d'))
                    else:
                        blocked_set.add(str(date_val))
                except:
                    pass
                    
        blocked_dates[name] = blocked_set

    return machines, blocked_dates

def parse_maestro(df):
    """
    Extrae el mapa de colores (Grupo -> Color hex)
    """
    color_map = {}
    for _, row in df.iterrows():
        group = str(row.iloc[0]).strip()
        color = str(row.iloc[1]).strip()
        if group and color and color != 'nan':
            color_map[group] = color
    return color_map
