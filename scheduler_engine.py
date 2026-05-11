from datetime import datetime, timedelta
import re

def clean_str(val):
    if pd_isna(val): return ''
    s = str(val)
    # Reemplazar todos los espacios extraños por un espacio normal y luego hacer strip
    s = re.sub(r'\s+', ' ', s)
    return s.strip()


def run_scheduler(df_prog, machines, blocked_dates, start_date=None):
    """
    Ejecuta el algoritmo de programación (Gantt) basado en las reglas del sistema.
    """
    if start_date is None:
        start_date = datetime.now().date()
        
    gantt_state = {m['name']: set() for m in machines}
    scheduled_tasks = []
    
    # Mapeo de duración fija (Requerimiento especial de versiones anteriores)
    hardcoded_durations = {
        ('RAY1', '9'): 2,
        ('RAY2', '9'): 2,
        ('RAY3', '9'): 2,
        ('RAY1', '10'): 3,
        ('RAY2', '10'): 3,
        ('RAY3', '10'): 3
    }
    
    # Procesar fila por fila (Ya viene ordenado por prioridad de data_loader)
    for idx, row in df_prog.iterrows():
        op_name = str(row.iloc[0]).strip()
        op_details = {
            'op': clean_str(op_name),
            'desc': clean_str(row.iloc[1]),
            'cliente': clean_str(row.iloc[2]),
            'zona': clean_str(row.iloc[3]),
            'eje': clean_str(row.iloc[4]),
            'tipo': clean_str(row.iloc[5]),
            'prioridad': clean_str(row.iloc[7]),
            'gr1_maquina': clean_str(row.iloc[8]) if len(row) > 8 else '',
            'fecha_fin': '9999-12-31' # Default, will be updated
        }
        
        current_stage_start = start_date
        
        # Guardar la OP en el resultado base
        op_schedule_entry = {
            'details': op_details,
            'tasks': []
        }
        
        # Etapas asumiendo que empiezan en la columna 8 (índice 8) y van de 2 en 2
        # Ejemplo: ETAPA1, DIAS1, ETAPA2, DIAS2...
        col_idx = 8
        while col_idx < len(row) - 1:
            group = str(row.iloc[col_idx]).strip()
            days = row.iloc[col_idx + 1]
            
            if pd_isna(group) or group == 'nan' or group == '' or pd_isna(days):
                col_idx += 2
                continue
                
            try:
                days = int(float(days))
            except:
                col_idx += 2
                continue
                
            if days <= 0:
                col_idx += 2
                continue
            
            success, machine_name, dates, end_date = assign_resource(
                group, days, current_stage_start, machines, blocked_dates, gantt_state, hardcoded_durations
            )
            
            if success:
                # Registrar ocupación global
                for d in dates:
                    gantt_state[machine_name].add(d.strftime('%Y-%m-%d'))
                
                # Guardar resultado para el frontend
                op_schedule_entry['tasks'].append({
                    'group': group,
                    'machine': machine_name,
                    'start': dates[0].strftime('%Y-%m-%d'),
                    'end': dates[-1].strftime('%Y-%m-%d'),
                    'days': len(dates),
                    'all_dates': [d.strftime('%Y-%m-%d') for d in dates]
                })
                
                # La siguiente etapa empieza al día siguiente del fin de esta
                current_stage_start = end_date + timedelta(days=1)
            else:
                # Si falla (no hay máquinas en ese grupo o algo), empujamos 1 día
                current_stage_start += timedelta(days=1)
                
            col_idx += 2
            
        # Calcular fecha fin de la OP
        if op_schedule_entry['tasks']:
            max_date = max(t['end'] for t in op_schedule_entry['tasks'])
            op_schedule_entry['details']['fecha_fin'] = max_date
            
        scheduled_tasks.append(op_schedule_entry)
            
    # Ordenar por fecha fin, del más antiguo al más nuevo
    scheduled_tasks.sort(key=lambda x: x['details']['fecha_fin'])
            
    return scheduled_tasks, gantt_state

def assign_resource(group, requested_days, start_date, machines, blocked_dates, gantt_state, hardcoded_durations):
    # Filtrar candidatas
    candidates = [m for m in machines if m['group'] == group]
    # Ordenar por prioridad original en la hoja
    candidates.sort(key=lambda x: x['priority'])
    
    if not candidates:
        return False, None, [], None
        
    search_date = start_date
    safety_counter = 0
    
    while safety_counter < 365:
        for machine in candidates:
            # Chequear cap de rendimiento (Máquinas 9 y 10 son más rápidas)
            days_needed = requested_days
            
            # Limitar a máximo 2 días para la máquina 9
            if machine['name'] == '9' and days_needed > 2:
                days_needed = 2
                
            # Limitar a máximo 3 días para la máquina 10
            elif machine['name'] == '10' and days_needed > 3:
                days_needed = 3
                
            temp_date = search_date
            days_found = 0
            schedule_dates = []
            machine_valid = True
            look_ahead_safety = 0
            
            while days_found < days_needed and look_ahead_safety < 100:
                date_str = temp_date.strftime('%Y-%m-%d')
                
                is_master_blocked = date_str in blocked_dates.get(machine['name'], set())
                is_gantt_occupied = date_str in gantt_state[machine['name']]
                
                if is_gantt_occupied:
                    machine_valid = False
                    break
                    
                if not is_master_blocked:
                    schedule_dates.append(temp_date)
                    days_found += 1
                    
                temp_date += timedelta(days=1)
                look_ahead_safety += 1
                
            if machine_valid and days_found == days_needed:
                return True, machine['name'], schedule_dates, schedule_dates[-1]
                
        search_date += timedelta(days=1)
        safety_counter += 1
        
    return False, None, [], None

def pd_isna(val):
    import pandas as pd
    return pd.isna(val)
