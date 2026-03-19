import pandas as pd
import numpy as np
import sys
from datetime import datetime

def parse_time(time_val):
    if pd.isna(time_val):
        return None
    
    # Handle datetime types properly if pandas parsed them as time objects or strings
    if type(time_val) == str:
        # Try converting string to time
        try:
            return pd.to_datetime(time_val).time()
        except:
            return None
    elif hasattr(time_val, 'hour'):
        return time_val
    return None

def calc_hours(t_in, t_out):
    if t_in is None or t_out is None:
        return None
    
    # Convert to minutes since midnight
    m_in = t_in.hour * 60 + t_in.minute
    m_out = t_out.hour * 60 + t_out.minute
    
    if m_out < m_in:
        m_out += 24 * 60 # Handle overnight
        
    return (m_out - m_in) / 60.0

try:
    df = pd.read_excel('clock-in-data.xlsx')
except Exception as e:
    print(f"Error reading file: {e}")
    sys.exit(1)

cols = df.columns.tolist()

# Attempt to smartly guess columns
date_col = next((c for c in cols if 'date' in c.lower()), None)
m_in_col = next((c for c in cols if 'morning' in c.lower() and 'in' in c.lower()), None)
m_out_col = next((c for c in cols if 'morning' in c.lower() and 'out' in c.lower()), None)
a_in_col = next((c for c in cols if 'afternoon' in c.lower() and 'in' in c.lower()), None)
a_out_col = next((c for c in cols if 'afternoon' in c.lower() and 'out' in c.lower()), None)

if not all([date_col, m_in_col, m_out_col, a_in_col, a_out_col]):
    # Fallback to column index mapping if naming doesn't exactly match
    date_col = cols[0]
    m_in_col = cols[1]
    m_out_col = cols[2]
    a_in_col = cols[3]
    a_out_col = cols[4]

grand_total = 0.0

print(f"Mapping Columns:")
print(f" - Date: {date_col}")
print(f" - Morning In: {m_in_col}")
print(f" - Morning Out: {m_out_col}")
print(f" - Afternoon In: {a_in_col}")
print(f" - Afternoon Out: {a_out_col}\n")

print("-" * 80)
print(f"{'Date':<15} | {'Morning (hrs)':<13} | {'Afternoon (hrs)':<15} | {'Total (hrs)':<11} | {'Notes'}")
print("-" * 80)

for idx, row in df.iterrows():
    date_val = row[date_col]
    
    # format date
    if hasattr(date_val, 'strftime'):
        date_str = date_val.strftime('%Y-%m-%d')
    else:
        date_str = str(date_val).split(' ')[0]
        
    m_in = parse_time(row[m_in_col])
    m_out = parse_time(row[m_out_col])
    a_in = parse_time(row[a_in_col])
    a_out = parse_time(row[a_out_col])
    
    m_hours = calc_hours(m_in, m_out)
    a_hours = calc_hours(a_in, a_out)
    
    notes = []
    
    if pd.notna(row[m_in_col]) and m_in and not m_out:
        notes.append("Missing Morning Out")
    if pd.notna(row[m_out_col]) and not m_in and m_out:
        notes.append("Missing Morning In")
    if pd.notna(row[a_in_col]) and a_in and not a_out:
        notes.append("Missing Afternoon Out")
    if pd.notna(row[a_out_col]) and not a_in and a_out:
        notes.append("Missing Afternoon In")
        
    m_h_val = m_hours if m_hours is not None else 0.0
    a_h_val = a_hours if a_hours is not None else 0.0
    
    total_day = m_h_val + a_h_val
    grand_total += total_day
    
    notes_str = ", ".join(notes) if notes else ""
    if hasattr(pd, "isna") and pd.isna(row[m_in_col]) and pd.isna(row[m_out_col]) and pd.isna(row[a_in_col]) and pd.isna(row[a_out_col]):
        notes_str = "No shifts logged"
    
    m_display = f"{m_h_val:.2f}" if m_hours is not None else "N/A"
    a_display = f"{a_h_val:.2f}" if a_hours is not None else "N/A"
    t_display = f"{total_day:.2f}"
    
    print(f"{date_str:<15} | {m_display:<13} | {a_display:<15} | {t_display:<11} | {notes_str}")

print("-" * 80)
print(f"GRAND TOTAL: {grand_total:.2f} hours")
print("-" * 80)
