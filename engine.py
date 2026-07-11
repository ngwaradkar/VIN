# -*- coding: utf-8 -*-
import openpyxl
import io
import os

def parse_html_wip_vcs(file_like_or_path):
    """
    Parses HTML-based .xls files dynamically using BeautifulSoup
    to extract VEHICLE CODE and PLATE PUNCH DATE from rows.
    """
    from bs4 import BeautifulSoup
    import pandas as pd
    
    # Read content
    if hasattr(file_like_or_path, 'read'):
        content = file_like_or_path.read()
        if hasattr(file_like_or_path, 'seek'):
            file_like_or_path.seek(0)
    else:
        with open(file_like_or_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
            
    soup = BeautifulSoup(content, 'html.parser')
    rows = []
    for tr in soup.find_all('tr'):
        cells = [td.get_text(strip=True) for td in tr.find_all(['td', 'th'])]
        if cells:
            rows.append(cells)
            
    if len(rows) < 2:
        return []
        
    headers = rows[0]
    data = rows[1:]
    df = pd.DataFrame(data, columns=headers)
    
    vc_col = None
    date_col = None
    for col in df.columns:
        col_clean = str(col).strip().upper()
        if 'VEHICLE CODE' in col_clean or 'VC' in col_clean:
            vc_col = col
        if 'PLATE PUNCH DATE' in col_clean or 'PUNCH' in col_clean:
            date_col = col
            
    if vc_col is None:
        vc_col = df.columns[4] if len(df.columns) > 4 else None
    if date_col is None:
        date_col = df.columns[8] if len(df.columns) > 8 else None
        
    # Detect description and color columns
    desc_col = None
    color_col = None
    for col in df.columns:
        col_clean = str(col).strip().upper()
        if 'SALES DESCRIPTION' in col_clean or 'DESC' in col_clean:
            desc_col = col
        if 'COLOR' in col_clean or 'COLOUR' in col_clean:
            color_col = col
            
    records = []
    for idx, row in df.iterrows():
        vc_val = row[vc_col] if vc_col is not None else None
        date_val = row[date_col] if date_col is not None else None
        desc_val = row[desc_col] if desc_col is not None else ''
        color_val = row[color_col] if color_col is not None else ''
        
        vc_str = str(vc_val).strip().upper() if vc_val is not None else ''
        desc_str = str(desc_val).strip().upper()
        color_str = str(color_val).strip().upper()
        
        # Apply Pune plant exclusions
        if vc_str.startswith('5442') or 'ALTROZ' in desc_str or color_str == '000' or '000' in color_str:
            continue
            
        if vc_val is not None and str(vc_val).strip() != '':
            records.append({
                'vc': vc_str,
                'date_str': str(date_val).strip() if date_val is not None else ''
            })
            
    return records

def initialize_tcf1_workbook(file_path_or_stream):
    """
    Initializes the TCF-1 workbook for the new day's planning run:
    1. Copies yesterday's Day 1 date and values (col H) to Today Plan date and values (col D) for rows 27-81.
    2. Clears 'biw Next 2 days' columns B and C.
    3. Clears 'Stage wise Float' columns B, D, E, F, G, H from row 2 downwards.
    4. Clears 'Day 1 Sequence' columns B, C, D from row 3 downwards.
    5. Clears 'Vin Plan Check' column F (Expected to Drop) from row 4 downwards.
    """
    # 1. Handle dual load (data-only for values, formula-only for writing)
    if hasattr(file_path_or_stream, 'read'):
        content = file_path_or_stream.read()
        file_path_or_stream.seek(0)
        wb_val = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
        wb = openpyxl.load_workbook(io.BytesIO(content), data_only=False)
    else:
        wb_val = openpyxl.load_workbook(file_path_or_stream, data_only=True)
        wb = openpyxl.load_workbook(file_path_or_stream, data_only=False)
        
    # Copy H (Day 1) values to D (Today Plan) in Plan Summary
    if 'Plan Summary' in wb.sheetnames and 'Plan Summary' in wb_val.sheetnames:
        ws_val = wb_val['Plan Summary']
        ws = wb['Plan Summary']
        
        # Copy yesterday's Day 1 date (col H) to Today's Today Plan date (col D)
        yesterday_day1_date = ws_val.cell(row=2, column=8).value
        ws.cell(row=2, column=4).value = yesterday_day1_date
        
        # Calculate next 4 planning days skipping Sundays
        import datetime
        start_date = yesterday_day1_date
        if isinstance(start_date, datetime.datetime):
            start_date = start_date.date()
            
        def get_next_planning_days_t1(s_date, count=4):
            days = []
            curr = s_date
            while len(days) < count:
                curr += datetime.timedelta(days=1)
                if curr.weekday() == 6: # Sunday
                    continue
                days.append(curr)
            return days
            
        if start_date is not None and not isinstance(start_date, str):
            planning_days = get_next_planning_days_t1(start_date, 4)
            ws.cell(row=2, column=8).value = planning_days[0]
            ws.cell(row=2, column=9).value = planning_days[1]
            ws.cell(row=2, column=10).value = planning_days[2]
            ws.cell(row=2, column=11).value = planning_days[3]
        
        for row in range(27, 82):
            # Read evaluated value from col H (8)
            day1_val = ws_val.cell(row=row, column=8).value
            # Write evaluated value to col D (4)
            ws.cell(row=row, column=4).value = day1_val
            
    # 2. biw Next 2 days: Clear B, C & D
    if 'biw Next 2 days' in wb.sheetnames:
        ws_biw = wb['biw Next 2 days']
        for r in range(2, ws_biw.max_row + 1):
            ws_biw.cell(row=r, column=2).value = None
            ws_biw.cell(row=r, column=3).value = None
            ws_biw.cell(row=r, column=4).value = None
            
    # 3. Stage wise Float: Clear B, D, E, F, G, H
    if 'Stage wise Float' in wb.sheetnames:
        ws_float = wb['Stage wise Float']
        cols_to_clear = [2, 4, 5, 6, 7, 8]  # B, D, E, F, G, H
        for r in range(2, ws_float.max_row + 1):
            for col in cols_to_clear:
                ws_float.cell(row=r, column=col).value = None
                
    # 4. Day 1 Sequence: Clear B, C, D
    if 'Day 1 Sequence' in wb.sheetnames:
        ws_seq = wb['Day 1 Sequence']
        cols_seq = [2, 3, 4]  # B, C, D
        for r in range(3, ws_seq.max_row + 1):
            for col in cols_seq:
                ws_seq.cell(row=r, column=col).value = None
                
    # 5. Vin Plan Check: Clear F (Expected to Drop)
    if 'Vin Plan Check' in wb.sheetnames:
        ws_check = wb['Vin Plan Check']
        for r in range(4, ws_check.max_row + 1):
            ws_check.cell(row=r, column=6).value = None
            
    return wb

def initialize_tcf2_workbook(file_path_or_stream):
    """
    Initializes the TCF-2 workbook for the new day's planning run:
    1. Copies yesterday's Day 1 date and values (col I) to Days Plan date and values (col E) for rows 4-206.
    2. Clears 'biw Next 2 days' columns B and C.
    3. Clears 'Stage wise float' columns B, D, E, F, G, H from row 2 downwards.
    4. Clears 'Day 1 Sequence' columns B, C, D from row 3 downwards.
    5. Clears 'Vin plan check' column G (Expected Drop) from row 3 downwards.
    """
    # 1. Handle dual load (data-only for values, formula-only for writing)
    if hasattr(file_path_or_stream, 'read'):
        content = file_path_or_stream.read()
        file_path_or_stream.seek(0)
        wb_val = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
        wb = openpyxl.load_workbook(io.BytesIO(content), data_only=False)
    else:
        wb_val = openpyxl.load_workbook(file_path_or_stream, data_only=True)
        wb = openpyxl.load_workbook(file_path_or_stream, data_only=False)
        
    # Copy I (Day 1) values to E (Days Plan) in Plan Summary
    if 'Plan Summary' in wb.sheetnames and 'Plan Summary' in wb_val.sheetnames:
        ws_val = wb_val['Plan Summary']
        ws = wb['Plan Summary']
        
        # Copy yesterday's Day 1 date (col I) to Today's Today Plan date (col E)
        yesterday_day1_date = ws_val.cell(row=2, column=9).value
        ws.cell(row=2, column=5).value = yesterday_day1_date
        
        # Calculate next 3 planning days skipping Sundays
        import datetime
        start_date = yesterday_day1_date
        if isinstance(start_date, datetime.datetime):
            start_date = start_date.date()
            
        def get_next_planning_days_t2(s_date, count=3):
            days = []
            curr = s_date
            while len(days) < count:
                curr += datetime.timedelta(days=1)
                if curr.weekday() == 6: # Sunday
                    continue
                days.append(curr)
            return days
            
        if start_date is not None and not isinstance(start_date, str):
            planning_days = get_next_planning_days_t2(start_date, 3)
            ws.cell(row=2, column=9).value = planning_days[0]
            ws.cell(row=2, column=10).value = planning_days[1]
            ws.cell(row=2, column=11).value = planning_days[2]
        
        for row in range(4, 207):
            # Read evaluated value from col I (9)
            day1_val = ws_val.cell(row=row, column=9).value
            # Write evaluated value to col E (5)
            ws.cell(row=row, column=5).value = day1_val
            
    # 2. biw Next 2 days: Clear B & C
    if 'biw Next 2 days' in wb.sheetnames:
        ws_biw = wb['biw Next 2 days']
        for r in range(2, ws_biw.max_row + 1):
            ws_biw.cell(row=r, column=2).value = None
            ws_biw.cell(row=r, column=3).value = None
            
    # 3. Stage wise float: Clear B, D, E, F, G, H
    sheet_name = 'Stage wise float'
    if 'Stage wise float ' in wb.sheetnames:
        sheet_name = 'Stage wise float '
    if sheet_name in wb.sheetnames:
        ws_float = wb[sheet_name]
        cols_to_clear = [2, 4, 5, 6, 7, 8]  # B, D, E, F, G, H
        for r in range(2, ws_float.max_row + 1):
            for col in cols_to_clear:
                ws_float.cell(row=r, column=col).value = None
                
    # 4. Day 1 Sequence: Clear B, C, D
    if 'Day 1 Sequence' in wb.sheetnames:
        ws_seq = wb['Day 1 Sequence']
        cols_seq = [2, 3, 4]  # B, C, D
        for r in range(3, ws_seq.max_row + 1):
            for col in cols_seq:
                ws_seq.cell(row=r, column=col).value = None
                
    # 5. Vin plan check: Clear G (Expected Drop)
    sheet_name_check = 'Vin plan check'
    if 'Vin Plan Check' in wb.sheetnames:
        sheet_name_check = 'Vin Plan Check'
    if sheet_name_check in wb.sheetnames:
        ws_check = wb[sheet_name_check]
        for r in range(3, ws_check.max_row + 1):
            ws_check.cell(row=r, column=7).value = None
            
    return wb


def update_today_vin_generation(wb, vin_list_file_or_stream, track):
    """
    Parses the Today's VIN List file and writes the VINs to Column B of
    the 'Stage wise Float' (TCF-1) or 'Stage wise float' (TCF-2) sheet.
    """
    import pandas as pd
    import io
    
    # 1. Read the VIN list file using pandas
    try:
        # Check if it's a file stream or a file path
        if hasattr(vin_list_file_or_stream, 'read'):
            # It's a stream (e.g. UploadedFile)
            content = vin_list_file_or_stream.read()
            # Reset seek position just in case
            vin_list_file_or_stream.seek(0)
            
            # Check for HTML signature (for .xls files exported as HTML tables)
            prefix = content[:300].lower()
            if b'<table' in prefix or b'<style' in prefix or b'<html' in prefix:
                dfs = pd.read_html(io.BytesIO(content), header=0)
                df = dfs[0]
            else:
                # Determine format based on name attribute
                filename = getattr(vin_list_file_or_stream, 'name', '').lower()
                if filename.endswith('.xlsb'):
                    df = pd.read_excel(io.BytesIO(content), engine='pyxlsb')
                elif filename.endswith(('.xlsx', '.xls')):
                    try:
                        df = pd.read_excel(io.BytesIO(content))
                    except Exception:
                        try:
                            dfs = pd.read_html(io.BytesIO(content), header=0)
                            df = dfs[0]
                        except Exception as e_inner:
                            raise ValueError(f"Could not read excel or HTML file: {e_inner}")
                else:
                    # Fallback to CSV
                    try:
                        df = pd.read_csv(io.BytesIO(content), encoding='utf-8')
                    except Exception:
                        df = pd.read_csv(io.BytesIO(content), encoding='latin1')
        else:
            # It's a file path
            with open(vin_list_file_or_stream, 'rb') as f:
                prefix = f.read(300).lower()
            
            if b'<table' in prefix or b'<style' in prefix or b'<html' in prefix:
                dfs = pd.read_html(vin_list_file_or_stream, header=0)
                df = dfs[0]
            else:
                filename = str(vin_list_file_or_stream).lower()
                if filename.endswith('.xlsb'):
                    df = pd.read_excel(vin_list_file_or_stream, engine='pyxlsb')
                elif filename.endswith(('.xlsx', '.xls')):
                    try:
                        df = pd.read_excel(vin_list_file_or_stream)
                    except Exception:
                        try:
                            dfs = pd.read_html(vin_list_file_or_stream, header=0)
                            df = dfs[0]
                        except Exception as e_inner:
                            raise ValueError(f"Could not read excel or HTML file: {e_inner}")
                else:
                    df = pd.read_csv(vin_list_file_or_stream)
    except Exception as e:
        raise ValueError(f"Failed to read VIN list file: {e}")
        
    # 2. Extract VINs from the dataframe
    vins = []
    if not df.empty:
        # Find the column containing VINs
        vin_col = None
        
        # S0: Look for exact or contains 'vehicle code' or 'vc' or 'vehicle_code' (case-insensitive)
        for col in df.columns:
            col_lower = str(col).lower()
            if 'vehicle code' in col_lower or 'vehicle_code' in col_lower or col_lower == 'vc':
                vin_col = col
                break
                
        # S1: Look for 'vehicle'
        if vin_col is None:
            for col in df.columns:
                if 'vehicle' in str(col).lower():
                    vin_col = col
                    break
                    
        # S2: Look for 'vin' or 'chassis'
        if vin_col is None:
            for col in df.columns:
                col_lower = str(col).lower()
                if 'vin' in col_lower or 'chassis' in col_lower:
                    vin_col = col
                    break
                    
        # S3: Look for other keywords
        if vin_col is None:
            keywords = ['serial', 'number']
            for col in df.columns:
                col_str = str(col).lower()
                if any(k in col_str for k in keywords):
                    vin_col = col
                    break
                
        # S4: If no column matched keywords, find first column with strings of length 8-20
        if vin_col is None:
            for col in df.columns:
                non_nulls = df[col].dropna()
                if not non_nulls.empty:
                    # Check if at least 50% of values are strings of length 8-20
                    string_lengths = non_nulls.astype(str).str.len()
                    valid_len_pct = ((string_lengths >= 8) & (string_lengths <= 20)).mean()
                    if valid_len_pct >= 0.5:
                        vin_col = col
                        break
                        
        # Fallback to the first column
        if vin_col is None:
            vin_col = df.columns[0]
            
        # Extract values and clean them
        vins = df[vin_col].dropna().astype(str).str.strip().tolist()
        
    # 3. Write VINs to the workbook
    sheet_name = 'Stage wise Float' if track == 'TCF1' else 'Stage wise float'
    if sheet_name not in wb.sheetnames and track == 'TCF2':
        # TCF-2 might have trailing space in sheet name
        if 'Stage wise float ' in wb.sheetnames:
            sheet_name = 'Stage wise float '
            
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # Write VINs starting from row 2
        for idx, vin in enumerate(vins):
            row_idx = idx + 2
            ws.cell(row=row_idx, column=2).value = vin
            
    return wb


def update_paint_float_data(wb, paint_float_file_or_stream, track, expected_qty=None,
                            wip_files=None, yest_plan_file_or_stream=None,
                            next_3days_biw_plan=None, daily_capacities=None,
                            sub_limits_1=None, sub_limits_2=None):
    """
    Parses the Paint Float Report and updates columns D, E, F, G, H of
    the 'Stage wise Float' (TCF-1) or 'Stage wise float' (TCF-2) sheet,
    populates the 'Expected to Drop' column in the Vin Plan Check sheet,
    and handles the 3-day plan sequencing if WIP files are provided.
    """
    import pandas as pd
    import io
    
    # 1. Read the Paint Float Report using pandas
    try:
        if hasattr(paint_float_file_or_stream, 'read'):
            content = paint_float_file_or_stream.read()
            paint_float_file_or_stream.seek(0)
            
            prefix = content[:300].lower()
            if b'<table' in prefix or b'<style' in prefix or b'<html' in prefix:
                dfs = pd.read_html(io.BytesIO(content), header=0)
                df = dfs[0]
            else:
                filename = getattr(paint_float_file_or_stream, 'name', '').lower()
                if filename.endswith('.xlsb'):
                    df = pd.read_excel(io.BytesIO(content), engine='pyxlsb')
                elif filename.endswith(('.xlsx', '.xls')):
                    try:
                        df = pd.read_excel(io.BytesIO(content))
                    except Exception:
                        try:
                            dfs = pd.read_html(io.BytesIO(content), header=0)
                            df = dfs[0]
                        except Exception as e_inner:
                            raise ValueError(f"Could not read excel or HTML file: {e_inner}")
                else:
                    try:
                        df = pd.read_csv(io.BytesIO(content), encoding='utf-8')
                    except Exception:
                        df = pd.read_csv(io.BytesIO(content), encoding='latin1')
        else:
            with open(paint_float_file_or_stream, 'rb') as f:
                prefix = f.read(300).lower()
            
            if b'<table' in prefix or b'<style' in prefix or b'<html' in prefix:
                dfs = pd.read_html(paint_float_file_or_stream, header=0)
                df = dfs[0]
            else:
                filename = str(paint_float_file_or_stream).lower()
                if filename.endswith('.xlsb'):
                    df = pd.read_excel(paint_float_file_or_stream, engine='pyxlsb')
                elif filename.endswith(('.xlsx', '.xls')):
                    try:
                        df = pd.read_excel(paint_float_file_or_stream)
                    except Exception:
                        try:
                            dfs = pd.read_html(paint_float_file_or_stream, header=0)
                            df = dfs[0]
                        except Exception as e_inner:
                            raise ValueError(f"Could not read excel or HTML file: {e_inner}")
                else:
                    df = pd.read_csv(paint_float_file_or_stream)
    except Exception as e:
        raise ValueError(f"Failed to read Paint Float Report file: {e}")
        
    if df.empty:
        return wb
        
    # Helper to find column dynamically
    def find_col(patterns):
        for pattern in patterns:
            for col in df.columns:
                if pattern.lower() in str(col).lower():
                    return col
        return None
        
    # Find key columns
    vin_col = find_col(['vehicle code', 'vehicle_code', 'vc', 'vin', 'chassis'])
    shop_col = find_col(['shop'])
    
    if not vin_col or not shop_col:
        raise ValueError("Paint Float Report must contain 'VEHICLE CODE' and 'SHOP' columns.")
        
    # Filter by track/shop (TCF1 or TCF2)
    shop_val = 'TCF1' if track == 'TCF1' else 'TCF2'
    df_filtered = df[df[shop_col].astype(str).str.strip().str.upper() == shop_val.upper()]
    
    # Target sheet setup
    sheet_name = 'Stage wise Float' if track == 'TCF1' else 'Stage wise float'
    if sheet_name not in wb.sheetnames and track == 'TCF2':
        if 'Stage wise float ' in wb.sheetnames:
            sheet_name = 'Stage wise float '
            
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found in the workbook.")
        
    ws = wb[sheet_name]
    
    # Stage mappings to target columns:
    # Column D: PBS
    # Column E: Top Coat
    # Column F: Sealant
    # Column G: Ptced
    # Column H: Biw Lift
    stage_mappings = [
        (['pbs lift', 'pbs', 'pbs_lift'], 4),        # Col D (4)
        (['topcoat', 'top coat', 'top_coat'], 5),    # Col E (5)
        (['sealant', 'seal'], 6),                    # Col F (6)
        (['ptced', 'ptc'], 7),                       # Col G (7)
        (['biw lifting', 'biw lift', 'biw_lifting'], 8) # Col H (8)
    ]
    
    # Use a copy of the filtered dataframe to progressively deduct matched vehicles
    available_df = df_filtered.copy()
    
    for patterns, col_idx in stage_mappings:
        src_col = find_col(patterns)
        if src_col and not available_df.empty:
            # Filter rows in remaining available_df where the stage timestamp is not blank/empty
            valid_rows = available_df[
                available_df[src_col].notna() & 
                (available_df[src_col].astype(str).str.strip() != '')
            ]
            # Extract vehicle codes
            stage_vins = valid_rows[vin_col].dropna().astype(str).str.strip().tolist()
            
            # Write to target column starting from row 2
            for idx, vin in enumerate(stage_vins):
                row_idx = idx + 2
                ws.cell(row=row_idx, column=col_idx).value = vin
                
            # Deduct the processed rows from the pool
            available_df = available_df.drop(valid_rows.index)
            
    # ─────────────────────────────────────────────
    # UPDATE EXPECTED TO DROP COLUMN IN VIN PLAN CHECK
    # ─────────────────────────────────────────────
    
    # 1. Build lookup dictionary from VIN Color Plan sheet first
    # so we can use it to filter out Pune plant excluded models on the float floor
    color_sheet_name = 'VIN Color Plan'
    if color_sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{color_sheet_name}' not found in the workbook.")
        
    color_ws = wb[color_sheet_name]
    
    color_vc_col = 2 if track == 'TCF1' else 3
    engine_code_col = 3 if track == 'TCF1' else 4
    sales_desc_col = 5
    colour_col = 6
    
    vc_lookup = {}
    for r_color in range(3, color_ws.max_row + 1):
        vc_val = color_ws.cell(row=r_color, column=color_vc_col).value
        if vc_val is not None:
            vc_str = str(vc_val).strip().upper()
            sales_desc = color_ws.cell(row=r_color, column=sales_desc_col).value
            colour_val = color_ws.cell(row=r_color, column=colour_col).value
            eng_code = color_ws.cell(row=r_color, column=engine_code_col).value
            vc_lookup[vc_str] = (sales_desc, colour_val, eng_code)

    # 2. Parse/resolve target quantity
    if expected_qty is not None:
        try:
            expected_qty = int(expected_qty)
        except Exception:
            expected_qty = 500 if track == 'TCF1' else 100
    else:
        expected_qty = 500 if track == 'TCF1' else 100

    # 3. Pull vehicle codes from the Stage wise Float sheet columns D, E, F, G, H in sequence
    # and highlight those that fall within expected_qty in light green
    from openpyxl.styles import PatternFill
    light_green_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
    
    all_stage_vcs = []
    picked_count = 0
    # Columns D (4), E (5), F (6), G (7), H (8)
    for col_idx in [4, 5, 6, 7, 8]:
        for r in range(2, ws.max_row + 1):
            val = ws.cell(row=r, column=col_idx).value
            if val is not None and str(val).strip() != '':
                val_clean = str(val).strip().upper()
                
                # Retrieve from lookup to check Pune plant exclusions
                desc_val, col_val, eng_val = vc_lookup.get(val_clean, (None, None, ""))
                desc_str = str(desc_val).strip().upper() if desc_val is not None else ''
                color_str = str(col_val).strip().upper() if col_val is not None else ''
                
                if val_clean.startswith('5442') or 'ALTROZ' in desc_str or color_str == '000' or '000' in color_str:
                    # Clear cell fill and skip
                    ws.cell(row=r, column=col_idx).fill = PatternFill(fill_type=None)
                    continue
                    
                all_stage_vcs.append(val_clean)
                if picked_count < expected_qty:
                    # Highlight cell with light green
                    ws.cell(row=r, column=col_idx).fill = light_green_fill
                    picked_count += 1
                else:
                    # Clear cell fill if not picked
                    ws.cell(row=r, column=col_idx).fill = PatternFill(fill_type=None)
                    
    # 4. Count occurrences of picked VCs
    picked_vcs = all_stage_vcs[:expected_qty]
    remaining_vcs = all_stage_vcs[expected_qty:]
    from collections import Counter
    vc_counts = Counter(picked_vcs)
    
    # 5. Open the Vin Plan Check sheet
    check_sheet_name = 'Vin Plan Check' if track == 'TCF1' else 'Vin plan check'
    if check_sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{check_sheet_name}' not found in the workbook.")
        
    check_ws = wb[check_sheet_name]
    
    # Define track-specific coordinates:
    # TCF-1: Start row = 4, VC col = 2 (B), Expected Drop col = 6 (F)
    # TCF-2: Start row = 3, VC col = 3 (C), Expected Drop col = 7 (G)
    if track == 'TCF1':
        start_row = 4
        vc_col_idx = 2
        drop_col_idx = 6
    else:
        start_row = 3
        vc_col_idx = 3
        drop_col_idx = 7
        
    # Scan and update data rows
    r = start_row
    while True:
        sr_no_val = check_ws.cell(row=r, column=1).value
        if sr_no_val is None:
            break
        try:
            # Check if it represents an integer
            int(float(str(sr_no_val).strip()))
        except ValueError:
            break
            
        # Get the Color VC
        vc_val = check_ws.cell(row=r, column=vc_col_idx).value
        if vc_val is not None:
            vc_str = str(vc_val).strip().upper()
            # Fetch count or default to 0
            count = vc_counts.get(vc_str, 0)
            check_ws.cell(row=r, column=drop_col_idx).value = count
        else:
            check_ws.cell(row=r, column=drop_col_idx).value = 0
            
        r += 1
            
    # 6. Open the Day 1 Sequence sheet
    day1_sheet_name = 'Day 1 Sequence'
    if day1_sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{day1_sheet_name}' not found in the workbook.")
        
    day1_ws = wb[day1_sheet_name]
    
    # 7. Complete the sequencing
    if wip_files is not None:
        # 3-Day Plan mode
        # Parse WIP files
        wip_records = []
        if isinstance(wip_files, list):
            for wip_file in wip_files:
                if wip_file is not None:
                    wip_records.extend(parse_html_wip_vcs(wip_file))
        elif wip_files is not None:
            wip_records.extend(parse_html_wip_vcs(wip_files))
            
        # Parse dates and sort
        for rec in wip_records:
            try:
                rec['datetime'] = pd.to_datetime(rec['date_str'], format='%d/%m/%Y %I:%M:%S %p')
            except Exception:
                try:
                    rec['datetime'] = pd.to_datetime(rec['date_str'])
                except Exception:
                    rec['datetime'] = pd.Timestamp.min
        wip_records.sort(key=lambda x: x['datetime'])
        wip_vcs = [(x['vc'], 'light_blue') for x in wip_records]
        
        # Load yesterday's planned sequence from biw Next 2 days sheet
        master_yest_queue = []
        if yest_plan_file_or_stream is not None:
            if hasattr(yest_plan_file_or_stream, 'read'):
                content = yest_plan_file_or_stream.read()
                yest_plan_file_or_stream.seek(0)
                yest_wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
            else:
                yest_wb = openpyxl.load_workbook(yest_plan_file_or_stream, data_only=True)
                
            if 'biw Next 2 days' in yest_wb.sheetnames:
                yest_ws = yest_wb['biw Next 2 days']
                cols_to_read = [2, 3, 4] if track == 'TCF1' else [2, 3]
                for col_idx in cols_to_read:
                    for r in range(2, yest_ws.max_row + 1):
                        val = yest_ws.cell(row=r, column=col_idx).value
                        if val is not None and str(val).strip() != '':
                            master_yest_queue.append((str(val).strip().upper(), 'light_yellow'))
                            
        # Load Next 3-Days BIW Plan sequence if provided
        if next_3days_biw_plan is not None:
            try:
                if hasattr(next_3days_biw_plan, 'read'):
                    biw_content = next_3days_biw_plan.read()
                    next_3days_biw_plan.seek(0)
                    biw_wb = openpyxl.load_workbook(io.BytesIO(biw_content), data_only=True)
                else:
                    biw_wb = openpyxl.load_workbook(next_3days_biw_plan, data_only=True)
                
                # Sheet patterns
                if track == 'TCF1':
                    sheets_to_read = ['Punch Seq', 'NOVA']
                else:
                    sheets_to_read = ['Q5', 'TGDI ', 'Eturna']
                
                for sh_pat in sheets_to_read:
                    # Find sheet matching sh_pat (case-insensitive and stripped)
                    target_sh = None
                    for name in biw_wb.sheetnames:
                        if name.strip().upper() == sh_pat.strip().upper():
                            target_sh = name
                            break
                    if target_sh is not None:
                        ws_biw = biw_wb[target_sh]
                        # find ITEMCODE / ITEM CODE / ITEM_CODE / VC column
                        headers = [str(c.value).strip().upper() if c.value is not None else '' for c in ws_biw[1]]
                        vc_col_idx = None
                        for idx, h in enumerate(headers):
                            if any(pat in h for pat in ['ITEMCODE', 'ITEM CODE', 'ITEM_CODE', 'VEHICLE CODE', 'VEHICLE_CODE', 'VC']):
                                vc_col_idx = idx + 1
                                break
                        if vc_col_idx is None:
                            vc_col_idx = 2 if ws_biw.max_column >= 2 else 1
                        
                        for r in range(2, ws_biw.max_row + 1):
                            val = ws_biw.cell(row=r, column=vc_col_idx).value
                            if val is not None:
                                vc_str = str(val).strip().upper()
                                if vc_str != '':
                                    # Apply Pune plant TCF exclusions
                                    desc_val, col_val, eng_val = vc_lookup.get(vc_str, (None, None, ""))
                                    desc_str = str(desc_val).strip().upper() if desc_val is not None else ''
                                    color_str = str(col_val).strip().upper() if col_val is not None else ''
                                    
                                    if vc_str.startswith('5442') or 'ALTROZ' in desc_str or color_str == '000' or '000' in color_str:
                                        continue
                                    master_yest_queue.append((vc_str, 'light_pink'))
            except Exception as e_biw:
                print(f"Error reading Next 3-Days BIW Plan: {e_biw}")
                            
        # Assemble master queue and apply constraints
        def identify_vehicle_type(vc, desc, eng_code=""):
            vc_upper = str(vc).upper()
            desc_upper = str(desc or '').upper()
            eng_upper = str(eng_code or '').strip().upper()
            
            # TCF-1
            # EV
            if eng_upper.startswith('5468') or 'PUNCH.EV' in desc_upper or 'PUNCH EV' in desc_upper or 'PUNCH.EV' in vc_upper or 'PUNCH EV' in vc_upper:
                return 'NOVA'
            # CNG
            if 'CNG' in desc_upper or 'CNG' in vc_upper:
                if 'PUNCH' in desc_upper or 'PUNCH' in vc_upper or eng_upper.startswith('5497'):
                    return 'CNG'
                    
            # TCF-2
            # EV (Eturna)
            if eng_upper.startswith('5473') or eng_upper.startswith('5483') or 'HARRIER.EV' in desc_upper or 'HARRIER EV' in desc_upper or 'SAFARI.EV' in desc_upper or 'SAFARI EV' in desc_upper or 'HARRIER.EV' in vc_upper or 'SAFARI.EV' in vc_upper:
                return 'ETURNA'
            # TGDI (Petrol)
            if eng_upper.startswith('5478') or eng_upper.startswith('5479') or 'TGDI' in desc_upper or 'TGDI' in vc_upper:
                return 'TGDI'
            if ' BS6 P' in desc_upper or ' BS6 P2 P' in desc_upper or 'BS6 P2 P' in desc_upper:
                if 'BS6 D' not in desc_upper and 'P2 D' not in desc_upper:
                    return 'TGDI'
                    
            return 'OTHER'

        def build_day_sequence(indexed_queue, max_capacity, lim1_val, lim2_val, is_tcf1=True):
            if max_capacity <= 0:
                return [], indexed_queue, 0, 0
                
            # 1. Separate into categories
            ev_list = []
            c2_list = []
            other_list = []
            
            for orig_idx, item in indexed_queue:
                vc, color_key = item
                desc_val, col_val, eng_val = vc_lookup.get(vc, (None, None, ""))
                v_type = identify_vehicle_type(vc, desc_val, eng_val)
                
                if is_tcf1:
                    if v_type == 'NOVA':
                        ev_list.append((orig_idx, item))
                    elif v_type == 'CNG':
                        c2_list.append((orig_idx, item))
                    else:
                        other_list.append((orig_idx, item))
                else:
                    if v_type == 'ETURNA':
                        ev_list.append((orig_idx, item))
                    elif v_type == 'TGDI':
                        c2_list.append((orig_idx, item))
                    else:
                        other_list.append((orig_idx, item))
            
            # 2. Slice according to target capacities
            rem_val = max(0, max_capacity - lim1_val - lim2_val)
            
            selected_ev = ev_list[:lim1_val]
            selected_c2 = c2_list[:lim2_val]
            selected_other = other_list[:rem_val]
            
            # 3. Combine selected items and sort by original index to preserve priority order
            selected_combined = selected_ev + selected_c2 + selected_other
            selected_combined.sort(key=lambda x: x[0])
            
            day_list = [item for orig_idx, item in selected_combined]
            
            # 4. Form postponed queue of non-selected items
            selected_orig_indices = set(orig_idx for orig_idx, item in selected_combined)
            postponed = [(orig_idx, item) for orig_idx, item in indexed_queue if orig_idx not in selected_orig_indices]
            
            c1_count = len(selected_ev)
            c2_count = len(selected_c2)
            
            return day_list, postponed, c1_count, c2_count

        # Picked VCs counts (Paint Float expected drops)
        picked_n, picked_c, picked_e, picked_t = 0, 0, 0, 0
        for vc in picked_vcs:
            desc_val, col_val, eng_val = vc_lookup.get(vc, (None, None, ""))
            v_type = identify_vehicle_type(vc, desc_val, eng_val)
            if v_type == 'NOVA':
                picked_n += 1
            elif v_type == 'CNG':
                picked_c += 1
            elif v_type == 'ETURNA':
                picked_e += 1
            elif v_type == 'TGDI':
                picked_t += 1

        # Assemble master queue
        master_queue = []
        master_queue.extend([(vc, 'light_green') for vc in remaining_vcs])
        master_queue.extend(wip_vcs)
        master_queue.extend(master_yest_queue)
        
        # Unique priority indexing
        indexed_master_queue = list(enumerate(master_queue))
        
        is_tcf1 = (track == 'TCF1')
        if daily_capacities is None:
            caps = [900, 900, 900, 900] if is_tcf1 else [250, 250, 250]
        else:
            caps = list(daily_capacities)
            expected_len = 4 if is_tcf1 else 3
            while len(caps) < expected_len:
                caps.append(900 if is_tcf1 else 250)
                
        # Resolve sub limits 1 (EV / Eturna)
        if sub_limits_1 is None:
            lim1 = [160, 160, 160, 160] if is_tcf1 else [160, 160, 160]
        else:
            lim1 = list(sub_limits_1)
            expected_len = 4 if is_tcf1 else 3
            while len(lim1) < expected_len:
                lim1.append(160 if is_tcf1 else 160)
                
        # Resolve sub limits 2 (CNG / TGDI)
        if sub_limits_2 is None:
            lim2 = [350, 200, 350, 350] if is_tcf1 else [40, 40, 40]
        else:
            lim2 = list(sub_limits_2)
            expected_len = 4 if is_tcf1 else 3
            while len(lim2) < expected_len:
                lim2.append(350 if is_tcf1 else 40)
        
        # Day 1 Sequence
        day1_list, remaining_queue, d1_c1, d1_c2 = build_day_sequence(indexed_master_queue, caps[0], lim1[0], lim2[0], is_tcf1)
        
        # Day 2 Sequence
        day2_list, remaining_queue, d2_c1, d2_c2 = build_day_sequence(remaining_queue, caps[1], lim1[1], lim2[1], is_tcf1)
        
        # Day 3 Sequence
        day3_list, remaining_queue, d3_c1, d3_c2 = build_day_sequence(remaining_queue, caps[2], lim1[2], lim2[2], is_tcf1)
        
        # Day 4 Sequence (TCF-1 only)
        day4_list = []
        d4_c1, d4_c2 = 0, 0
        if is_tcf1:
            day4_list, remaining_queue, d4_c1, d4_c2 = build_day_sequence(remaining_queue, caps[3], lim1[3], lim2[3], is_tcf1)
            
        # Store summary counts on workbook object
        all_planned_items = day1_list + day2_list + day3_list + day4_list
        color_counts = {
            'light_green': 0,
            'light_blue': 0,
            'light_yellow': 0,
            'light_pink': 0
        }
        vc_lists = {
            'light_green': [],
            'light_blue': [],
            'light_yellow': [],
            'light_pink': []
        }
        for item in all_planned_items:
            vc, color_key = item
            if color_key in color_counts:
                color_counts[color_key] += 1
            if color_key in vc_lists:
                vc_lists[color_key].append(vc)
                
        if is_tcf1:
            wb.summary_counts = {
                'picked': {'nova': picked_n, 'cng': picked_c, 'eturna': 0, 'tgdi': 0},
                'day1': {'nova': d1_c1, 'cng': d1_c2, 'eturna': 0, 'tgdi': 0},
                'day2': {'nova': d2_c1, 'cng': d2_c2, 'eturna': 0, 'tgdi': 0},
                'day3': {'nova': d3_c1, 'cng': d3_c2, 'eturna': 0, 'tgdi': 0},
                'day4': {'nova': d4_c1, 'cng': d4_c2, 'eturna': 0, 'tgdi': 0},
                'targets': {
                    'day1': {'nova': lim1[0], 'cng': lim2[0], 'eturna': 0, 'tgdi': 0},
                    'day2': {'nova': lim1[1], 'cng': lim2[1], 'eturna': 0, 'tgdi': 0},
                    'day3': {'nova': lim1[2], 'cng': lim2[2], 'eturna': 0, 'tgdi': 0},
                    'day4': {'nova': lim1[3], 'cng': lim2[3], 'eturna': 0, 'tgdi': 0},
                },
                'colors': color_counts,
                'vc_lists': vc_lists
            }
        else:
            wb.summary_counts = {
                'picked': {'nova': 0, 'cng': 0, 'eturna': picked_e, 'tgdi': picked_t},
                'day1': {'nova': 0, 'cng': 0, 'eturna': d1_c1, 'tgdi': d1_c2},
                'day2': {'nova': 0, 'cng': 0, 'eturna': d2_c1, 'tgdi': d2_c2},
                'day3': {'nova': 0, 'cng': 0, 'eturna': d3_c1, 'tgdi': d3_c2},
                'targets': {
                    'day1': {'nova': 0, 'cng': 0, 'eturna': lim1[0], 'tgdi': lim2[0]},
                    'day2': {'nova': 0, 'cng': 0, 'eturna': lim1[1], 'tgdi': lim2[1]},
                    'day3': {'nova': 0, 'cng': 0, 'eturna': lim1[2], 'tgdi': lim2[2]},
                },
                'colors': color_counts,
                'vc_lists': vc_lists
            }
        
        # Define highlight patterns
        from openpyxl.styles import PatternFill
        fills = {
            'light_green': PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid'),
            'light_blue': PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid'),
            'light_yellow': PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid'),
            'light_pink': PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
        }
        
        # Write Day 1 Sequence
        for idx, item in enumerate(day1_list):
            vc, color_key = item
            row_idx = idx + 3
            if row_idx <= day1_ws.max_row:
                day1_ws.cell(row=row_idx, column=2).value = vc
                desc_val, col_val, eng_val = vc_lookup.get(vc, (None, None, ""))
                day1_ws.cell(row=row_idx, column=3).value = desc_val
                day1_ws.cell(row=row_idx, column=4).value = col_val
                if day1_ws.cell(row=row_idx, column=1).value is None:
                    day1_ws.cell(row=row_idx, column=1).value = idx + 1
                if color_key in fills:
                    day1_ws.cell(row=row_idx, column=2).fill = fills[color_key]
                    
        # Write Day 2, Day 3, Day 4 to biw Next 2 days sheet
        if 'biw Next 2 days' in wb.sheetnames:
            biw_ws = wb['biw Next 2 days']
            # Day 2 -> B (2)
            for idx, item in enumerate(day2_list):
                vc, color_key = item
                row_idx = idx + 2
                if row_idx <= biw_ws.max_row:
                    biw_ws.cell(row=row_idx, column=2).value = vc
                    if biw_ws.cell(row=row_idx, column=1).value is None:
                        biw_ws.cell(row=row_idx, column=1).value = idx + 1
                    if color_key in fills:
                        biw_ws.cell(row=row_idx, column=2).fill = fills[color_key]
            # Day 3 -> C (3)
            for idx, item in enumerate(day3_list):
                vc, color_key = item
                row_idx = idx + 2
                if row_idx <= biw_ws.max_row:
                    biw_ws.cell(row=row_idx, column=3).value = vc
                    if biw_ws.cell(row=row_idx, column=1).value is None:
                        biw_ws.cell(row=row_idx, column=1).value = idx + 1
                    if color_key in fills:
                        biw_ws.cell(row=row_idx, column=3).fill = fills[color_key]
            # Day 4 -> D (4) - TCF-1 only
            if track == 'TCF1':
                for idx, item in enumerate(day4_list):
                    vc, color_key = item
                    row_idx = idx + 2
                    if row_idx <= biw_ws.max_row:
                        biw_ws.cell(row=row_idx, column=4).value = vc
                        if biw_ws.cell(row=row_idx, column=1).value is None:
                            biw_ws.cell(row=row_idx, column=1).value = idx + 1
                        if color_key in fills:
                            biw_ws.cell(row=row_idx, column=4).fill = fills[color_key]
    else:
        # Standard Phase 4 mode: only remaining Float VCs in Day 1 Sequence
        from openpyxl.styles import PatternFill
        lg_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
        for idx, vc in enumerate(remaining_vcs):
            row_idx = idx + 3
            if row_idx <= day1_ws.max_row:
                day1_ws.cell(row=row_idx, column=2).value = vc
                desc_val, col_val, eng_val = vc_lookup.get(vc, (None, None, ""))
                day1_ws.cell(row=row_idx, column=3).value = desc_val
                day1_ws.cell(row=row_idx, column=4).value = col_val
                if day1_ws.cell(row=row_idx, column=1).value is None:
                    day1_ws.cell(row=row_idx, column=1).value = idx + 1
                day1_ws.cell(row=row_idx, column=2).fill = lg_fill
                    
    return wb
