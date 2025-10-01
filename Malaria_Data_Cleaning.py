import pandas as pd
import numpy as np
from datetime import datetime


import streamlit as st
import pandas as pd
import io
import numpy as np
from datetime import datetime


# Cleaning Function
def clean_malaria_data(df):

     # Strip column names
    df.columns = df.columns.str.strip()

    # Clean string cells (strip; keep NaN)
    obj_cols = df.select_dtypes(include=['object']).columns
    df[obj_cols] = (df[obj_cols].applymap(lambda x: str(x).strip() if pd.notna(x) else x)
                                .replace('nan', np.nan))

    # Case-insensitive column lookup
    col_map = {c.lower(): c for c in df.columns}
    get_col = lambda name: col_map.get(name.lower())

    # Helpers
    def add_comment(idx, msg):
        prev = df.at[idx, 'COMMENT']
        df.at[idx, 'COMMENT'] = (prev + '; ' if str(prev).strip() else '') + msg

    def to_num(x):
        try: return float(str(x).strip())
        except: return None

    def is_blank(x):
        s = '' if pd.isna(x) else str(x).strip().lower()
        return s in ['', 'na', 'n/a', 'none', 'nan', 'not treated']
    
        # --- robust rule matcher (numeric-safe / string-safe / blank-aware) ---
    def matches_rule(val, rule):
        # None => must be blank
        if rule is None:
            return is_blank(val)

        # numeric rule => compare numerically (handles "10" vs 10, "7.5" vs 7.5)
        if isinstance(rule, (int, float)):
            v = to_num(val)
            return v is not None and float(v) == float(rule)

        # string rule => trim + exact match (e.g., "7.5mg")
        return ('' if pd.isna(val) else str(val)).strip() == str(rule)


    def nonempty_row(idx, skip=()):
        cols = [c for c in df.columns if c not in skip]
        mask = df.loc[idx, cols].notna() & (df.loc[idx, cols].astype(str).str.strip() != '')
        return mask.any()

    def rows_with_data(skip=()):
        return [i for i in df.index if nonempty_row(i, skip=skip)]

    # Ensure COMMENT exists
    if 'COMMENT' not in df.columns:
        df['COMMENT'] = ''
    else:
        df['COMMENT'] = df['COMMENT'].fillna('')

    # === DATE PARSING & VALIDATION ===
    date_col = get_col('SCREENING_DATE')
    if date_col:
        # Sort by the raw column first (as original code)
        df = df.sort_values(by=date_col)

        start_date = pd.to_datetime('2024-01-01')
        today      = pd.to_datetime(datetime.now().strftime('%Y-%m-%d'))

        # STRICTLY allowed input formats (no dotted formats allowed)
        strict_formats = [
            '%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%Y/%m/%d',
            '%m/%d/%Y', '%m-%d-%Y', '%d/%m/%Y', '%d-%m-%Y',
            '%m/%d/%y', '%m-%d-%y', '%d/%m/%y', '%d-%m-%y'
        ]

        for idx, val in df[date_col].items():
            if pd.isna(val) or str(val).strip() == '':
                continue

            valid, parsed = False, None

            # Already a proper Timestamp from Excel
            if isinstance(val, pd.Timestamp):
                parsed, valid = val, True
            else:
                s = str(val).strip()
                # Try ONLY the formats above; no permissive fallback
                for fmt in strict_formats:
                    try:
                        parsed = datetime.strptime(s, fmt)
                        if 2024 < parsed.year < 2100:
                            valid = True
                            break
                    except Exception:
                        continue

            # If invalid, add error and PRESERVE original cell value
            if not valid:
                add_comment(idx, f"Date_Format Error: [{date_col}]")
                continue

            # Range check
            if not (start_date <= pd.to_datetime(parsed) <= today):
                add_comment(idx, f"Date_Range Error: [{date_col}] (Date out of range)")

            # Normalize valid dates to DD-MMM-YY (e.g., 04-Aug-25)
            df.at[idx, date_col] = pd.to_datetime(parsed).strftime('%d-%b-%y')


    # === NUMERIC & RANGE CHECKS ===
    numeric_vars       = ['AGE_YEAR', 'ACT_TABLET', 'CQ_TABLET', 'PQ_TABLET']
    integer_only       = {'ACT_TABLET'}
    tablet_vars        = ['ACT_TABLET', 'CQ_TABLET', 'PQ_TABLET']
    age_col = get_col('AGE_YEAR')

    for var in numeric_vars:
        c = get_col(var)
        if not c: continue
        for idx, v in df[c].items():
            if pd.isna(v) or str(v).strip()=='' : continue
            n = to_num(v)
            if n is None:
                add_comment(idx, f"Numeric_Check Error: [{c}]"); continue
            if var in integer_only and n != int(n):
                add_comment(idx, f"Numeric_Check Error: [{c}] (must be integer)")

    # Age 0–100
    if age_col:
        for idx, v in df[age_col].items():
            if pd.isna(v) or str(v).strip()=='' : continue
            n = to_num(v)
            if n is None or not (0 <= n <= 100):
                add_comment(idx, f"Range_Check Error: [{age_col}] (must be between 0 and 100)")

    # Tablets >0
    for var in tablet_vars:
        c = get_col(var)
        if not c: continue
        for idx, v in df[c].items():
            if pd.isna(v) or str(v).strip()=='' : continue
            n = to_num(v)
            if n is None or n <= 0:
                add_comment(idx, f"Range_Check Error: [{c}] (must be > 0)")

    # === MISSING CHECK (only for rows with some data) ===
    required = ['SR_NO','ORGANIZATION','STATE_REGION','TOWNSHIP','CHANNEL','REPORTING_MONTH',
                'SCREENING_DATE','PATIENT_NAME','AGE_YEAR','SEX','RDT']
    required_cols = [get_col(x) for x in required if get_col(x)]
    sr_no_col = get_col('SR_NO')
    for idx in df.index:
        # skip fully empty rows
        other = [c for c in df.columns if c != sr_no_col]
        non_empty = (df.loc[idx, other].notna() & (df.loc[idx, other].astype(str).str.strip()!='')).sum()
        if non_empty == 0: continue
        for c in required_cols:
            val = df.at[idx, c] if c in df.columns else None
            if pd.isna(val) or (isinstance(val, str) and val.strip()==''):
                add_comment(idx, f"Missing_Check Error: [{c}]")

    # === CHOICE / VALIDATION ===
    validation = {
        'STATE_REGION': ['Ayeyarwady', 'Bago (East)', 'Bago (West)', 'Chin', 'Kachin', 'Kayah', 'Kayin',
                        'Magway', 'Mandalay', 'Mon', 'Nay Pyi Taw', 'Rakhine', 'Sagaing', 'Shan (East)',
                        'Shan (North)', 'Shan (South)', 'Tanintharyi', 'Yangon'],
        'CHANNEL' : ['clinic', 'mobile team', 'ICHV'],
        'REPORTING_MONTH' : ['January','February','March','April','May','June','July','August',
                            'September','October','November','December'],
        'SEX' : ['Male','Female'],
        'DISABILITY' : ['Yes','No'],
        'IDP' : ['Yes','No'],
        'PATIENT_ADDRESS_CATEGORY' : ['Same as village','Inside catchment','Inside township',
                                    'Outside township',"Don't know"],
        'PREG' : ['Yes','No','N/A'],
        'RDT' : ['Negative','Pf','Pv/ Non Pf','Mix','Not exam by RDT'],
        'SYMPTOMS' : ['normal','severe','N/A'],
        'ACT' : ['Not Treated','ACT 3','ACT 6','ACT 12','ACT 18','ACT 24','N/A'],
        'PQ_MG' : ['7.5mg','15mg'],
        'REFERRAL' : ['Yes','No','N/A'],
        'MALARIA_DEATH' : ['Yes','No','N/A'],
        'TREATMENT_24HRS' : ['<=24hr','>24hr','N/A'],
        'TRAVELLING_HISTORY' : ['Yes','No']
    }

    for var, vals in validation.items():
        c = get_col(var)
        if not c: continue
        lower_map = {str(v).lower(): v for v in vals}
        for idx, v in df[c].items():
            if pd.isna(v) or str(v)=='' : continue
            s = str(v).strip()
            if var == 'RDT':
                if s not in vals:
                    add_comment(idx, f"Choice_Check Error: [{c}] (Valid values: {', '.join(vals)})")
            else:
                key = s.lower()
                if key in lower_map:
                    df.at[idx, c] = lower_map[key]  # normalize case
                else:
                    add_comment(idx, f"Choice_Check Error: [{c}] (Valid values: {', '.join(vals)})")

    # Tidy COMMENT
    df['COMMENT'] = df['COMMENT'].astype(str).str.rstrip('; ').str.strip().fillna('')

    # === DUPLICATE CHECK ===
    dup_keys = ['TOWNSHIP','SCREENING_DATE','PATIENT_NAME','AGE_YEAR','SEX','RDT']
    actual_dup = [get_col(x) for x in dup_keys if get_col(x)]
    if len(actual_dup) == len(dup_keys):
        scr_c = get_col('SCREENING_DATE')
        if 'DUPLICATE' not in df.columns:
            # place before COMMENT if exists
            pos = df.columns.get_loc('COMMENT') if 'COMMENT' in df.columns else len(df.columns)
            df.insert(pos, 'DUPLICATE', None)

        # convert to dt for sort
        tmp = df[actual_dup].copy()
        tmp['original_index'] = df.index
        tmp['screening_date_dt'] = pd.to_datetime(df[scr_c], format='mixed', errors='coerce')
        tmp = tmp.dropna(subset=['screening_date_dt']).sort_values(actual_dup + ['screening_date_dt'])
        tmp['is_dup'] = tmp.duplicated(subset=actual_dup, keep='first')

        df.loc[tmp[~tmp['is_dup']]['original_index'], 'DUPLICATE'] = 'First'
        df.loc[tmp[tmp['is_dup']]['original_index'], 'DUPLICATE']   = 'Duplicate'

    print("\n=== Duplicate Check Results ===")

    # === RETESTING (First/Repeat within groups by 28 days) ===
    def ensure_col_before(col_name, before='COMMENT'):
        if col_name not in df.columns:
            pos = df.columns.get_loc(before) if before in df.columns else len(df.columns)
            df.insert(pos, col_name, None)

    ensure_col_before('RETESTING', before='COMMENT')

    scr_c = get_col('SCREENING_DATE')
    town_c, name_c, sex_c, rdt_c = map(get_col, ['TOWNSHIP','PATIENT_NAME','SEX','RDT'])
    age_c = get_col('AGE_YEAR')

    needed = [scr_c, town_c, name_c, age_c, sex_c, rdt_c]
    if all(needed):
        idxs = rows_with_data(skip=(get_col('NO'), 'COMMENT','RETESTING'))
        sub = df.loc[idxs, [scr_c, town_c, name_c, age_c, sex_c, rdt_c]].copy()
        sub['date_temp'] = pd.to_datetime(sub[scr_c], format='mixed', errors='coerce')
        sub = sub.dropna(subset=['date_temp']).reset_index()

        for _, g in sub.groupby([town_c, name_c, age_c, sex_c, rdt_c], dropna=False):
            g = g.sort_values('date_temp')
            df.at[g.iloc[0]['index'], 'RETESTING'] = 'First'
            for i in range(1, len(g)):
                cur, prev = g.iloc[i], g.iloc[i-1]
                df.at[cur['index'], 'RETESTING'] = 'First' if (cur['date_temp'] - prev['date_temp']).days > 28 else 'Repeat'
    else:
        print("Warning: Required columns for consistency check are missing. Retesting status will not be assigned.")

    # === CONSISTENCY CHECKS ===
    # 1) REPORTING_MONTH vs SCREENING_DATE (26th prev → 25th curr)
    rep_c = get_col('REPORTING_MONTH')
    if rep_c and scr_c:
        for idx, row in df.iterrows():
            rep, scr = row.get(rep_c), row.get(scr_c)
            if pd.isna(rep) or pd.isna(scr) or str(scr).strip()=='' : continue
            scr_dt = pd.to_datetime(scr, format='%d-%b-%y', errors='coerce')
            if pd.isna(scr_dt): continue
            ranges = {
                'January':   ((scr_dt.year-1,12,26),(scr_dt.year,1,25)),
                'February':  ((scr_dt.year,1,26),(scr_dt.year,2,25)),
                'March':     ((scr_dt.year,2,26),(scr_dt.year,3,25)),
                'April':     ((scr_dt.year,3,26),(scr_dt.year,4,25)),
                'May':       ((scr_dt.year,4,26),(scr_dt.year,5,25)),
                'June':      ((scr_dt.year,5,26),(scr_dt.year,6,25)),
                'July':      ((scr_dt.year,6,26),(scr_dt.year,7,25)),
                'August':    ((scr_dt.year,7,26),(scr_dt.year,8,25)),
                'September': ((scr_dt.year,8,26),(scr_dt.year,9,25)),
                'October':   ((scr_dt.year,9,26),(scr_dt.year,10,25)),
                'November':  ((scr_dt.year,10,26),(scr_dt.year,11,25)),
                'December':  ((scr_dt.year,11,26),(scr_dt.year,12,25)),
            }
            if rep in ranges:
                start = datetime(*ranges[rep][0]); end = datetime(*ranges[rep][1])
                if not (start <= scr_dt <= end):
                    add_comment(idx, f"Consistency_Check Error: [{scr_c} vs {rep_c}]")

    # 2) PATIENT_NAME vs SEX (prefix heuristics)
    name_c = get_col('PATIENT_NAME'); sex_c = get_col('SEX')
    if name_c and sex_c:
        male_keys   = ['U ', 'Ko ', 'Mg ', 'Maung']
        female_keys = ['Daw ', 'Ma ']
        for idx, row in df.iterrows():
            pname, sex = str(row.get(name_c, '') or ''), str(row.get(sex_c, '') or '')
            if not pname.strip() or not sex.strip(): continue
            lower = pname.lower()
            if any(lower.startswith(k.lower()) for k in male_keys) and sex.strip().lower()!='male':
                add_comment(idx, f"Consistency_Check Error: [{name_c} suggests Male but {sex_c}='{sex}']")
            if any(lower.startswith(k.lower()) for k in female_keys) and sex.strip().lower()!='female':
                add_comment(idx, f"Consistency_Check Error: [{name_c} suggests Female but {sex_c}='{sex}']")

    # 3) SEX vs PREG
    preg_c = get_col('PREG')
    if sex_c and preg_c:
        for idx, row in df.iterrows():
            sex = str(row.get(sex_c, '') or '').lower()
            preg = str(row.get(preg_c, '') or '').lower()
            if preg=='yes' and sex!='female':
                add_comment(idx, f"Consistency_Check Error: [If {preg_c} is 'Yes' {sex_c} must be 'Female']")

    # 4) SYMPTOMS vs RDT (match original logic: Pv & ACT blank & SYMPTOMS != normal)
    rdt_c = get_col('RDT'); act_tab_c = get_col('ACT_TABLET'); sym_c = get_col('SYMPTOMS')
    if rdt_c and act_tab_c and sym_c:
        for idx, row in df.iterrows():
            # NaN-safe reads
            rdt_txt = '' if pd.isna(row.get(rdt_c)) else str(row.get(rdt_c)).strip().lower()
            act_val = row.get(act_tab_c)
            act_blank = (pd.isna(act_val) or str(act_val).strip() == '')
            sym_txt = '' if pd.isna(row.get(sym_c)) else str(row.get(sym_c)).strip().lower()

            # Original rule: if 'pv' anywhere, ACT_TABLET is blank, and SYMPTOMS == 'severe' -> error
            if ('pv' in rdt_txt) and act_blank and (sym_txt == 'severe'):
                add_comment(idx, f"Consistency_Check Error: [{act_tab_c} is not given, and {rdt_c} is Pv, is {sym_c} severe??]")


    # 5) ACT vs ACT_TABLET  — only error if ACT indicates treatment AND ACT_TABLET is blank
    act_c = get_col('ACT')
    if act_c and act_tab_c:
        for idx, row in df.iterrows():
            raw_act = row.get(act_c)
            act_type = '' if pd.isna(raw_act) else str(raw_act).strip()

            raw_tab = row.get(act_tab_c)
            act_tab_blank = (pd.isna(raw_tab) or str(raw_tab).strip() == '')

            # If ACT is blank / N/A / Not Treated, no error even if ACT_TABLET is blank
            if act_type not in ['', 'N/A', 'Not Treated'] and act_tab_blank:
                add_comment(idx, f"Consistency Error: [{act_c}='{act_type}' but {act_tab_c} is blank]")

    # 6) PQ_TABLET vs PQ_MG
    pq_tab_c = get_col('PQ_TABLET'); pq_mg_c = get_col('PQ_MG')
    if pq_tab_c and pq_mg_c:
        for idx, row in df.iterrows():
            pq_tab = row.get(pq_tab_c, '')
            pq_mg  = str(row.get(pq_mg_c, '') or '').strip()
            if not is_blank(pq_tab) and pq_mg not in ['7.5mg','15mg']:
                add_comment(idx, f"Consistency Error: [{pq_tab_c} filled but {pq_mg_c} is blank]")
            if pq_mg in ['7.5mg','15mg'] and is_blank(pq_tab):
                add_comment(idx, f"Consistency Error: [{pq_mg_c} is '{pq_mg}' but {pq_tab_c} is blank]")

    # === IHRP FLAG ===
    ihrp_townships = {
        "Bhamo","Mohnyin","Shwegu","Buthidaung","Kyauktaw","Maungdaw","Minbya","Mrauk-U","Myebon","Paletwa","Ponnagyun",
        "Hseni","Kunlong","Kutkai","Kyaukme","Laukkaing","Manton","Muse","Namhkan","Namhsan","Namtu","Lashio",
        "Chinshwehaw Sub-township (Kokang SAZ)","Chinshwehaw",
        "Bawlake","Demoso","Hpasawng","Hpruso","Loikaw","Mese","Shadaw","Thandaunggyi","Pyinmana","Pekon","Hsihseng",
        "Falam","Hakha","Kanpetlet","Matupi","Mindat","Tedim","Thantlang","Tonzang",
        "Ayadaw","Banmauk","Budalin","Chaung-U","Homalin","Indaw","Kale","Kanbalu","Kani","Katha","Kawlin","Khin-U",
        "Kyunhla","Mawlaik","Mingin","Myaung","Myinmu","Pale","Pinlebu","Sagaing","Salingyi","Shwebo","Tabayin","Taze",
        "Tigyaing","Wetlet","Wuntho","Ye-U","Yinmarbin",
        "Gangaw","Myaing","Pauk","Saw","Tilin","Yesagyo","Mindon","Pakokku","Salin","Seikphyu","Thayet","Tamu","Sinbaungwe"
    }
    town_c = get_col("TOWNSHIP"); sdp_c = get_col("SDP")
    df['IHRP'] = None
    if town_c:
        idxs = rows_with_data(skip=(get_col("SR_NO"), "COMMENT","Retesting","Duplicate","IHRP"))
        for idx in idxs:
            t = str(df.at[idx, town_c]).strip() if pd.notna(df.at[idx, town_c]) else ''
            sdp = str(df.at[idx, sdp_c]).strip() if sdp_c and pd.notna(df.at[idx, sdp_c]) else ''
            df.at[idx, "IHRP"] = "IHRP" if t in ihrp_townships else "Non-IHRP"
            if t=="Mohnyin" and sdp=="KNSM":
                df.at[idx, "IHRP"] = "Non-IHRP"

    # place IHRP before COMMENT (after RETESTING if present)
    try:
        cols = df.columns.tolist()
        if "IHRP" in cols:
            cols.remove("IHRP")
            insert_after = 'RETESTING' if 'RETESTING' in cols else 'COMMENT'
            pos = df.columns.get_loc(insert_after)+1 if insert_after in df.columns else len(cols)
            cols.insert(pos, "IHRP")
            df = df[cols]
    except: pass

    # === ORAL_TX_FOR_NON_IHRP ===
    oral_non_c = "ORAL_TX_FOR_NON_IHRP"
    df[oral_non_c] = None
    age_c = get_col("AGE_YEAR"); rdt_c = get_col("RDT")
    act_tab_c = get_col("ACT_TABLET"); cq_tab_c = get_col("CQ_TABLET"); pq_tab_c = get_col("PQ_TABLET")
    idxs = rows_with_data(skip=(get_col("SR_NO"), "COMMENT","Retesting","Duplicate","IHRP", oral_non_c))

    for idx in idxs:
        age = to_num(df.at[idx, age_c]) if age_c else None
        rdt = (str(df.at[idx, rdt_c]).strip().lower() if rdt_c and pd.notna(df.at[idx, rdt_c]) else '')
        act = to_num(df.at[idx, act_tab_c]) if act_tab_c else None
        cq  = to_num(df.at[idx, cq_tab_c])  if cq_tab_c  else None
        pq  = to_num(df.at[idx, pq_tab_c])  if pq_tab_c  else None

        result = "No"
        if age is not None:
            if age < 1:
                if rdt in ["pf","mix"] and (act and act>0) and is_blank(df.at[idx, cq_tab_c]): result = "Yes"
                elif rdt == "pv/ non pf" and (cq and cq>0) and is_blank(df.at[idx, act_tab_c]): result = "Yes"
            elif age > 1:
                if rdt in ["pf","mix"] and (act and act>0) and (pq and pq>0) and is_blank(df.at[idx, cq_tab_c]): result = "Yes"
            if rdt=="pv/ non pf" and (cq and cq>0) and (pq and pq>0) and is_blank(df.at[idx, act_tab_c]): result = "Yes"

        df.at[idx, oral_non_c] = result

    # insert after IHRP
    try:
        cols = df.columns.tolist(); cols.remove(oral_non_c)
        pos = df.columns.get_loc("IHRP")+1 if "IHRP" in df.columns else df.columns.get_loc("COMMENT")
        cols.insert(pos, oral_non_c); df = df[cols]
    except: pass

    # === ORAL_TX_1 & ORAL_TX_2 (rule engines) ===
    def apply_tx_rules(rules, target_col, skip_extra=()):
        df[target_col] = None
        idxs = rows_with_data(skip=(get_col('SR_NO'), 'COMMENT','Retesting','Duplicate','IHRP', oral_non_c, target_col, *skip_extra))
        a = get_col('AGE_YEAR'); r = get_col('RDT'); act = get_col('ACT_TABLET'); cq = get_col('CQ_TABLET'); pq = get_col('PQ_TABLET'); pqmg = get_col('PQ_MG')
        for idx in idxs:
            age = to_num(df.at[idx, a]) if a else None
            rdt = str(df.at[idx, r]).strip().lower() if r and pd.notna(df.at[idx, r]) else ''
            act_v, cq_v, pq_v = df.at[idx, act], df.at[idx, cq], df.at[idx, pq]
            pqmg_v = str(df.at[idx, pqmg]).strip() if pqmg and pd.notna(df.at[idx, pqmg]) else ''
            res = "No"
            if age is not None and rdt in rules:
                for (mn, mx, act_r, cq_r, pqmg_r, pq_r) in rules[rdt]:
                    if mn <= age <= mx:
                        conds = [
                            matches_rule(act_v,  act_r),
                            matches_rule(cq_v,   cq_r),
                            matches_rule(pqmg_v, pqmg_r),
                            matches_rule(pq_v,   pq_r),
                        ]

                        if all(conds): res = "Yes"; break
            df.at[idx, target_col] = res

    tx1_rules = {
        "pf":[(0,0.99,3,None,None,None),(1,4,6,None,"7.5mg",1),(5,9,12,None,"7.5mg",2),(10,14,18,None,"7.5mg",4),(15,100,24,None,"7.5mg",6)],
        "pv/ non pf":[(0,0.99,None,1,None,None),(1,4,None,4,"7.5mg",7),(5,9,None,5,"7.5mg",14),(10,14,None,7.5,"7.5mg",21),(15,100,None,10,"7.5mg",28)],
        "mix":[(0,0.99,3,None,None,None),(1,4,6,None,"7.5mg",7),(5,9,12,None,"7.5mg",14),(10,14,18,None,"7.5mg",21),(15,100,24,None,"7.5mg",28)],
    }
    tx2_rules = {
        "pf":[(0,0.99,3,None,None,None),(1,4,6,None,"7.5mg",0.5),(5,9,12,None,"7.5mg",1),(10,14,18,None,"7.5mg",1.5),(15,100,24,None,"7.5mg",2)],
        "pv/ non pf":[(0,0.99,None,1,None,None),(1,4,None,4,"7.5mg",7),(5,9,None,5,"7.5mg",14),(10,14,None,7.5,"7.5mg",21),(15,100,None,10,"7.5mg",28)],
        "mix":[(0,0.99,3,None,None,None),(1,4,6,None,"7.5mg",7),(5,9,12,None,"7.5mg",14),(10,14,18,None,"7.5mg",21),(15,100,24,None,"7.5mg",28)],
    }

    apply_tx_rules(tx1_rules, 'ORAL_TX_1')
    apply_tx_rules(tx2_rules, 'ORAL_TX_2', skip_extra=('ORAL_TX_1',))

    # reorder ORAL_TX_1/2
    try:
        cols = df.columns.tolist()
        for col in ['ORAL_TX_2','ORAL_TX_1']:
            if col in cols: cols.remove(col)
        pos = df.columns.get_loc('ORAL_TX_FOR_NON_IHRP')+1 if 'ORAL_TX_FOR_NON_IHRP' in df.columns else df.columns.get_loc('COMMENT')
        cols[pos:pos] = ['ORAL_TX_1','ORAL_TX_2']
        df = df[cols]
    except: pass

    # === TX_GUIDELINE ===
    txg_c = 'TX_GUIDELINE'
    df[txg_c] = None

    preg_c = get_col('PREG'); pq_c = get_col('PQ_TABLET'); ihrp_c = 'IHRP'
    refer_c = get_col('REFERRAL'); pqmg_c = get_col('PQ_MG')
    age_c, act_cq_vals = get_col('AGE_YEAR'), (get_col('ACT_TABLET'), get_col('CQ_TABLET'))
    act_tab_c, cq_tab_c = act_cq_vals

    idxs = rows_with_data(skip=(get_col('SR_NO'), 'COMMENT', txg_c))
    for idx in idxs:
        preg = str(df.at[idx, preg_c]).strip().lower() if preg_c and pd.notna(df.at[idx, preg_c]) else ''
        pq_tab = df.at[idx, pq_c] if pq_c and pd.notna(df.at[idx, pq_c]) else ''
        ihrp = str(df.at[idx, ihrp_c]).strip().lower() if ihrp_c in df.columns and pd.notna(df.at[idx, ihrp_c]) else ''
        oral_non = str(df.at[idx, 'ORAL_TX_FOR_NON_IHRP']).strip().lower() if 'ORAL_TX_FOR_NON_IHRP' in df.columns and pd.notna(df.at[idx, 'ORAL_TX_FOR_NON_IHRP']) else ''
        oral1 = str(df.at[idx, 'ORAL_TX_1']).strip().lower() if 'ORAL_TX_1' in df.columns and pd.notna(df.at[idx, 'ORAL_TX_1']) else ''
        oral2 = str(df.at[idx, 'ORAL_TX_2']).strip().lower() if 'ORAL_TX_2' in df.columns and pd.notna(df.at[idx, 'ORAL_TX_2']) else ''
        rdt = str(df.at[idx, rdt_c]).strip().lower() if rdt_c and pd.notna(df.at[idx, rdt_c]) else ''
        refer = str(df.at[idx, refer_c]).strip().lower() if refer_c and pd.notna(df.at[idx, refer_c]) else ''

        # quick numeric helpers
        age_v = to_num(df.at[idx, age_c]) if age_c else None
        act_v = to_num(df.at[idx, act_tab_c]) if act_tab_c else None
        cq_v  = to_num(df.at[idx, cq_tab_c])  if cq_tab_c  else None

        # Disqualify: PREG with Non-IHRP oral
        if oral_non == 'yes' and preg == 'yes':
            df.at[idx, txg_c] = 'No'; continue

        # All YES shortcuts
        if ihrp == 'non-ihrp' and oral_non == 'yes':
            df.at[idx, txg_c] = 'Yes'; continue
        if ihrp == 'ihrp' and rdt == 'pf' and oral1 == 'yes':
            df.at[idx, txg_c] = 'Yes (Old guideline)'; continue
        if ihrp == 'ihrp' and rdt == 'pf' and oral2 == 'yes':
            df.at[idx, txg_c] = 'Yes (New guideline)'; continue
        if ihrp == 'ihrp' and oral1 == 'yes' and oral2 == 'yes':
            df.at[idx, txg_c] = 'Yes'; continue

        # Non-IHRP with oral=No
        if ihrp == 'non-ihrp' and oral_non == 'no':
            if refer == 'yes':
                df.at[idx, txg_c] = 'Refer & No'
            elif preg == 'yes' and str(pq_tab).strip().lower() in ['', 'n/a', 'not treated', 'nan']:
                if rdt in ['pf','mix'] and (act_v and act_v>0) and is_blank(df.at[idx, cq_tab_c] if cq_tab_c in df.columns else ''):
                    df.at[idx, txg_c] = 'Yes'
                elif rdt in ['pv', 'pv/ non pf'] and (cq_v and cq_v>0) and is_blank(df.at[idx, act_tab_c] if act_tab_c in df.columns else ''):
                    df.at[idx, txg_c] = 'Yes'
                else:
                    df.at[idx, txg_c] = 'No'
            else:
                df.at[idx, txg_c] = 'No'
            continue

        # IHRP with both oral flags No
        if ihrp == 'ihrp' and oral1 == 'no' and oral2 == 'no':
            if refer == 'yes':
                df.at[idx, txg_c] = 'Refer & No'; continue
            elif preg == 'yes' and str(pq_tab).strip().lower() in ['', 'n/a', 'not treated', 'nan']:
                assigned = False
                if rdt == 'pf':
                    rules = [(0,0.99,3), (1,4,6), (5,9,12), (10,14,18), (15,100,24)]
                    for mn,mx,req in rules:
                        if age_v is not None and mn <= age_v <= mx and act_v == req and is_blank(df.at[idx, cq_tab_c] if cq_tab_c in df.columns else ''):
                            df.at[idx, txg_c] = 'Yes'; assigned=True; break
                elif rdt in ['pv','pv/ non pf']:
                    rules = [(0,0.99,1),(1,4,4),(5,9,5),(10,14,7.5),(15,100,10)]
                    for mn,mx,req in rules:
                        if age_v is not None and mn <= age_v <= mx and is_blank(df.at[idx, act_tab_c] if act_tab_c in df.columns else '') and cq_v == req:
                            df.at[idx, txg_c] = 'Yes'; assigned=True; break
                elif rdt == 'mix':
                    rules = [(0,0.99,3),(1,4,6),(5,9,12),(10,14,18),(15,100,24)]
                    for mn,mx,req in rules:
                        if age_v is not None and mn <= age_v <= mx and act_v == req and is_blank(df.at[idx, cq_tab_c] if cq_tab_c in df.columns else ''):
                            df.at[idx, txg_c] = 'Yes'; assigned=True; break
                if not assigned: df.at[idx, txg_c] = 'No'
            else:
                df.at[idx, txg_c] = 'No'
            continue

    # place TX_GUIDELINE just after ORAL_TX_2 (or before COMMENT fallback)
    try:
        cols = df.columns.tolist()
        if txg_c in cols: cols.remove(txg_c)
        pos = df.columns.get_loc('ORAL_TX_2')+1 if 'ORAL_TX_2' in df.columns else df.columns.get_loc('COMMENT')
        cols.insert(pos, txg_c); df = df[cols]
    except: pass

    # === REMARK/REASON check when TX not according to guideline ===
    remark_c = get_col('REMARK') or get_col('REASON')
    if remark_c:
        for idx in df.index:
            v = str(df.at[idx, txg_c]).strip().lower() if not pd.isna(df.at[idx, txg_c]) else ''
            r = str(df.at[idx, remark_c]).strip() if not pd.isna(df.at[idx, remark_c]) else ''
            if v in ['no','refer & no'] and r=='':
                if 'Reason for treatment not according to guideline' not in df.at[idx, 'COMMENT']:
                    add_comment(idx, "Reason for treatment not according to guideline")
    else:
        print("Column 'REMARK' (or 'REASON') not found. No additional check for guideline reason performed.")

    # Final tidy
    df['COMMENT'] = df['COMMENT'].astype(str).str.rstrip('; ').str.strip().fillna('')
    return df
