# Publication example: personal paths and shipment outputs removed.
# Run in IPython/Jupyter for styled table display.
import pandas as pd
from IPython.display import display
import glob
import os
import time

start_time = time.time()

# --- Invoice file (fixed path) ---
file_path_inv = r"data/Invoice Excel.xls"

# --- Draft file (latest file in folder) ---
folder = r"data/"

files = glob.glob(os.path.join(folder, "*.xlsx"))  # ✅ แก้ตรงนี้

if not files:
    print("ไม่พบไฟล์ในโฟลเดอร์:", folder)
else:
    file_path_draft = max(files, key=os.path.getmtime)
    print("Reading draft file:", file_path_draft)

    # --- Read raw file to extract row 9 ---
    df_raw = pd.read_csv(file_path_inv, encoding='utf-16', sep=None,
                         engine='python', header=None, on_bad_lines='skip')

    col_name_A = df_raw.iloc[8, 0]
    col_val_A  = df_raw.iloc[8, 1]

    print(f"New columns: '{col_name_A}' = {col_val_A}")

    # --- Read Invoice data from row 13 (header=12) ---
    df_inv   = pd.read_csv(file_path_inv, encoding='utf-16', sep=None,
                           engine='python', header=12, on_bad_lines='skip')
    df_draft = pd.read_excel(file_path_draft, engine='openpyxl')

    # --- Filter rows where first column is not NaN ---
    first_col_inv   = df_inv.columns[0]
    first_col_draft = df_draft.columns[0]

    df_inv   = df_inv[df_inv[first_col_inv].notna()].reset_index(drop=True)
    df_draft = df_draft[df_draft[first_col_draft].notna()].reset_index(drop=True)

    # --- Convert numeric columns ใน df_inv ให้มีทศนิยม 1 ตำแหน่ง ---
    for col in df_inv.columns:
        converted = pd.to_numeric(df_inv[col], errors='coerce')
        if converted.notna().sum() > len(df_inv) * 0.5:
            df_inv[col] = converted.round(1)

    # --- เพิ่ม column ใหม่ต่อท้าย ---
    df_inv[col_name_A] = col_val_A

    # --- Mapping: df_draft column -> df_inv column ---
    col_mapping = {
        'เลขที่อินวอยซ์'                  : 'Invoice No.',
        'Part No'                     : 'รหัสสินค้า',
        'Brand'                       : 'Brand',
        'ชื่อสินค้าภาษาอังกฤษ 1'          : 'Description',
        'ชื่อสินค้าภาษาอังกฤษ 2'          : 'Size',
        'ชื่อสินค้าภาษาไทย 1'            : 'Thai Description',
        'พิกัด'                        : 'Tariff Code',
        'ประเภทพิกัด'                   : 'Privilege',
        'ปริมาณ (Invoice)'             : 'Qty',
        'หน่วยปริมาณ (Invoice)'        : 'Unit',
        'น้ำหนักสุทธิ'                   : 'Total G.W.(Kgs)',
        'Unit price (เงินต่างประเทศ)'   : 'Unit Price',
        'รหัสสกุลเงิน'                  : 'Currency',
        'ชื่อเรือ'                      : 'VESSEL NAME',
        'รหัสประเทศกำเนิด'             : 'ORIGIN',
        'เลขที่ใบอนุญาต1'              : 'เลขรับ/เลขอนุมัติหลักการ'
    }

    # --- Unit equivalents (PC = C62) ---
    unit_equivalents = {
        'c62': 'pc',
        'pc' : 'c62',
    }

    # --- normalize: ลบ space + lowercase + แปลงตัวเลขให้เป็น float format ---
    def normalize(val):
        s = str(val).strip().lower().replace(' ', '')
        try:
            s = str(float(s))
        except:
            pass
        return s

    # --- เปรียบเทียบโดยคำนึง unit equivalents ---
    def is_equal(cmp_draft, cmp_inv, draft_col):
        if cmp_draft == cmp_inv:
            return True
        if draft_col == 'หน่วยปริมาณ (Invoice)':
            if unit_equivalents.get(cmp_draft) == cmp_inv:
                return True
            if unit_equivalents.get(cmp_inv) == cmp_draft:
                return True
        return False

    # --- เช็ค row ต่อ row ตามลำดับ ---
    results = []
    total_rows = min(len(df_draft), len(df_inv))

    for i in range(total_rows):
        row_draft = df_draft.iloc[i]
        row_inv   = df_inv.iloc[i]

        for draft_col, inv_col in col_mapping.items():
            val_draft = str(row_draft.get(draft_col, '')).strip()
            val_inv   = str(row_inv.get(inv_col, '')).strip()

            cmp_draft = normalize(val_draft)
            cmp_inv   = normalize(val_inv)

            status = '✅' if is_equal(cmp_draft, cmp_inv, draft_col) else '⚠️ ไม่ตรง'

            results.append({
                'Row'      : i + 1,
                'Field'    : draft_col,
                'df_draft' : val_draft,
                'df_inv'   : val_inv,
                'Status'   : status,
            })

    df_result = pd.DataFrame(results)

    # --- แสดงเฉพาะที่ไม่ตรง ---
    df_mismatch = df_result[df_result['Status'] == '⚠️ ไม่ตรง'].reset_index(drop=True)

    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)

    if df_mismatch.empty:
        print("✅ ข้อมูลทุก field ตรงกันหมด!")
    else:
        print(f"⚠️ พบข้อมูลไม่ตรงกัน {len(df_mismatch)} รายการ\n")
        display(df_mismatch.style
            .applymap(lambda v: 'background-color: #ffe0e0' if v == '⚠️ ไม่ตรง' else '')
            .set_properties(**{'text-align': 'left'})
        )

    # --- แจ้งถ้าจำนวน row ไม่เท่ากัน ---
    if len(df_draft) != len(df_inv):
        print(f"\n⚠️ จำนวน row ไม่เท่ากัน: df_draft={len(df_draft)} rows, df_inv={len(df_inv)} rows")

# --- Timestamp ---
elapsed = time.time() - start_time
print(f"\n⏱️ ใช้เวลา run ทั้งหมด: {int(elapsed//60)}m {elapsed%60:.2f}s")