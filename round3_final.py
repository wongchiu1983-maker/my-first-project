"""
HK IPO Full Analysis Report 更新腳本
----------------------------------------
讀取 HK_IPO_Full_Analysis_Report_v2.xlsx 的 "IPO全量數據分析" 工作表，
根據 Stock Code 精準更新指定欄位的數據，並保留原工作表的格式。
"""

import pandas as pd
from openpyxl import load_workbook

# ------------------------------------------------------------------
# 0. 文件路徑與工作表名稱設定
# ------------------------------------------------------------------
INPUT_FILE = "HK_IPO_Full_Analysis_Report_v2.xlsx"
OUTPUT_FILE = "HK_IPO_Full_Analysis_Report_v2.xlsx"  # 原地覆寫；如需另存請修改此路徑
SHEET_NAME = "IPO全量數據分析"

# ------------------------------------------------------------------
# 1. 欄位名稱 → 欄位索引（1-based，對應 openpyxl）
#    依據用戶提供的嚴格欄位順序
# ------------------------------------------------------------------
COLUMNS = [
    "Stock Code",                        # A  (1)
    "Company Name",                      # B  (2)
    "Date of Prospectus",                # C  (3)
    "Date of Listing",                   # D  (4)
    "IPO Subscription Price",            # E  (5)
    "公開零售認購倍數",                   # F  (6)
    "國際配售認購倍數",                   # G  (7)
    "國際配售認購人數",                   # H  (8)
    "基石投資者占比 (%)",                 # I  (9)
    "公司大股東持股占比 (上市後 %)",        # J  (10)
    "暗盤收盤回報率 (%)",                 # K  (11)
    "首日收市回報率 (%)",                 # L  (12)
    "一個月回報率 (%)",                   # M  (13)
    "一手中簽率 (%)",                     # N  (14)
]
COL_IDX = {name: i + 1 for i, name in enumerate(COLUMNS)}  # 1-based

# ------------------------------------------------------------------
# 2. 更新資料定義（以 Stock Code 為鍵）
#    僅填入已確認的真實數據；未列出的欄位保持原值「待公告/需手動核查」
# ------------------------------------------------------------------
UPDATES = {
    "06681": {  # BrainAurora Medical Technology Limited - B
        "公開零售認購倍數": "11.39倍",
        "國際配售認購倍數": "接近1倍",
        "首日收市回報率 (%)": "+3% ~ +5%(約)",
        "一手中簽率 (%)": "約50.01%",
    },
    "00100": {  # MiniMax Group Inc. - W - P
        "公開零售認購倍數": "1837.17倍",
        "國際配售認購倍數": "36.76倍",
        "國際配售認購人數": "482",
        "基石投資者占比 (%)": "56.53%",
        "公司大股東持股占比 (上市後 %)": "89.30%",
        "暗盤收盤回報率 (%)": "+26.80%",
        "首日收市回報率 (%)": "+109.09%",
        "一個月回報率 (%)": "+185.50%",
        "一手中簽率 (%)": "2.81%",
    },
    "02560": {  # Anhui Conch Material Technology Co., Ltd. - H shares
        "公開零售認購倍數": "相對溫和",
        "首日收市回報率 (%)": "約-47%",
    },
}


def main():
    # --------------------------------------------------------------
    # 3. 讀取 Excel（使用 openpyxl 以保留格式）
    # --------------------------------------------------------------
    print(f"[1/4] 讀取文件: {INPUT_FILE}")
    wb = load_workbook(INPUT_FILE)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(
            f"找不到工作表 '{SHEET_NAME}'，可用工作表為: {wb.sheetnames}"
        )
    ws = wb[SHEET_NAME]
    print(f"      工作表尺寸: {ws.max_row} 行 x {ws.max_column} 欄")

    # --------------------------------------------------------------
    # 4. 驗證標題欄（確保欄位順序與預期一致）
    # --------------------------------------------------------------
    print("[2/4] 驗證欄位標題")
    actual_headers = [ws.cell(row=1, column=i + 1).value for i in range(len(COLUMNS))]
    for expected, actual in zip(COLUMNS, actual_headers):
        if expected != actual:
            raise ValueError(
                f"欄位標題不符: 預期 '{expected}' / 實際 '{actual}'"
            )
    print("      欄位標題與預期完全一致 ✓")

    # --------------------------------------------------------------
    # 5. 建立 Stock Code → 行號索引
    # --------------------------------------------------------------
    stock_code_col = COL_IDX["Stock Code"]
    code_to_row = {}
    for r in range(2, ws.max_row + 1):
        code_val = ws.cell(row=r, column=stock_code_col).value
        if code_val is not None:
            code_to_row[str(code_val).strip()] = r

    # --------------------------------------------------------------
    # 6. 根據 UPDATES 定義，精準寫入指定儲存格
    # --------------------------------------------------------------
    print("[3/4] 更新資料")
    total_cells_updated = 0
    for stock_code, fields in UPDATES.items():
        if stock_code not in code_to_row:
            print(f"      ⚠ 找不到 Stock Code: {stock_code}，跳過")
            continue
        row = code_to_row[stock_code]
        company_name = ws.cell(row=row, column=COL_IDX["Company Name"]).value
        company_short = (company_name or "").replace("\n", " ").strip()[:45]
        print(f"      • {stock_code} (Row {row}) - {company_short}")
        for field, new_value in fields.items():
            col = COL_IDX[field]
            old_value = ws.cell(row=row, column=col).value
            ws.cell(row=row, column=col).value = new_value
            total_cells_updated += 1
            print(f"          [{field}] '{old_value}' → '{new_value}'")
    print(f"      共更新 {total_cells_updated} 個儲存格")

    # --------------------------------------------------------------
    # 7. 儲存
    # --------------------------------------------------------------
    print(f"[4/4] 儲存至: {OUTPUT_FILE}")
    wb.save(OUTPUT_FILE)
    print("      儲存完成 ✓")

    # --------------------------------------------------------------
    # 8. 使用 pandas 驗證結果
    # --------------------------------------------------------------
    print("\n=== 驗證更新後的資料 (使用 pandas 讀取) ===")
    df = pd.read_excel(OUTPUT_FILE, sheet_name=SHEET_NAME, dtype=str)
    for stock_code in UPDATES:
        row_df = df[df["Stock Code"] == stock_code]
        if row_df.empty:
            print(f"  ⚠ 驗證時找不到 {stock_code}")
            continue
        print(f"\n  [{stock_code}] {row_df.iloc[0]['Company Name'].strip()}")
        for field in UPDATES[stock_code]:
            print(f"      {field}: {row_df.iloc[0][field]}")


if __name__ == "__main__":
    main()
