"""
data_loader.py — อ่านและ normalize ทุก sheet จากไฟล์หลัก
"""
import pandas as pd
import numpy as np
from io import BytesIO


# ── column-name maps ──────────────────────────────────────────────────────────
NO_SAP_COLS = {
    "NO": "NO",
    "料號\nเลขวัสดุ": "Material_ID",
    "品名\nชื่อวัสดุ": "Material_Name",
    "數量\nจำนวน": "Quantity",
    "單位\nหน่วย": "Unit",
    "製成單位\nหน่วยงาน": "Dept",
    "供應商通知日期\nวันที่ทางซัพพลายเออร์แจ้ง": "Supplier_Notice_Date",
    "收貨日期\nวันที่รับเข้า": "Receive_Date",
    "製造日期\nวันที่ผลิต": "Mfg_Date",
    "到期日\nวันที่หมดอายุ": "Expiry_Date",
    "已領取的總數量\nจำนวนที่เบิกแล้วทั้งหมด": "Total_Withdrawn",
    "其餘的\nคงเหลือ": "Remaining",
    "วัสดุ": "Material_Type",
    "備註\nหมายเหตุ": "Remark",
}

SAP_ZRMM0001_COLS = {
    "Material Type": "Material_Type",
    "Material Group": "Material_Group",
    "MRP Controller": "MRP_Controller",
    "Plant": "Plant",
    "Stor. Location": "Storage_Location",
    "Stor. Loc. Desc": "Storage_Loc_Desc",
    "Storage Cell": "Storage_Cell",
    "Material Num": "Material_ID",
    "Material Desc": "Material_Desc",
    "Unit": "Unit",
    "Unrestricted": "Unrestricted",
    "Inspection": "Inspection",
    "In Transit": "In_Transit",
    "Blocked": "Blocked",
}

SAP_ZRMM0004_COLS = {
    "PR Number": "PR_Number",
    "PO Number": "PO_Number",
    "PR Item": "PR_Item",
    "Mat. Number": "Material_ID",
    "Product Spec.": "Product_Spec",
    "Unit": "Unit",
    "Vendor": "Vendor",
    "Storage Loca.": "Storage_Location",
    "PR Date": "PR_Date",
    "Require Date": "Require_Date",
    "Delivery Date": "Delivery_Date",
    "Ordered QTY": "Ordered_QTY",
    "Required QTY": "Required_QTY",
    "Accu. Rece.QTY": "Accu_Rece_QTY",
    "QTY AwaitInsp": "QTY_AwaitInsp",
    "OutstandigQTY": "Outstanding_QTY",
    "Item Remark": "Item_Remark",
    "Mat. Group": "Mat_Group",
    "Mat. Grp. Descr.": "Mat_Grp_Desc",
    "PR Requester": "PR_Requester",
    "MRP Controller": "MRP_Controller",
    "PR Del. Indicator": "PR_Del_Indicator",
    "PO Del. Indicator": "PO_Del_Indicator",
}

PR_PO_UPDATE_COLS = {
    "申請日期\nวันที่ขอซื้อ (PR Date)": "PR_Date",
    "PR單號\nเลขที่ PR": "PR_Number",
    "\nPR項次ลำดับ PR": "PR_Item",
    "PO單號\nเลขที่ PO": "PO_Number",
    "物料編號\nรหัสวัสดุ": "Material_ID",
    "物料名稱／規格\nรายละเอียดสินค้า / Spec": "Material_Name",
    "請購數量จำนวนที่ขอซื้อ": "PR_QTY",
    "單位หน่วย": "Unit",
    "供應商\nผู้ขาย (Vendor)": "Vendor",
    "需求日期\nวันที่ต้องการ": "Require_Date",
    "入庫日期\nวันที่รับเข้า": "Receive_Date",
    "已收貨\nรับสินค้าแล้ว": "Received_QTY",
    "待收數量\nค้างรับ (Outstanding)": "Outstanding_QTY",
    "REMARK": "Remark",
}

SAP_MB51_COLS = {
    "Storage location": "Storage_Location",
    "Movement type": "Movement_Type",
    "Material": "Material_ID",
    "Material description": "Material_Desc",
    "Movement Type Text": "Movement_Type_Text",
    "Special Stock": "Special_Stock",
    "Material Document": "Material_Document",
    "Material Doc.Item": "Doc_Item",
    "Qty in unit of entry": "Quantity",
    "Unit of Entry": "Unit",
    "Document Date": "Document_Date",
    "Posting Date": "Posting_Date",
    "Entry Date": "Entry_Date",
    "Time of Entry": "Time_Entry",
    "Batch": "Batch",
    "Purchase order": "PO_Number",
    "Vendor": "Vendor",
    "Document Header Text": "Doc_Header_Text",
    "User Name": "User_Name",
}

MANUAL_SUMMARY_COLS = {
    "料號 เลขวัสดุ": "Material_ID",
    "品名 ชื่อวัสดุ": "Material_Name",
    "單位 หน่วย": "Unit",
    "領用數量 จำนวนที่เบิก": "Withdrawn_QTY",
}

STOCK_MATERIAL_COLS = {
    "Material Code": "Material_ID",
    "Specification": "Specification",
    "รายละเอียด": "Description_TH",
    "Unit": "Unit",
    "Category": "Category",
    "Process": "Process",
    "Safety Stock": "Safety_Stock",
    "Warehouse Stock": "Warehouse_Stock",
    "Remind to buy": "Remind_To_Buy",
    "PR (1)": "PR_1",
    "Delivery Date": "Delivery_Date_1",
    "PR (2)": "PR_2",
    "Delivery Date.1": "Delivery_Date_2",
    "LT": "Lead_Time",
    "Periodic Consumption": "Periodic_Consumption",
    "Daily Consumption": "Daily_Consumption",
    "Weekly Consumption": "Weekly_Consumption",
    "DOH": "DOH",
    "Days Remaining": "Days_Remaining",
    "Out-of-Stock Date": "Out_of_Stock_Date",
    "Total Usage Time": "Total_Usage_Time",
    "Remarks": "Remarks",
}


def _clean_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def load_all_sheets(file) -> dict:
    """
    file: UploadedFile จาก st.file_uploader หรือ path string
    Returns dict ของ DataFrames ที่ clean แล้ว
    """
    if hasattr(file, "read"):
        raw = file.read()
        xf = pd.ExcelFile(BytesIO(raw))
    else:
        xf = pd.ExcelFile(file)

    sheets = {}

    # Stock Material (header row 2, index 2)
    try:
        df = pd.read_excel(xf, sheet_name="Stock Material", header=2)
        df = df.rename(columns={c: STOCK_MATERIAL_COLS.get(c, c) for c in df.columns})
        df = _clean_df(df)
        df = df[df["Material_ID"].notna() & (df["Material_ID"] != "Material Code")]
        sheets["stock"] = df
    except Exception as e:
        sheets["stock"] = pd.DataFrame()

    # No_SAP
    try:
        df = pd.read_excel(xf, sheet_name="No_SAP", header=0)
        df = df.rename(columns={c: NO_SAP_COLS.get(c, c) for c in df.columns})
        df = _clean_df(df)
        df = df[df["Material_ID"].notna() & (df["NO"].notna())]
        for col in ["Supplier_Notice_Date", "Receive_Date", "Mfg_Date", "Expiry_Date"]:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
        sheets["no_sap"] = df
    except Exception as e:
        sheets["no_sap"] = pd.DataFrame()

    # Chemical
    try:
        df = pd.read_excel(xf, sheet_name="Chemical", header=0)
        df = df.rename(columns={c: NO_SAP_COLS.get(c, c) for c in df.columns})
        df = _clean_df(df)
        df = df[df["Material_ID"].notna() & (df["NO"].notna())]
        for col in ["Supplier_Notice_Date", "Receive_Date", "Mfg_Date", "Expiry_Date"]:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
        sheets["chemical"] = df
    except Exception as e:
        sheets["chemical"] = pd.DataFrame()

    # SAP_ZRMM0001
    try:
        df = pd.read_excel(xf, sheet_name="SAP_ZRMM0001", header=0)
        df = df.rename(columns={c: SAP_ZRMM0001_COLS.get(c, c) for c in df.columns})
        df = _clean_df(df)
        df = df[df["Material_ID"].notna()]
        sheets["sap_stock"] = df
    except Exception as e:
        sheets["sap_stock"] = pd.DataFrame()

    # SAP_ZRMM0004
    try:
        df = pd.read_excel(xf, sheet_name="SAP_ZRMM0004", header=0)
        df = df.rename(columns={c: SAP_ZRMM0004_COLS.get(c, c) for c in df.columns})
        df = _clean_df(df)
        df = df[df["Material_ID"].notna()]
        for col in ["PR_Date", "Require_Date", "Delivery_Date"]:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")
        sheets["sap_pr_po"] = df
    except Exception as e:
        sheets["sap_pr_po"] = pd.DataFrame()

    # PR & PO Update
    try:
        df = pd.read_excel(xf, sheet_name="PR & PO Update", header=0)
        df = df.rename(columns={c: PR_PO_UPDATE_COLS.get(c, c) for c in df.columns})
        df = _clean_df(df)
        df = df[df["Material_ID"].notna()]
        for col in ["PR_Date", "Require_Date", "Receive_Date"]:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")
        sheets["pr_po_update"] = df
    except Exception as e:
        sheets["pr_po_update"] = pd.DataFrame()

    # SAP_MB51
    try:
        df = pd.read_excel(xf, sheet_name="SAP_MB51", header=0)
        df = df.rename(columns={c: SAP_MB51_COLS.get(c, c) for c in df.columns})
        df = _clean_df(df)
        df = df[df["Material_ID"].notna()]
        for col in ["Document_Date", "Posting_Date", "Entry_Date"]:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")
        sheets["mb51"] = df
    except Exception as e:
        sheets["mb51"] = pd.DataFrame()

    # Manual_Summary
    try:
        df = pd.read_excel(xf, sheet_name="Manual_Summary", header=0)
        df = df.rename(columns={c: MANUAL_SUMMARY_COLS.get(c, c) for c in df.columns})
        df = _clean_df(df)
        df = df[df["Material_ID"].notna()]
        sheets["manual"] = df
    except Exception as e:
        sheets["manual"] = pd.DataFrame()

    return sheets
