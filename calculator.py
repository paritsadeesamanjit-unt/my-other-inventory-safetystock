"""
calculator.py — คำนวณ Stock ที่แท้จริงโดยรวม No_SAP, Chemical, SAP_ZRMM0001
และหักด้วย PR & PO Update กรณีรับสินค้าแล้วแต่ยังไม่เข้าระบบ
"""
import pandas as pd
import numpy as np
from datetime import date


def compute_true_stock(sheets: dict) -> pd.DataFrame:
    """
    รวม stock จากทุกแหล่ง:
      - SAP_ZRMM0001  : Unrestricted + Inspection
      - No_SAP / Chemical : Remaining (วัสดุนอกระบบ)

    Returns DataFrame: Material_ID, Material_Name, Unit, Source,
                       SAP_Stock, NoSAP_Stock, Total_Stock
    """
    rows = []

    # SAP stock
    sap = sheets.get("sap_stock", pd.DataFrame())
    if not sap.empty:
        for _, r in sap.iterrows():
            rows.append({
                "Material_ID": str(r.get("Material_ID", "")).strip(),
                "Material_Name": str(r.get("Material_Desc", "")).strip(),
                "Unit": str(r.get("Unit", "")).strip(),
                "Source": "SAP",
                "SAP_Unrestricted": float(r.get("Unrestricted", 0) or 0),
                "SAP_Inspection": float(r.get("Inspection", 0) or 0),
                "SAP_InTransit": float(r.get("In_Transit", 0) or 0),
                "NoSAP_Remaining": 0.0,
                "Material_Type": str(r.get("Material_Type", "")),
                "Material_Group": str(r.get("Material_Group", "")),
            })

    # No_SAP stock
    no_sap = sheets.get("no_sap", pd.DataFrame())
    if not no_sap.empty:
        for _, r in no_sap.iterrows():
            rows.append({
                "Material_ID": str(r.get("Material_ID", "")).strip(),
                "Material_Name": str(r.get("Material_Name", "")).strip(),
                "Unit": str(r.get("Unit", "")).strip(),
                "Source": "No_SAP",
                "SAP_Unrestricted": 0.0,
                "SAP_Inspection": 0.0,
                "SAP_InTransit": 0.0,
                "NoSAP_Remaining": float(r.get("Remaining", 0) or 0),
                "Material_Type": str(r.get("Material_Type", "")),
                "Material_Group": "",
            })

    # Chemical (No SAP) stock
    chem = sheets.get("chemical", pd.DataFrame())
    if not chem.empty:
        for _, r in chem.iterrows():
            rows.append({
                "Material_ID": str(r.get("Material_ID", "")).strip(),
                "Material_Name": str(r.get("Material_Name", "")).strip(),
                "Unit": str(r.get("Unit", "")).strip(),
                "Source": "Chemical_NoSAP",
                "SAP_Unrestricted": 0.0,
                "SAP_Inspection": 0.0,
                "SAP_InTransit": 0.0,
                "NoSAP_Remaining": float(r.get("Remaining", 0) or 0),
                "Material_Type": str(r.get("Material_Type", "")),
                "Material_Group": "",
            })

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)

    # Aggregate by Material_ID (รวม SAP + NoSAP)
    agg = df.groupby("Material_ID", as_index=False).agg(
        Material_Name=("Material_Name", "first"),
        Unit=("Unit", "first"),
        Sources=("Source", lambda x: "+".join(sorted(set(x)))),
        SAP_Unrestricted=("SAP_Unrestricted", "sum"),
        SAP_Inspection=("SAP_Inspection", "sum"),
        SAP_InTransit=("SAP_InTransit", "sum"),
        NoSAP_Remaining=("NoSAP_Remaining", "sum"),
        Material_Type=("Material_Type", "first"),
        Material_Group=("Material_Group", "first"),
    )
    agg["Total_Stock"] = (
        agg["SAP_Unrestricted"] + agg["SAP_Inspection"] + agg["NoSAP_Remaining"]
    )
    return agg


def compute_pr_po_adjusted(sheets: dict) -> pd.DataFrame:
    """
    คำนวณ Outstanding PO/PR ที่แท้จริง:
    - เริ่มจาก SAP_ZRMM0004 (Outstanding_QTY)
    - ลบด้วย PR & PO Update ที่ได้รับสินค้าแล้วแต่รับเข้าระบบไม่ได้
      (Received_QTY > 0)
    """
    sap_po = sheets.get("sap_pr_po", pd.DataFrame())
    update = sheets.get("pr_po_update", pd.DataFrame())

    if sap_po.empty:
        return pd.DataFrame()

    df = sap_po.copy()
    df["Outstanding_QTY"] = pd.to_numeric(df.get("Outstanding_QTY", 0), errors="coerce").fillna(0)
    df["Adjusted_Outstanding"] = df["Outstanding_QTY"].copy()

    if not update.empty:
        # วัสดุที่รับสินค้าแล้วจาก PR&PO Update
        received = update[update["Received_QTY"].notna() & (update["Received_QTY"] > 0)].copy()
        received["Received_QTY"] = pd.to_numeric(received["Received_QTY"], errors="coerce").fillna(0)

        for _, upd_row in received.iterrows():
            mat = str(upd_row.get("Material_ID", "")).strip()
            pr_num = str(upd_row.get("PR_Number", "")).strip()
            rcv_qty = float(upd_row.get("Received_QTY", 0))

            mask = df["Material_ID"].astype(str).str.strip() == mat
            if pr_num and pr_num != "nan":
                pr_mask = df["PR_Number"].astype(str).str.strip() == pr_num
                combined = mask & pr_mask
                if combined.any():
                    df.loc[combined, "Adjusted_Outstanding"] = (
                        df.loc[combined, "Adjusted_Outstanding"] - rcv_qty
                    ).clip(lower=0)
                    continue
            # fallback: match by material only
            if mask.any():
                df.loc[mask, "Adjusted_Outstanding"] = (
                    df.loc[mask, "Adjusted_Outstanding"] - rcv_qty
                ).clip(lower=0)

    df["Adjusted_Outstanding"] = df["Adjusted_Outstanding"].clip(lower=0)
    return df


def compute_usage_summary(sheets: dict) -> pd.DataFrame:
    """
    รวม usage จาก SAP_MB51 + Manual_Summary
    Returns: Material_ID, Material_Name, Unit, SAP_Usage, Manual_Usage, Total_Usage
    """
    mb51 = sheets.get("mb51", pd.DataFrame())
    manual = sheets.get("manual", pd.DataFrame())

    rows = []

    if not mb51.empty:
        mb_clean = mb51.copy()
        mb_clean["Quantity"] = pd.to_numeric(mb_clean["Quantity"], errors="coerce").fillna(0)
        # เอาเฉพาะ movement type ที่เป็นการเบิก (ค่าลบ = ออก)
        mb_clean["Usage"] = mb_clean["Quantity"].abs()
        sap_usage = mb_clean.groupby("Material_ID", as_index=False).agg(
            Material_Name=("Material_Desc", "first"),
            Unit=("Unit", "first"),
            SAP_Usage=("Usage", "sum"),
        )
        rows.append(sap_usage.assign(Manual_Usage=0.0))

    if not manual.empty:
        manual_clean = manual.copy()
        manual_clean["Withdrawn_QTY"] = pd.to_numeric(
            manual_clean["Withdrawn_QTY"], errors="coerce"
        ).fillna(0)
        man_usage = manual_clean.groupby("Material_ID", as_index=False).agg(
            Material_Name=("Material_Name", "first"),
            Unit=("Unit", "first"),
            Manual_Usage=("Withdrawn_QTY", "sum"),
        )
        rows.append(man_usage.assign(SAP_Usage=0.0))

    if not rows:
        return pd.DataFrame()

    combined = pd.concat(rows, ignore_index=True)
    result = combined.groupby("Material_ID", as_index=False).agg(
        Material_Name=("Material_Name", "first"),
        Unit=("Unit", "first"),
        SAP_Usage=("SAP_Usage", "sum"),
        Manual_Usage=("Manual_Usage", "sum"),
    )
    result["Total_Usage"] = result["SAP_Usage"] + result["Manual_Usage"]
    return result.sort_values("Total_Usage", ascending=False).reset_index(drop=True)


def compute_expiry_alerts(sheets: dict, days_ahead: int = 90) -> pd.DataFrame:
    """
    หาวัสดุที่ใกล้หมดอายุ (No_SAP + Chemical)
    """
    today = pd.Timestamp(date.today())
    alert_date = today + pd.Timedelta(days=days_ahead)

    frames = []
    for key, label in [("no_sap", "No_SAP"), ("chemical", "Chemical_NoSAP")]:
        df = sheets.get(key, pd.DataFrame())
        if df.empty or "Expiry_Date" not in df.columns:
            continue
        df2 = df[df["Expiry_Date"].notna()].copy()
        df2 = df2[df2["Expiry_Date"] <= alert_date]
        df2["Category"] = label
        df2["Days_Until_Expiry"] = (df2["Expiry_Date"] - today).dt.days
        frames.append(df2[["Material_ID", "Material_Name", "Unit", "Remaining",
                             "Expiry_Date", "Days_Until_Expiry", "Category"]])

    if not frames:
        return pd.DataFrame()

    result = pd.concat(frames, ignore_index=True)
    result = result.sort_values("Days_Until_Expiry").reset_index(drop=True)
    result["Urgency"] = result["Days_Until_Expiry"].apply(
        lambda d: "🔴 หมดอายุแล้ว" if d < 0
        else ("🟠 ด่วน (<30 วัน)" if d < 30
              else ("🟡 ใกล้หมด (<90 วัน)" if d < 90 else "🟢 ปกติ"))
    )
    return result


def compute_stock_alerts(sheets: dict) -> pd.DataFrame:
    """
    เปรียบ Stock ใน SAP กับ Safety Stock จาก Stock Material sheet
    """
    stock_master = sheets.get("stock", pd.DataFrame())
    sap_stock = sheets.get("sap_stock", pd.DataFrame())

    if stock_master.empty or sap_stock.empty:
        return pd.DataFrame()

    master = stock_master[["Material_ID", "Specification", "Safety_Stock",
                            "Warehouse_Stock", "Daily_Consumption", "Lead_Time"]].copy()
    master["Safety_Stock"] = pd.to_numeric(master["Safety_Stock"], errors="coerce").fillna(0)

    sap = sap_stock[["Material_ID", "Unrestricted", "Inspection"]].copy()
    sap["Unrestricted"] = pd.to_numeric(sap["Unrestricted"], errors="coerce").fillna(0)
    sap["Inspection"] = pd.to_numeric(sap["Inspection"], errors="coerce").fillna(0)
    sap["SAP_Available"] = sap["Unrestricted"] + sap["Inspection"]

    merged = master.merge(sap, on="Material_ID", how="left")
    merged["SAP_Available"] = merged["SAP_Available"].fillna(0)
    merged["Stock_Gap"] = merged["SAP_Available"] - merged["Safety_Stock"]

    alerts = merged[merged["SAP_Available"] < merged["Safety_Stock"]].copy()
    alerts["Alert_Level"] = alerts["Stock_Gap"].apply(
        lambda g: "🔴 Critical (0)" if g <= -merged["Safety_Stock"].max() * 0.5
        else ("🟠 Low" if g < 0 else "🟢 OK")
    )
    return alerts.sort_values("Stock_Gap").reset_index(drop=True)
