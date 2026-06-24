"""
app.py — Material & Chemical Management System (TH Plant)
"""
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import date, datetime
import io

from data_loader import load_all_sheets
from calculator import (
    compute_true_stock,
    compute_pr_po_adjusted,
    compute_usage_summary,
    compute_expiry_alerts,
    compute_stock_alerts,
)
from export_utils import export_full_report

# ─── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TH Plant — Material & Chemical System",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Main background */
.main { background-color: #F0F4F8; }

/* KPI card */
.kpi-card {
    background: white;
    border-radius: 12px;
    padding: 16px 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    border-left: 5px solid #1A5276;
    margin-bottom: 8px;
}
.kpi-value { font-size: 2rem; font-weight: 700; color: #1A5276; line-height:1.1; }
.kpi-label { font-size: 0.8rem; color: #666; text-transform: uppercase; letter-spacing:.05em; }
.kpi-sub { font-size: 0.75rem; color: #999; margin-top:2px; }

/* Alert cards */
.alert-red   { border-left-color: #C0392B !important; }
.alert-orange{ border-left-color: #E67E22 !important; }
.alert-green { border-left-color: #1E8449 !important; }
.alert-teal  { border-left-color: #1A7A6E !important; }

/* Section header */
.section-header {
    background: linear-gradient(90deg, #1A5276 0%, #1A7A6E 100%);
    color: white; border-radius: 8px;
    padding: 10px 18px; margin-bottom: 12px; margin-top: 8px;
    font-size: 1rem; font-weight: 600;
}

/* Dataframe styling */
.dataframe-container { border-radius: 8px; overflow: hidden; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] { gap: 4px; }
.stTabs [data-baseweb="tab"] {
    border-radius: 6px 6px 0 0;
    font-weight: 500;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1C2833 0%, #1A5276 100%);
}
section[data-testid="stSidebar"] * { color: white !important; }
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stSlider label { color: #BFC9CA !important; }

/* Upload area */
.upload-area {
    border: 2px dashed #1A5276; border-radius: 12px;
    padding: 30px; text-align: center; background: white;
    margin: 20px 0;
}
</style>
""", unsafe_allow_html=True)


# ─── Helpers ─────────────────────────────────────────────────────────────────
def kpi(label: str, value, sub: str = "", color_class: str = "") -> str:
    return f"""<div class="kpi-card {color_class}">
  <div class="kpi-value">{value}</div>
  <div class="kpi-label">{label}</div>
  {"<div class='kpi-sub'>" + sub + "</div>" if sub else ""}
</div>"""


def fmt_num(n, decimals=0) -> str:
    if pd.isna(n):
        return "—"
    try:
        if decimals == 0:
            return f"{int(n):,}"
        return f"{n:,.{decimals}f}"
    except Exception:
        return str(n)


def section(title: str):
    st.markdown(f'<div class="section-header">▶ {title}</div>', unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def load_cached(file_bytes: bytes) -> dict:
    return load_all_sheets(io.BytesIO(file_bytes))


# ─── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏭 TH Plant")
    st.markdown("### Material & Chemical\nManagement System")
    st.divider()

    uploaded = st.file_uploader(
        "📎 อัปโหลดไฟล์หลัก (.xlsx)",
        type=["xlsx"],
        help="ไฟล์ที่ประกอบด้วย Sheet: No_SAP, Chemical, SAP_ZRMM0001, SAP_ZRMM0004, PR & PO Update, SAP_MB51, Manual_Summary",
    )

    st.divider()
    expiry_days = st.slider("⏰ แจ้งเตือนใกล้หมดอายุ (วัน)", 30, 180, 90, 10)
    st.divider()

    st.markdown("**📋 Sheet ที่ต้องการ:**")
    st.markdown("""
- No_SAP  
- Chemical  
- SAP_ZRMM0001  
- SAP_ZRMM0004  
- PR & PO Update  
- SAP_MB51  
- Manual_Summary  
""")
    st.divider()
    st.caption(f"วันที่: {date.today().strftime('%d %b %Y')}")


# ─── Main ─────────────────────────────────────────────────────────────────────
st.markdown("# 🏭 TH Plant — ระบบคำนวณวัสดุและสารเคมี")

if not uploaded:
    st.markdown("""
    <div class="upload-area">
        <h3>📂 กรุณาอัปโหลดไฟล์หลัก</h3>
        <p style="color:#666;">ลากไฟล์ .xlsx มาวางที่นี่ หรือคลิกปุ่มอัปโหลดในแถบซ้าย</p>
        <p style="font-size:0.85rem; color:#999;">รองรับ Sheet: No_SAP · Chemical · SAP_ZRMM0001 · SAP_ZRMM0004 · PR &amp; PO Update · SAP_MB51 · Manual_Summary</p>
    </div>
    """, unsafe_allow_html=True)

    st.info("💡 **วิธีใช้:** อัปโหลดไฟล์ Excel หลักที่มีข้อมูลครบทุก Sheet แล้วระบบจะคำนวณและแสดงผลโดยอัตโนมัติ")
    st.stop()


# ─── Load & Compute ──────────────────────────────────────────────────────────
with st.spinner("⏳ กำลังโหลดและประมวลผลข้อมูล..."):
    file_bytes = uploaded.read()
    try:
        sheets = load_cached(file_bytes)
    except Exception as e:
        st.error(f"❌ ไม่สามารถอ่านไฟล์ได้: {e}")
        st.stop()

    true_stock    = compute_true_stock(sheets)
    pr_po_adj     = compute_pr_po_adjusted(sheets)
    usage_summary = compute_usage_summary(sheets)
    expiry_alerts = compute_expiry_alerts(sheets, days_ahead=expiry_days)
    stock_alerts  = compute_stock_alerts(sheets)


# ─── KPI Row ─────────────────────────────────────────────────────────────────
section("📊 ภาพรวม (KPI)")

k1, k2, k3, k4, k5 = st.columns(5)

total_items      = len(true_stock) if not true_stock.empty else 0
sap_items        = len(sheets.get("sap_stock", pd.DataFrame()))
no_sap_items     = len(sheets.get("no_sap", pd.DataFrame()))
chem_items       = len(sheets.get("chemical", pd.DataFrame()))
open_pr          = len(pr_po_adj[pr_po_adj["Adjusted_Outstanding"] > 0]) if not pr_po_adj.empty else 0
expiry_critical  = len(expiry_alerts[expiry_alerts["Days_Until_Expiry"] < 30]) if not expiry_alerts.empty else 0
stock_low        = len(stock_alerts) if not stock_alerts.empty else 0
total_usage      = usage_summary["Total_Usage"].sum() if not usage_summary.empty else 0

with k1:
    st.markdown(kpi("รายการวัสดุทั้งหมด", fmt_num(total_items),
                    f"SAP:{sap_items} / NoSAP:{no_sap_items} / Chem:{chem_items}"), unsafe_allow_html=True)
with k2:
    st.markdown(kpi("PR/PO ที่ยังค้างรับ", fmt_num(open_pr),
                    "รายการที่ Adjusted Outstanding > 0", "alert-teal"), unsafe_allow_html=True)
with k3:
    st.markdown(kpi("ใกล้หมดอายุ (<30 วัน)", fmt_num(expiry_critical),
                    f"ตรวจสอบทั้งหมด {expiry_days} วันข้างหน้า",
                    "alert-red" if expiry_critical > 0 else ""), unsafe_allow_html=True)
with k4:
    st.markdown(kpi("Stock ต่ำกว่า Safety", fmt_num(stock_low),
                    "รายการที่ต้องจัดซื้อ",
                    "alert-orange" if stock_low > 0 else ""), unsafe_allow_html=True)
with k5:
    st.markdown(kpi("Usage รวมทั้งหมด", fmt_num(total_usage),
                    "SAP_MB51 + Manual Summary"), unsafe_allow_html=True)

st.divider()


# ─── TABS ─────────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "📦 Stock รวม",
    "🧪 No SAP / Chemical",
    "📋 PR & PO",
    "📈 Usage",
    "⚠️ แจ้งเตือน",
    "🔍 ค้นหาวัสดุ",
    "📥 Export",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Stock รวม
# ══════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    section("📦 สรุปสต็อกวัสดุทั้งหมด (SAP + No SAP + Chemical)")

    if true_stock.empty:
        st.warning("ไม่พบข้อมูล Stock")
    else:
        col_f1, col_f2 = st.columns([3, 1])
        with col_f1:
            search_stock = st.text_input("🔍 ค้นหา Material ID / ชื่อ", key="search_stock")
        with col_f2:
            source_filter = st.selectbox("แหล่งข้อมูล", ["ทั้งหมด", "SAP", "No_SAP", "Chemical_NoSAP"], key="src_filter")

        df_show = true_stock.copy()
        if search_stock:
            mask = (
                df_show["Material_ID"].str.contains(search_stock, case=False, na=False) |
                df_show["Material_Name"].str.contains(search_stock, case=False, na=False)
            )
            df_show = df_show[mask]
        if source_filter != "ทั้งหมด":
            df_show = df_show[df_show["Sources"].str.contains(source_filter, na=False)]

        st.caption(f"แสดง {len(df_show):,} รายการ จากทั้งหมด {len(true_stock):,} รายการ")
        st.dataframe(
            df_show[["Material_ID", "Material_Name", "Unit", "Sources",
                     "SAP_Unrestricted", "SAP_Inspection", "SAP_InTransit",
                     "NoSAP_Remaining", "Total_Stock"]].reset_index(drop=True),
            use_container_width=True, height=450,
            column_config={
                "Material_ID":       st.column_config.TextColumn("Material ID", width=130),
                "Material_Name":     st.column_config.TextColumn("ชื่อวัสดุ", width=280),
                "Unit":              st.column_config.TextColumn("หน่วย", width=70),
                "Sources":           st.column_config.TextColumn("แหล่งข้อมูล", width=140),
                "SAP_Unrestricted":  st.column_config.NumberColumn("SAP Unrestricted", format="%,.2f"),
                "SAP_Inspection":    st.column_config.NumberColumn("SAP Inspection", format="%,.2f"),
                "SAP_InTransit":     st.column_config.NumberColumn("In Transit", format="%,.2f"),
                "NoSAP_Remaining":   st.column_config.NumberColumn("NoSAP คงเหลือ", format="%,.2f"),
                "Total_Stock":       st.column_config.NumberColumn("Total Stock", format="%,.2f"),
            },
        )

        # Chart — Top 20 by Total Stock
        st.markdown("---")
        section("📊 Top 20 วัสดุที่มี Stock มากที่สุด")
        top20 = df_show.nlargest(20, "Total_Stock").copy()
        if not top20.empty:
            fig = px.bar(
                top20, x="Total_Stock", y="Material_ID",
                orientation="h", color="Sources",
                color_discrete_map={"SAP": "#1A5276", "No_SAP": "#1A7A6E", "Chemical_NoSAP": "#935116"},
                hover_data=["Material_Name", "Unit"],
                labels={"Total_Stock": "จำนวนคงเหลือ", "Material_ID": ""},
                height=500,
            )
            fig.update_layout(
                plot_bgcolor="white", paper_bgcolor="white",
                yaxis={"categoryorder": "total ascending"},
                legend_title="แหล่งข้อมูล",
            )
            st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — No SAP / Chemical
# ══════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    section("🧪 วัสดุและสารเคมีนอกระบบ SAP")

    sub1, sub2 = st.tabs(["📦 No SAP", "⚗️ Chemical"])

    with sub1:
        no_sap_df = sheets.get("no_sap", pd.DataFrame())
        if no_sap_df.empty:
            st.info("ไม่พบข้อมูล No_SAP")
        else:
            search_ns = st.text_input("🔍 ค้นหา", key="ns_search")
            df_ns = no_sap_df.copy()
            if search_ns:
                df_ns = df_ns[
                    df_ns["Material_ID"].astype(str).str.contains(search_ns, case=False, na=False) |
                    df_ns["Material_Name"].astype(str).str.contains(search_ns, case=False, na=False)
                ]
            st.caption(f"{len(df_ns):,} รายการ")
            show_cols = ["NO", "Material_ID", "Material_Name", "Unit", "Dept",
                         "Quantity", "Total_Withdrawn", "Remaining",
                         "Receive_Date", "Expiry_Date", "Remark"]
            show_cols = [c for c in show_cols if c in df_ns.columns]
            st.dataframe(df_ns[show_cols].reset_index(drop=True), use_container_width=True, height=430)

    with sub2:
        chem_df = sheets.get("chemical", pd.DataFrame())
        if chem_df.empty:
            st.info("ไม่พบข้อมูล Chemical")
        else:
            search_ch = st.text_input("🔍 ค้นหา", key="ch_search")
            df_ch = chem_df.copy()
            if search_ch:
                df_ch = df_ch[
                    df_ch["Material_ID"].astype(str).str.contains(search_ch, case=False, na=False) |
                    df_ch["Material_Name"].astype(str).str.contains(search_ch, case=False, na=False)
                ]
            st.caption(f"{len(df_ch):,} รายการ")
            show_cols = ["NO", "Material_ID", "Material_Name", "Unit", "Dept",
                         "Quantity", "Total_Withdrawn", "Remaining",
                         "Receive_Date", "Mfg_Date", "Expiry_Date", "Remark"]
            show_cols = [c for c in show_cols if c in df_ch.columns]
            st.dataframe(df_ch[show_cols].reset_index(drop=True), use_container_width=True, height=430)

            # Expiry chart
            if "Expiry_Date" in df_ch.columns:
                today_ts = pd.Timestamp(date.today())
                df_exp = df_ch[df_ch["Expiry_Date"].notna()].copy()
                df_exp["Days_Left"] = (df_exp["Expiry_Date"] - today_ts).dt.days
                df_exp["Status"] = df_exp["Days_Left"].apply(
                    lambda d: "หมดอายุแล้ว" if d < 0
                    else ("< 30 วัน" if d < 30 else ("< 90 วัน" if d < 90 else "> 90 วัน"))
                )
                status_count = df_exp["Status"].value_counts().reset_index()
                status_count.columns = ["Status", "Count"]
                fig2 = px.pie(status_count, names="Status", values="Count",
                              title="สถานะอายุสารเคมีนอกระบบ",
                              color="Status",
                              color_discrete_map={
                                  "หมดอายุแล้ว": "#C0392B", "< 30 วัน": "#E67E22",
                                  "< 90 วัน": "#F1C40F", "> 90 วัน": "#1E8449",
                              })
                fig2.update_layout(paper_bgcolor="white")
                st.plotly_chart(fig2, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — PR & PO
# ══════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    section("📋 PR & PO — ข้อมูลคำสั่งซื้อ (Adjusted)")

    pr_sub1, pr_sub2 = st.tabs(["📊 SAP PR/PO (Adjusted)", "🔄 PR & PO Update"])

    with pr_sub1:
        if pr_po_adj.empty:
            st.info("ไม่พบข้อมูล SAP_ZRMM0004")
        else:
            col_a, col_b = st.columns([3, 1])
            with col_a:
                search_pr = st.text_input("🔍 ค้นหา PR/Material", key="pr_search")
            with col_b:
                show_open_only = st.checkbox("แสดงเฉพาะ Outstanding > 0", value=True, key="open_pr")

            df_pr = pr_po_adj.copy()
            if show_open_only:
                df_pr = df_pr[df_pr["Adjusted_Outstanding"] > 0]
            if search_pr:
                df_pr = df_pr[
                    df_pr["Material_ID"].astype(str).str.contains(search_pr, case=False, na=False) |
                    df_pr["PR_Number"].astype(str).str.contains(search_pr, case=False, na=False) |
                    df_pr["Product_Spec"].astype(str).str.contains(search_pr, case=False, na=False)
                ]

            st.caption(f"{len(df_pr):,} รายการ (Outstanding_QTY เดิม → Adjusted หลักหักรับสินค้าที่รอเข้าระบบ)")

            show_cols = ["PR_Number", "PO_Number", "Material_ID", "Product_Spec",
                         "Unit", "PR_Date", "Require_Date",
                         "Required_QTY", "Outstanding_QTY", "Adjusted_Outstanding"]
            show_cols = [c for c in show_cols if c in df_pr.columns]
            st.dataframe(
                df_pr[show_cols].reset_index(drop=True),
                use_container_width=True, height=430,
                column_config={
                    "Outstanding_QTY":   st.column_config.NumberColumn("Outstanding (SAP)", format="%,.2f"),
                    "Adjusted_Outstanding": st.column_config.NumberColumn("Adjusted Outstanding", format="%,.2f"),
                    "Required_QTY":      st.column_config.NumberColumn("Required QTY", format="%,.2f"),
                },
            )

    with pr_sub2:
        update_df = sheets.get("pr_po_update", pd.DataFrame())
        if update_df.empty:
            st.info("ไม่พบข้อมูล PR & PO Update")
        else:
            st.caption(f"{len(update_df):,} รายการ — วัสดุที่รับสินค้าแล้วแต่ยังรับเข้าระบบไม่ได้")
            st.dataframe(update_df.reset_index(drop=True), use_container_width=True, height=430)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — Usage
# ══════════════════════════════════════════════════════════════════════════════
with tabs[3]:
    section("📈 ประวัติการเบิกจ่าย (SAP MB51 + Manual)")

    use_sub1, use_sub2, use_sub3 = st.tabs(["📊 Usage Summary", "🗂️ SAP MB51 Raw", "📝 Manual Summary"])

    with use_sub1:
        if usage_summary.empty:
            st.info("ไม่พบข้อมูล Usage")
        else:
            search_use = st.text_input("🔍 ค้นหาวัสดุ", key="use_search")
            df_use = usage_summary.copy()
            if search_use:
                df_use = df_use[
                    df_use["Material_ID"].str.contains(search_use, case=False, na=False) |
                    df_use["Material_Name"].str.contains(search_use, case=False, na=False)
                ]
            st.caption(f"{len(df_use):,} รายการ")
            st.dataframe(
                df_use.reset_index(drop=True),
                use_container_width=True, height=400,
                column_config={
                    "SAP_Usage":    st.column_config.NumberColumn("SAP Usage", format="%,.2f"),
                    "Manual_Usage": st.column_config.NumberColumn("Manual Usage", format="%,.2f"),
                    "Total_Usage":  st.column_config.NumberColumn("Total Usage", format="%,.2f"),
                },
            )

            # Stacked bar chart — Top 20
            top20u = df_use.nlargest(20, "Total_Usage").copy()
            if not top20u.empty:
                fig3 = go.Figure()
                fig3.add_trace(go.Bar(
                    name="SAP Usage", x=top20u["Material_ID"], y=top20u["SAP_Usage"],
                    marker_color="#1A5276",
                ))
                fig3.add_trace(go.Bar(
                    name="Manual Usage", x=top20u["Material_ID"], y=top20u["Manual_Usage"],
                    marker_color="#1A7A6E",
                ))
                fig3.update_layout(
                    barmode="stack", title="Top 20 วัสดุที่ใช้มากที่สุด",
                    xaxis_tickangle=-45, plot_bgcolor="white",
                    paper_bgcolor="white", height=420,
                )
                st.plotly_chart(fig3, use_container_width=True)

    with use_sub2:
        mb51_df = sheets.get("mb51", pd.DataFrame())
        if mb51_df.empty:
            st.info("ไม่พบข้อมูล SAP_MB51")
        else:
            search_mb = st.text_input("🔍 ค้นหา", key="mb_search")
            df_mb = mb51_df.copy()
            if search_mb:
                df_mb = df_mb[
                    df_mb["Material_ID"].astype(str).str.contains(search_mb, case=False, na=False) |
                    df_mb["Material_Desc"].astype(str).str.contains(search_mb, case=False, na=False)
                ]
            st.caption(f"{len(df_mb):,} รายการ")
            st.dataframe(df_mb.reset_index(drop=True), use_container_width=True, height=430)

    with use_sub3:
        man_df = sheets.get("manual", pd.DataFrame())
        if man_df.empty:
            st.info("ไม่พบข้อมูล Manual_Summary")
        else:
            st.caption(f"{len(man_df):,} รายการ")
            st.dataframe(man_df.reset_index(drop=True), use_container_width=True, height=430)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — แจ้งเตือน
# ══════════════════════════════════════════════════════════════════════════════
with tabs[4]:
    section("⚠️ แจ้งเตือน — Expiry & Stock Alert")

    alert_sub1, alert_sub2 = st.tabs(["⏰ แจ้งเตือนหมดอายุ", "📉 Stock ต่ำกว่า Safety"])

    with alert_sub1:
        if expiry_alerts.empty:
            st.success("✅ ไม่พบวัสดุที่ใกล้หมดอายุในช่วง {expiry_days} วัน")
        else:
            # Summary metrics
            c1, c2, c3 = st.columns(3)
            expired = expiry_alerts[expiry_alerts["Days_Until_Expiry"] < 0]
            critical = expiry_alerts[(expiry_alerts["Days_Until_Expiry"] >= 0) & (expiry_alerts["Days_Until_Expiry"] < 30)]
            warning  = expiry_alerts[(expiry_alerts["Days_Until_Expiry"] >= 30) & (expiry_alerts["Days_Until_Expiry"] < 90)]
            with c1:
                st.markdown(kpi("หมดอายุแล้ว", len(expired), "", "alert-red"), unsafe_allow_html=True)
            with c2:
                st.markdown(kpi("ด่วน < 30 วัน", len(critical), "", "alert-orange"), unsafe_allow_html=True)
            with c3:
                st.markdown(kpi("ใกล้หมด < 90 วัน", len(warning)), unsafe_allow_html=True)

            st.dataframe(
                expiry_alerts.reset_index(drop=True),
                use_container_width=True, height=430,
                column_config={
                    "Days_Until_Expiry": st.column_config.NumberColumn("วันที่เหลือ", format="%d"),
                    "Remaining": st.column_config.NumberColumn("คงเหลือ", format="%,.2f"),
                    "Expiry_Date": st.column_config.DateColumn("วันหมดอายุ", format="DD/MM/YYYY"),
                },
            )

    with alert_sub2:
        if stock_alerts.empty:
            st.success("✅ Stock ทุกรายการอยู่ในระดับปกติ หรือยังไม่มีข้อมูล Safety Stock")
        else:
            st.caption(f"⚠️ พบ {len(stock_alerts):,} รายการที่ Stock ต่ำกว่า Safety Stock")
            st.dataframe(
                stock_alerts.reset_index(drop=True),
                use_container_width=True, height=430,
                column_config={
                    "Safety_Stock":  st.column_config.NumberColumn("Safety Stock", format="%,.2f"),
                    "SAP_Available": st.column_config.NumberColumn("SAP Available", format="%,.2f"),
                    "Stock_Gap":     st.column_config.NumberColumn("Gap", format="%,.2f"),
                },
            )

            # Chart
            chart_df = stock_alerts.copy().head(20)
            if not chart_df.empty:
                fig4 = go.Figure()
                fig4.add_trace(go.Bar(name="SAP Available", x=chart_df["Material_ID"],
                                      y=chart_df["SAP_Available"], marker_color="#1A7A6E"))
                fig4.add_trace(go.Bar(name="Safety Stock", x=chart_df["Material_ID"],
                                      y=chart_df["Safety_Stock"], marker_color="#C0392B",
                                      opacity=0.7))
                fig4.update_layout(
                    barmode="group", title="Stock vs Safety Stock (Top 20 ที่ต่ำสุด)",
                    xaxis_tickangle=-45, plot_bgcolor="white",
                    paper_bgcolor="white", height=420,
                )
                st.plotly_chart(fig4, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — ค้นหาวัสดุ (All-in-one)
# ══════════════════════════════════════════════════════════════════════════════
with tabs[5]:
    section("🔍 ค้นหาวัสดุแบบครอบคลุม (Cross-sheet)")

    q = st.text_input("พิมพ์ Material ID หรือชื่อวัสดุ", placeholder="เช่น T11-1001 หรือ Hydrochloric")

    if q:
        results = {}

        for key, label, id_col, name_col in [
            ("sap_stock",    "SAP_ZRMM0001 (Stock)",  "Material_ID",  "Material_Desc"),
            ("sap_pr_po",    "SAP_ZRMM0004 (PR/PO)",  "Material_ID",  "Product_Spec"),
            ("no_sap",       "No_SAP",                "Material_ID",  "Material_Name"),
            ("chemical",     "Chemical NoSAP",        "Material_ID",  "Material_Name"),
            ("mb51",         "SAP_MB51 (Usage)",      "Material_ID",  "Material_Desc"),
            ("manual",       "Manual Summary",        "Material_ID",  "Material_Name"),
            ("pr_po_update", "PR & PO Update",        "Material_ID",  "Material_Name"),
        ]:
            df = sheets.get(key, pd.DataFrame())
            if df.empty:
                continue
            id_c  = id_col  if id_col  in df.columns else None
            nm_c  = name_col if name_col in df.columns else None
            masks = []
            if id_c:
                masks.append(df[id_c].astype(str).str.contains(q, case=False, na=False))
            if nm_c:
                masks.append(df[nm_c].astype(str).str.contains(q, case=False, na=False))
            if masks:
                combined = masks[0]
                for m in masks[1:]:
                    combined = combined | m
                found = df[combined]
                if len(found) > 0:
                    results[label] = found

        if results:
            st.success(f"พบข้อมูลใน {len(results)} แหล่ง")
            for label, df_r in results.items():
                with st.expander(f"📂 {label} — {len(df_r)} รายการ", expanded=True):
                    st.dataframe(df_r.reset_index(drop=True), use_container_width=True)
        else:
            st.warning(f"ไม่พบข้อมูลสำหรับ: **{q}**")
    else:
        st.info("พิมพ์ Material ID หรือชื่อวัสดุเพื่อค้นหาข้ามทุก Sheet")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — Export
# ══════════════════════════════════════════════════════════════════════════════
with tabs[6]:
    section("📥 Export รายงาน Excel")

    st.markdown("""
    รายงานที่ Export จะประกอบด้วย:
    - **Stock Summary** — สต็อกรวมจากทุกแหล่ง
    - **PR-PO Adjusted** — คำสั่งซื้อที่ปรับ Outstanding แล้ว
    - **Usage Summary** — สรุปการเบิกจ่ายรวม SAP + Manual
    - **Expiry Alerts** — วัสดุที่ใกล้หมดอายุ
    - **Stock Alerts** — วัสดุที่ Stock ต่ำกว่า Safety
    """)

    col_ex1, col_ex2 = st.columns([2, 1])
    with col_ex1:
        export_name = st.text_input(
            "ชื่อไฟล์",
            value=f"TH_Plant_Material_Report_{date.today().strftime('%Y%m%d')}.xlsx",
        )
    with col_ex2:
        st.write("")
        st.write("")
        do_export = st.button("🖨️ สร้างรายงาน Excel", type="primary", use_container_width=True)

    if do_export:
        with st.spinner("กำลังสร้างรายงาน..."):
            try:
                excel_bytes = export_full_report(
                    true_stock, pr_po_adj, usage_summary,
                    expiry_alerts, stock_alerts,
                )
                st.success("✅ รายงานพร้อมดาวน์โหลด")
                st.download_button(
                    label="⬇️ ดาวน์โหลด Excel",
                    data=excel_bytes,
                    file_name=export_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"❌ Export ล้มเหลว: {e}")
