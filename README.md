# 🏭 TH Plant — Material & Chemical Management System

ระบบคำนวณและติดตามวัสดุ + สารเคมี สำหรับโรงงาน TH Plant  
สร้างด้วย Python + Streamlit · ใช้งานผ่าน GitHub + Streamlit Community Cloud

---

## 📋 Sheet ที่รองรับในไฟล์ Excel หลัก

| Sheet | คำอธิบาย |
|---|---|
| `No_SAP` | วัสดุที่ยังไม่อยู่ในระบบ SAP |
| `Chemical` | สารเคมีที่ยังไม่อยู่ในระบบ SAP |
| `SAP_ZRMM0001` | ข้อมูล Stock ในระบบ SAP |
| `SAP_ZRMM0004` | ข้อมูล PR & PO ในระบบ SAP |
| `PR & PO Update` | วัสดุที่รับสินค้าแล้วแต่รับเข้าระบบไม่ได้ |
| `SAP_MB51` | ประวัติการเบิกจ่ายจาก SAP |
| `Manual_Summary` | การเบิกจ่ายที่ไม่อยู่ในระบบ SAP |
| `Stock Material` | ข้อมูลหลัก: Safety Stock, Daily Consumption |

---

## 🚀 วิธีติดตั้งและรันบน Local

```bash
# 1. Clone repo
git clone https://github.com/<YOUR_USERNAME>/<YOUR_REPO>.git
cd <YOUR_REPO>

# 2. ติดตั้ง dependencies
pip install -r requirements.txt

# 3. รันแอป
streamlit run app.py
```

เปิดเบราว์เซอร์ที่ `http://localhost:8501`

---

## ☁️ Deploy บน Streamlit Community Cloud (ฟรี)

### ขั้นตอน:

1. **Push โค้ดขึ้น GitHub**
   ```bash
   git init
   git add .
   git commit -m "Initial commit: TH Plant Material System"
   git branch -M main
   git remote add origin https://github.com/<YOUR_USERNAME>/<YOUR_REPO>.git
   git push -u origin main
   ```

2. **เปิด Streamlit Cloud**
   - ไปที่ [share.streamlit.io](https://share.streamlit.io)
   - Login ด้วย GitHub account

3. **Deploy**
   - กด **"New app"**
   - เลือก Repository / Branch `main` / Main file `app.py`
   - กด **"Deploy!"**

4. **รอประมาณ 1–2 นาที** → ได้ URL สาธารณะ เช่น  
   `https://<app-name>.streamlit.app`

---

## 🗂️ โครงสร้างไฟล์

```
material_system/
├── app.py              # หน้าหลัก Streamlit
├── data_loader.py      # โหลดและ normalize ทุก sheet
├── calculator.py       # คำนวณ Stock, PR/PO, Usage, Alerts
├── export_utils.py     # สร้างรายงาน Excel
├── requirements.txt    # Python packages
├── .streamlit/
│   └── config.toml     # Theme & settings
└── .github/
    └── workflows/
        └── ci.yml      # GitHub Actions CI
```

---

## ⚙️ ฟีเจอร์หลัก

| ฟีเจอร์ | คำอธิบาย |
|---|---|
| 📦 Stock รวม | รวม SAP + No_SAP + Chemical → Total Stock ต่อ Material ID |
| 📋 PR/PO Adjusted | ปรับ Outstanding โดยหักวัสดุที่รับแล้วแต่ยังไม่เข้าระบบ |
| 📈 Usage Summary | รวม SAP_MB51 + Manual_Summary แสดง Total Usage |
| ⚠️ Expiry Alert | แจ้งเตือนวัสดุใกล้หมดอายุ (ปรับช่วงวันได้) |
| 📉 Stock Alert | แจ้งเตือน Stock ต่ำกว่า Safety Stock |
| 🔍 Cross-search | ค้นหาวัสดุข้ามทุก Sheet พร้อมกัน |
| 📥 Export | ดาวน์โหลดรายงาน Excel สรุปครบทุก Sheet |

---

## 📦 Dependencies

```
streamlit>=1.32.0
pandas>=2.0.0
openpyxl>=3.1.0
plotly>=5.18.0
xlsxwriter>=3.1.0
numpy>=1.26.0
```

---

## 🔒 ข้อมูลสำคัญ

- ไม่มีการเก็บข้อมูลบนเซิร์ฟเวอร์ ไฟล์จะถูกประมวลผลใน session เท่านั้น
- ทุกครั้งที่อัปโหลดไฟล์ใหม่ระบบจะคำนวณใหม่ทั้งหมด
- ใช้ `@st.cache_data` เพื่อลดเวลาโหลดซ้ำในระหว่าง session

---

*TH Plant Material System · Built with ❤️ using Streamlit*
