import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------
# 1. إعدادات التنسيق والألوان
# ------------------------------------------------------------
STANDARD_SLOTS = [
    {"label": "8:30-9:45",   "start_min": 8*60+30, "end_min": 9*60+45},
    {"label": "10:00-11:15", "start_min": 10*60,   "end_min": 11*60+15},
    {"label": "11:30-12:45", "start_min": 11*60+30, "end_min": 12*60+45},
    {"label": "13:00-14:15", "start_min": 13*60,   "end_min": 14*60+15},
]

# الألوان الرسمية
HEADER_FILL = PatternFill(start_color="366092", end_color="366092", fill_type="solid") # أزرق للهيدر
EMPTY_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # أخضر للفارغ
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
GRAY_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")   # رمادي فاتح للتبادل
WHITE_FONT = Font(color="FFFFFF", bold=True)
BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# ------------------------------------------------------------
# 2. وظائف المعالجة والكشف
# ------------------------------------------------------------
def split_course_code(code):
    match = re.match(r"([a-zA-Z]+)([0-9]+)", str(code))
    if match: return match.group(1), match.group(2)
    return str(code), ""

def parse_time_range(t_str):
    try:
        parts = str(t_str).replace('–', '-').split('-')
        if len(parts) != 2: return 0, 0
        s_dt = datetime.strptime(parts[0].strip(), "%H:%M")
        e_dt = datetime.strptime(parts[1].strip(), "%H:%M")
        return s_dt.hour * 60 + s_dt.minute, e_dt.hour * 60 + e_dt.minute
    except: return 0, 0

def parse_excel_files(files):
    all_data = []
    days_map = {'Sunday': 'الاحد', 'Monday': 'الاثنين', 'Tuesday': 'الثلاثاء', 'Wednesday': 'الأربعاء', 'Thursday': 'الخميس'}
    
    for f in files:
        xl = pd.ExcelFile(io.BytesIO(f.read()))
        for sheet in xl.sheet_names:
            df = xl.parse(sheet, header=None).reset_index(drop=True)
            section = "Unknown"
            for r in range(len(df)):
                val = str(df.iloc[r, 0]).strip()
                if val in days_map:
                    day_en = val
                    for c in range(1, df.shape[1]):
                        code_full = str(df.iloc[r, c]).strip()
                        if not code_full or code_full.lower() in ["nan", "1", "2", "3", "4", "5"]: continue
                        
                        prefix, num = split_course_code(code_full)
                        time_str = str(df.iloc[r+2, c])
                        start_m, end_m = parse_time_range(time_str)
                        
                        all_data.append({
                            "التخصص": sheet,
                            "المقرر": str(df.iloc[r+1, c]).strip(),
                            "الرمز": prefix, "الرقم": num, "الشعبة": section,
                            "Day": day_en, "Time": time_str,
                            "القاعة": str(df.iloc[r+3, c]).strip(),
                            "StartMin": start_m, "EndMin": end_m
                        })
                elif val and val != "nan" and not val.isdigit():
                    section = val
    return pd.DataFrame(all_data)

def detect_conflicts(df):
    conflicts = []
    valid_df = df[~df['القاعة'].isin(["?", "nan", "", "N/A"])].copy()
    valid_df = valid_df.sort_values(['القاعة', 'Day', 'StartMin'])

    for (room, day), group in valid_df.groupby(['القاعة', 'Day']):
        rows = group.to_dict('records')
        for i in range(len(rows)-1):
            curr, nxt = rows[i], rows[i+1]
            if nxt['StartMin'] < curr['EndMin']:
                if curr['الرمز'] != nxt['الرمز'] or curr['الشعبة'] != nxt['الشعبة']:
                    conflicts.append({
                        "القاعة": room, "اليوم": day,
                        "التفاصيل": f"تداخل بين [{curr['الرمز']}{curr['الرقم']}] و [{nxt['الرمز']}{nxt['الرقم']}]",
                        "الوقت": f"{curr['Time']} مع {nxt['Time']}"
                    })
    return conflicts

# ------------------------------------------------------------
# 3. وظائف التصدير المنسقة
# ------------------------------------------------------------
def export_linear_format(df):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Linear Schedule"
    
    # إضافة التخصص للترتيب
    headers = ["التخصص", "المقرر", "الرمز", "الرقم", "الشعبة", "الاحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "القاعة"]
    ws.append(headers)
    
    for cell in ws[1]:
        cell.fill = HEADER_FILL; cell.font = WHITE_FONT; cell.alignment = Alignment(horizontal='center')

    # الترتيب حسب التخصص ثم المقرر
    df_sorted = df.sort_values(['التخصص', 'المقرر', 'الشعبة'])
    grouped = df_sorted.groupby(['التخصص', 'المقرر', 'الرمز', 'الرقم', 'الشعبة', 'القاعة'], sort=False)
    days_cols = {'Sunday': 6, 'Monday': 7, 'Tuesday': 8, 'Wednesday': 9, 'Thursday': 10}
    
    last_item_id = None
    use_gray = False

    for (spec, name, pref, num, sec, room), group in grouped:
        # تبديل اللون عند تغير المادة (المقرر + الشعبة)
        current_item_id = (spec, name, sec)
        if current_item_id != last_item_id:
            use_gray = not use_gray
            last_item_id = current_item_id
        
        current_fill = GRAY_FILL if use_gray else WHITE_FILL
        row_data = [spec, name, pref, num, sec, "", "", "", "", "", room]
        for _, item in group.iterrows():
            col_idx = days_cols.get(item['Day'])
            if col_idx: row_data[col_idx-1] = item['Time']
        
        ws.append(row_data)
        for cell in ws[ws.max_row]:
            cell.fill = current_fill; cell.border = BORDER; cell.alignment = Alignment(horizontal='center', vertical='center')

    # ضبط أحجام الأعمدة
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 25
    for col in range(3, 12): ws.column_dimensions[get_column_letter(col)].width = 15
    
    wb.save(output)
    return output.getvalue()

def export_rooms_styled(df, mode="occupied"):
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    days_en = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday']
    slots = [s["label"] for s in STANDARD_SLOTS]

    if mode == "occupied":
        for room in sorted(df['القاعة'].unique()):
            if room in ["nan", "?", ""]: continue
            ws = wb.create_sheet(title=str(room)[:31])
            ws.append(["اليوم / الحصة"] + slots)
            for cell in ws[1]: cell.fill = HEADER_FILL; cell.font = WHITE_FONT; cell.alignment = Alignment(horizontal='center')

            for d_en in days_en:
                row_vals = [d_en]
                for s in slots:
                    m = df[(df['القاعة'] == room) & (df['Day'] == d_en)]
                    cell_text = ""
                    for _, r in m.iterrows():
                        for slot_cfg in STANDARD_SLOTS:
                            if slot_cfg["label"] == s and abs(r['StartMin'] - slot_cfg["start_min"]) <= 30:
                                cell_text += f"{r['الرمز']}{r['الرقم']}\n{r['المقرر']}\n({r['الشعبة']})\n"
                    row_vals.append(cell_text.strip())
                ws.append(row_vals)

            for row in ws.iter_rows(min_row=2):
                ws.row_dimensions[row[0].row].height = 65
                for cell in row:
                    cell.border = BORDER; cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
                    if cell.column > 1 and not cell.value: cell.fill = EMPTY_FILL
            
            ws.column_dimensions['A'].width = 15
            for col in range(2, 7): ws.column_dimensions[get_column_letter(col)].width = 25
    else:
        # ملف القاعات الفارغة
        ws = wb.create_sheet(title="Available Rooms")
        ws.append(["Day / Slot"] + slots)
        for cell in ws[1]: cell.fill = HEADER_FILL; cell.font = WHITE_FONT; cell.alignment = Alignment(horizontal='center')
        
        all_rooms = set(df['القاعة'].unique()) - {"", "nan", "?"}
        for d_en in days_en:
            row = [days_en.index(d_en)] # placeholder
            row[0] = d_en
            for s in slots:
                occ = set(df[(df['Day'] == d_en) & (df['StartMin'].apply(lambda x: any(abs(x - sc['start_min']) <= 30 for sc in STANDARD_SLOTS if sc['label'] == s)))]['القاعة'].unique())
                row.append(", ".join(sorted(list(all_rooms - occ))))
            ws.append(row)
            
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = BORDER; cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

    wb.save(output)
    return output.getvalue()

# ------------------------------------------------------------
# 4. واجهة Streamlit
# ------------------------------------------------------------
st.set_page_config(page_title="Scheduler Pro V6", layout="wide")
st.title("📅 نظام الجداول (الكلية التطبيقية-جامعة طيبة)")

files = st.file_uploader("ارفع ملفات Excel", type="xlsx", accept_multiple_files=True)

if files:
    data = parse_excel_files(files)
    conflicts = detect_conflicts(data)
    
    tab1, tab2, tab3 = st.tabs(["📋 التنسيق الخطي (Linear)", "🏢 القاعات والشواغر", "⚠️ التعارضات"])

    with tab1:
        st.subheader("المعاينة الخطية (مرتبة حسب التخصص)")
        st.download_button("📥 تحميل التنسيق الخطي ", export_linear_format(data), "Linear_Clean_Schedule.xlsx")
        st.dataframe(data.sort_values(["التخصص", "المقرر"])[["التخصص", "المقرر", "الرمز", "الرقم", "الشعبة", "Day", "Time", "القاعة"]], use_container_width=True)

    with tab2:
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("🏢 تحميل جدول القاعات (ملون)", export_rooms_styled(data, "occupied"), "Rooms_Schedule.xlsx")
        with col2:
            st.download_button("✅ تحميل القاعات الفارغة", export_rooms_styled(data, "empty"), "Available_Rooms.xlsx")

    with tab3:
        if conflicts:
            st.error(f"يوجد {len(conflicts)} تداخل في القاعات!")
            st.table(pd.DataFrame(conflicts))
        else:
            st.success("✅ لا توجد تداخلات زمنية.")
