import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------
# Define standard time slots (label, start, end in minutes)
# ------------------------------------------------------------
STANDARD_SLOTS = [
    {"label": "8:30-9:45",   "start_min": 8*60+30, "end_min": 9*60+45},
    {"label": "10:00-11:15", "start_min": 10*60,   "end_min": 11*60+15},
    {"label": "11:30-12:45", "start_min": 11*60+30, "end_min": 12*60+45},
    {"label": "13:00-14:15", "start_min": 13*60,   "end_min": 14*60+15},
]

# Mapping from actual booking start minutes to standard slot label
def map_to_standard_slot(start_min):
    for slot in STANDARD_SLOTS:
        if abs(start_min - slot["start_min"]) <= 10:
            return slot["label"]
    return f"{start_min//60:02d}:{start_min%60:02d}"

# ------------------------------------------------------------
# Helper: parse a time string like "8:30 - 9:45"
# ------------------------------------------------------------
def parse_time_slot(time_str):
    parts = time_str.replace('–', '-').split('-')
    start = parts[0].strip()
    end = parts[1].strip()
    start_dt = datetime.strptime(start, "%H:%M")
    end_dt = datetime.strptime(end, "%H:%M")
    start_min = start_dt.hour * 60 + start_dt.minute
    end_min = end_dt.hour * 60 + end_dt.minute
    return start_min, end_min, start, end

# ------------------------------------------------------------
# Parse a single Excel file (in-memory bytes)
# Supports both 3-row and 4-row formats.
# ------------------------------------------------------------
def parse_file(file_bytes, filename):
    all_bookings = []
    excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
    for sheet_name in excel_file.sheet_names:
        df = excel_file.parse(sheet_name, header=None)
        num_cols = df.shape[1]
        # Find section rows (F1, F2, ...)
        section_rows = []
        for idx, row in df.iterrows():
            cell = str(row[0]) if pd.notna(row[0]) else ""
            if cell.startswith('F') and cell[1:].isdigit():
                section_rows.append((idx, cell))
        if not section_rows:
            continue
        # Process each section
        for sec_idx, (start_row, section_name) in enumerate(section_rows):
            end_row = section_rows[sec_idx+1][0] if sec_idx+1 < len(section_rows) else len(df)
            # Find day rows (Sunday..Thursday) in this section
            day_rows = []
            for r in range(start_row, end_row):
                first_cell = str(df.iloc[r, 0]) if pd.notna(df.iloc[r, 0]) else ""
                if first_cell in ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday']:
                    day_rows.append(r)
            # For each day, detect format (3-row or 4-row)
            for day_row in day_rows:
                if day_row + 3 >= len(df):
                    continue
                row2 = df.iloc[day_row + 2]  # third row after day
                is_3row = any(":" in str(cell) for cell in row2 if pd.notna(cell))
                if is_3row:
                    courses_row = df.iloc[day_row + 1]
                    times_row = df.iloc[day_row + 2]
                    rooms_row = df.iloc[day_row + 3]
                else:
                    courses_row = df.iloc[day_row + 1]
                    names_row = df.iloc[day_row + 2]
                    times_row = df.iloc[day_row + 3]
                    rooms_row = df.iloc[day_row + 4]
                day_name = str(df.iloc[day_row, 0])
                for col in range(1, num_cols):
                    course_code = courses_row[col] if pd.notna(courses_row[col]) else ""
                    if not course_code:
                        continue
                    # Get Arabic name if present
                    course_name = ""
                    if not is_3row:
                        course_name = names_row[col] if pd.notna(names_row[col]) else ""
                    time_str = times_row[col] if pd.notna(times_row[col]) else ""
                    room = rooms_row[col] if pd.notna(rooms_row[col]) else ""
                    if not time_str or not room:
                        continue
                    try:
                        start_min, end_min, start_str, end_str = parse_time_slot(time_str)
                    except:
                        continue
                    all_bookings.append({
                        "Specialty": sheet_name,
                        "Section": section_name,
                        "Day": day_name,
                        "Start": start_str,
                        "End": end_str,
                        "StartMin": start_min,
                        "EndMin": end_min,
                        "Room": room,
                        "CourseCode": course_code,
                        "CourseName": course_name,
                        "File": filename,
                        "OriginalTime": time_str
                    })
    return all_bookings

# ------------------------------------------------------------
# Conflict detection
# ------------------------------------------------------------
def find_conflicts(df):
    conflicts = []
    for (room, day), group in df.groupby(['Room', 'Day']):
        group = group.sort_values('StartMin')
        for i in range(len(group)-1):
            current = group.iloc[i]
            next_ = group.iloc[i+1]
            if next_['StartMin'] < current['EndMin']:
                conflicts.append({
                    "Room": room,
                    "Day": day,
                    "Conflict": f"{current['CourseCode']} {current['CourseName']} ({current['Section']}) {current['Start']}-{current['End']} overlaps with {next_['CourseCode']} {next_['CourseName']} ({next_['Section']}) {next_['Start']}-{next_['End']}"
                })
    return pd.DataFrame(conflicts)

# ------------------------------------------------------------
# Weekly grid for a room using standard slots.
# Each cell contains:
#   line 1: original time
#   line 2: course code
#   line 3: course name (if present)
#   line 4: section
# ------------------------------------------------------------
def weekly_grid_for_room(df, room_name):
    room_df = df[df['Room'] == room_name]
    room_df['Slot'] = room_df['StartMin'].apply(map_to_standard_slot)
    
    days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday']
    slot_labels = [slot["label"] for slot in STANDARD_SLOTS]
    grid = {day: {slot: [] for slot in slot_labels} for day in days}
    for _, row in room_df.iterrows():
        day = row['Day']
        slot = row['Slot']
        if slot in grid[day]:
            # Build the cell content as lines
            lines = [row['OriginalTime']]
            lines.append(row['CourseCode'])
            if row['CourseName'] and not pd.isna(row['CourseName']):
                lines.append(row['CourseName'])
            lines.append(row['Section'])
            grid[day][slot].append("\n".join(lines))
    # Convert to DataFrame
    data = []
    for day in days:
        row = []
        for slot in slot_labels:
            entries = grid[day][slot]
            cell_text = "\n\n".join(entries) if entries else ""
            row.append(cell_text)
        data.append(row)
    return pd.DataFrame(data, index=days, columns=slot_labels)

# ------------------------------------------------------------
# Export all rooms to a formatted Excel file
# ------------------------------------------------------------
def export_all_rooms_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for room in sorted(df['Room'].unique()):
            weekly = weekly_grid_for_room(df, room)
            if not weekly.empty:
                sheet_name = room[:31]
                weekly.to_excel(writer, sheet_name=sheet_name, index=True)
    wb = load_workbook(output)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    for sheet in wb.worksheets:
        # Set column widths based on longest line
        for col in sheet.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    lines = str(cell.value).split('\n')
                    for line in lines:
                        max_length = max(max_length, len(line))
            adjusted_width = min(max_length + 2, 50)
            sheet.column_dimensions[col_letter].width = adjusted_width
        # Style header row
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
        # Style data cells
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = thin_border
                if cell.value is None or cell.value == "":
                    cell.fill = green_fill
        sheet.freeze_panes = 'A2'
    output_final = io.BytesIO()
    wb.save(output_final)
    return output_final.getvalue()

# ------------------------------------------------------------
# Streamlit App
# ------------------------------------------------------------
st.set_page_config(page_title="Room Schedule Viewer", layout="wide")
st.title("📅 Room Schedule Viewer")

uploaded_files = st.file_uploader(
    "Upload Excel schedule files",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    all_bookings = []
    for file in uploaded_files:
        with st.spinner(f"Parsing {file.name}..."):
            bookings = parse_file(file.read(), file.name)
            all_bookings.extend(bookings)
    if not all_bookings:
        st.warning("No bookings found. Check file format.")
        st.stop()
    
    bookings_df = pd.DataFrame(all_bookings)
    st.success(f"Loaded {len(bookings_df)} bookings from {len(uploaded_files)} file(s).")
    
    # Conflicts
    conflicts_df = find_conflicts(bookings_df)
    if not conflicts_df.empty:
        st.error(f"⚠️ Found {len(conflicts_df)} conflicts!")
        with st.expander("Show conflicts"):
            st.dataframe(conflicts_df)
    else:
        st.success("✅ No conflicts detected.")
    
    # Room selection and schedule
    all_rooms = sorted(bookings_df['Room'].unique())
    selected_room = st.sidebar.selectbox("Select a room", all_rooms)
    st.subheader(f"Weekly Schedule for Room {selected_room}")
    weekly = weekly_grid_for_room(bookings_df, selected_room)
    st.dataframe(weekly, use_container_width=True)
    
    # Heatmap
    st.subheader("Room Usage Heatmap (by day and standard time slot)")
    bookings_df['Slot'] = bookings_df['StartMin'].apply(map_to_standard_slot)
    slot_labels = [slot["label"] for slot in STANDARD_SLOTS]
    heatmap = pd.crosstab(index=[bookings_df['Room'], bookings_df['Day']], columns=bookings_df['Slot'])
    heatmap = heatmap.reindex(columns=slot_labels, fill_value=0)
    st.dataframe(heatmap, use_container_width=True)
    
    # Course schedule
    st.subheader("📖 Course Schedule")
    # Create a friendly display name for the dropdown: "Code - Name (if any)"
    bookings_df['CourseDisplay'] = bookings_df.apply(
        lambda row: f"{row['CourseCode']} - {row['CourseName']}" if row['CourseName'] else row['CourseCode'],
        axis=1
    )
    unique_courses = sorted(bookings_df[['CourseCode', 'CourseName', 'CourseDisplay']].drop_duplicates().values.tolist())
    course_display_options = [f"{c[0]} - {c[1]}" if c[1] else c[0] for c in unique_courses]
    selected_course_display = st.selectbox("Select a course", course_display_options)
    # Find the corresponding code (first part)
    selected_code = selected_course_display.split(" - ")[0] if " - " in selected_course_display else selected_course_display
    course_schedule = bookings_df[bookings_df['CourseCode'] == selected_code][
        ['Section', 'Day', 'Start', 'End', 'Room', 'Specialty']
    ].reset_index(drop=True)
    if not course_schedule.empty:
        st.dataframe(course_schedule, use_container_width=True)
        csv_course = course_schedule.to_csv(index=False).encode('utf-8')
        st.download_button(
            label=f"Download schedule for {selected_course_display} (CSV)",
            data=csv_course,
            file_name=f"{selected_code}_schedule.csv",
            mime="text/csv"
        )
    else:
        st.write("No schedule found.")
    
    # Export all rooms as Excel
    st.sidebar.header("📤 Export Schedules")
    if len(all_rooms) > 0:
        excel_data = export_all_rooms_excel(bookings_df)
        st.sidebar.download_button(
            label="Download all rooms schedule (Excel)",
            data=excel_data,
            file_name="all_rooms_schedule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload Excel files to start.")
