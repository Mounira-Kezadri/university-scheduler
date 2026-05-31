# 📅 نظام تصميم وإدارة الجداول الدراسية الجامعية
### University Academic Schedule Design & Management System

**Developer:** Dr. Mounira Kezadri  
**Institution:** Applied College — Taibah University  
**Unit:** Computer Science & Information Programs  
**Version:** 1.0 (Initial Release)

---

## 📌 Overview

An interactive web application built with **Python + Streamlit** that automates the creation and analysis of academic schedules for university departments. The system reads structured Excel files, detects room conflicts, generates occupancy maps, and exports professionally formatted reports with full Arabic RTL support.

---

## ✨ Features

| Feature | Description |
|---|---|
| **Multi-file Import** | Upload and merge multiple Excel files simultaneously (one per specialization) |
| **Auto Pattern Detection** | Automatically detects 3-row or 4-row data formats |
| **Conflict Detection** | Instantly flags time overlaps in rooms with full details |
| **Room Schedule View** | Weekly grid per room with color-coded empty slots |
| **Occupancy Heatmap** | Statistical cross-table showing room usage by day and period |
| **Linear Schedule Export** | Sorted by specialization, with one row per section-day |
| **Available Rooms Export** | Shows free rooms per time slot across all days |
| **Full Arabic RTL Support** | Arabic course names and right-to-left interface |

---

## ☁️ Cloud Deployment (No Installation Required)

Deploy for free on [Streamlit Community Cloud](https://streamlit.io/cloud):

1. Push the project to a public GitHub repository
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo and select `V3Scheduler.py` as the main file
4. Click **Deploy** — your app will be live at a shareable URL

---

## 🖥️ Requirements

- Python 3.8 or higher
- See `requirements.txt` for all dependencies

---

## ⚙️ Installation

**1. Clone or download the project:**
```bash
git clone https://github.com/Mounira-Kezadri/university-scheduler.git
cd university-scheduler
```

**2. (Recommended) Create a virtual environment:**
```bash
python -m venv venv

# Windows:
venv\Scripts\activate

# macOS/Linux:
source venv/bin/activate
```

**3. Install dependencies:**
```bash
pip install -r requirements.txt
```

---

## 🚀 Running the Application

```bash
streamlit run Scheduler.py
```

The app will open automatically at: **http://localhost:8501**

If it doesn't open automatically, copy the URL and paste it into your browser.

**To stop the app:** Press `Ctrl + C` in the terminal.


---

## 📁 Input File Structure

Each Excel file should represent **one academic specialization** and follow this structure:

```
Sheet name  → Specialization code (e.g., CS, IS, IT)

4-row layout per course per day cell:
  Row +0 : Day name in English  (Sunday … Thursday)
  Row +1 : Course name in Arabic
  Row +2 : Course code           (e.g., CS101)
  Row +3 : Time                  (HH:MM-HH:MM  e.g., 08:00-09:30)
  Row +4 : Room                  (e.g., B201)

3-row layout (no Arabic name row):
  Row +0 : Day name
  Row +1 : Course code
  Row +2 : Time
  Row +3 : Room

Section headers appear as standalone cell values (e.g., F1, F2, F3)
```

> ⚠️ A sample input file (`sample_input.xlsx`) is included in the repository. Review it before preparing your own files.

**Supported time formats:**
- Standard hyphen: `08:00-09:30`
- Em-dash: `08:00–09:30`

**Standard time slots used by the system:**

| Slot | Start | End |
|---|---|---|
| 1 | 08:30 | 09:45 |
| 2 | 10:00 | 11:15 |
| 3 | 11:30 | 12:45 |
| 4 | 13:00 | 14:15 |

---

## ⚙️ Customising Time Slots

Create a `config.json` file in the same folder as the script:

```json
{
  "time_slots": [
    {"label": "8:00-9:30",   "start_min": 480, "end_min": 570},
    {"label": "9:45-11:15",  "start_min": 585, "end_min": 675},
    {"label": "11:30-13:00", "start_min": 690, "end_min": 780}
  ]
}
```

The app will load these on startup — no code changes required.

---

## 📤 Outputs

| File | Contents |
|---|---|
| `Linear_Clean_Schedule.xlsx` | All sections in a flat linear table, sorted by specialization |
| `Rooms_Schedule.xlsx` | One sheet per room — weekly grid with color-coded occupancy |
| `Available_Rooms.xlsx` | Available (free) rooms per day and time slot |

---

## 📂 Project Structure

```
university-scheduler/
│
├── Scheduler.py            # Main application
├── requirements.txt        # Python dependencies
├── config.json             # (optional) custom time slots
├── README.md               # This file
└── sample_input.xlsx       # Example input file
```

---

## 🔧 Troubleshooting

| Problem | Solution |
|---|---|
| No results after uploading | Verify the file is `.xlsx` and matches the required structure |
| Library install error | Ensure Python 3.8+ is installed; run terminal as Administrator (Windows) |
| Browser doesn't open | Manually navigate to `http://localhost:8501` |
| Unexpected conflicts | Check that all time values follow the `HH:MM-HH:MM` format consistently |
| Arabic text not displaying | Use a modern browser: Chrome, Edge, or Firefox |

---

## 📖 How to Cite

If you use this software in academic work, please cite it as:

```
Kezadri, M. (2025). University Academic Schedule Design & Management System
(Version 1.0). Applied College, Taibah University.
```

---

## 📜 License

This project is licensed under the **MIT License** — see the `LICENSE` file for details.

---

## 👩‍💻 Author

**Dr. Mounira Kezadri**  
Applied College, Taibah University  
Computer Science & Information Programs Unit
