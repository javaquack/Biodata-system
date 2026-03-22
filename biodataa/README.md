# Web-Based Academic Biodata Management System

A Flask web application for managing student academic records across 8 semesters. It features a data integration system that parses PDFs and Excel files to merge attendance and marks into a single student biodata profile using **Register Number** and **Subject Codes**.

## 🚀 Key Features

- **Integrated Data Parser**: Upload multiple PDFs/Excel files at once. The system automatically identifies student IDs and subject codes to update records.
- **Teacher Dashboard**:
  - Semester-wise filtering (S1 to S8).
  - Visual alerts for students with attendance below 75%.
  - Unified Biodata view with history of all 8 semesters.
  - Delete student records.
- **Student Module**:
  - View personal academic records across all semesters.
  - Self-service profile updates (phone, email, address, etc.).
- **Multi-Semester Architecture**: Supports data from S1 through S8.

## 🛠 Prerequisites

- Python 3.8 or higher
- Pip (Python package manager)

## 📥 Setup (Step-by-Step)

### 1. Clone the repository
```bash
git clone https://github.com/your-username/Biodata-system.git
cd Biodata-system/biodataa
```

### 2. Create a virtual environment
```bash
python -m venv venv
```

### 3. Activate the virtual environment

**Windows:**
```bash
venv\Scripts\activate
```

**Mac / Linux:**
```bash
source venv/bin/activate
```

### 4. Install dependencies
```bash
pip install -r requirements.txt
```

### 5. Run the application
```bash
python app.py
```

### 6. Open in browser
Go to: [http://localhost:5000](http://localhost:5000)

The database and admin account are created automatically on first run.

## 🔑 Default Login

| Role | Username | Password |
|------|----------|----------|
| Teacher (Admin) | `admin` | `admin123` |
| Student | Their Register Number (e.g., `TL23BTCS0218`) | Same as register number in lowercase |

> **Note:** Student accounts are created automatically when the teacher uploads data.

## 📝 How to Use (For Teachers)

1. Login with **admin / admin123**
2. Go to **Upload Data**
3. Select the **Semester** (S1 - S8)
4. Select the **Data Type**:
   - `End Semester` — for KTU results (grades, SGPA, CGPA)
   - `Series 1` — for Series 1 internal marks
   - `Series 2` — for Series 2 internal marks
   - `Internals` — for internal assessment with attendance
5. Upload your Excel/PDF files and click **Process**
6. View results in **Dashboard** or **Attendance** page

## 📁 Project Structure

```
biodataa/
├── app.py              — Main Flask app (routes, models, config)
├── requirements.txt    — Python dependencies
├── README.md           — This file
├── .gitignore          — Git ignore rules
├── utils/
│   └── parsers.py      — Data parser (reads Excel/PDF, saves to DB)
├── templates/
│   ├── base.html       — Base layout template
│   ├── login.html      — Login page
│   ├── teacher/        — Teacher pages (dashboard, upload, etc.)
│   └── student/        — Student pages (profile, records)
├── static/
│   └── css/            — Stylesheets
└── uploads/            — Temporary storage for uploaded files
```

## 🔧 Tech Stack

- **Backend**: Python, Flask
- **Database**: SQLite (via Flask-SQLAlchemy)
- **Authentication**: Flask-Login
- **Data Parsing**: Pandas, OpenPyXL, PDFPlumber
