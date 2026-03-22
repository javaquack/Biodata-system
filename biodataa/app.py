"""
=============================================================
 ACADEMIC BIODATA SYSTEM - Main Application (app.py)
=============================================================
 This is the main file that runs our Flask web application.
 It handles:
   1. Setting up the database (SQLite)
   2. Defining what data we store (Models)
   3. User login/logout (Authentication)
   4. All the web pages and their logic (Routes)
=============================================================
"""

# -------------------- IMPORTS --------------------
# These are the libraries (tools) we need to run our app

import os                          # For working with files and folders
import re                          # For pattern matching in text (e.g., extracting dept from reg no)
from flask import (                # Flask is our web framework
    Flask,                         # The main app object
    render_template,               # To show HTML pages
    request,                       # To read data from forms
    redirect,                      # To send user to another page
    url_for,                       # To generate URLs for our routes
    flash,                         # To show success/error messages
    jsonify                        # To send JSON responses
)
from flask_sqlalchemy import SQLAlchemy          # Database helper (makes SQL easy)
from flask_login import (                        # Login system
    LoginManager,                                # Manages user sessions
    UserMixin,                                   # Adds login methods to our User model
    login_user,                                  # Logs a user in
    login_required,                              # Protects pages (must be logged in)
    logout_user,                                 # Logs a user out
    current_user                                 # Gets the currently logged-in user
)
from werkzeug.security import generate_password_hash, check_password_hash  # Password encryption
from werkzeug.utils import secure_filename       # Sanitize uploaded file names
from datetime import datetime                    # For date/time operations


# ==================== APP SETUP ====================
# This section configures our Flask application

app = Flask(__name__)

# Secret key: Used to encrypt session cookies (keep this secret in production!)
app.config['SECRET_KEY'] = 'academic-biodata-secret-key-2024'

# Database location: We use SQLite, which stores everything in a single .db file
basedir = os.path.abspath(os.path.dirname(__file__))  # Get the folder this file is in
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(basedir, 'biodata_v2.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False   # Turns off unnecessary warnings

# Upload folder: Where uploaded Excel/PDF files are saved temporarily
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)  # Create folder if it doesn't exist

# Initialize database and login manager
db = SQLAlchemy(app)
print("Database connected:", app.config['SQLALCHEMY_DATABASE_URI'])

login_manager = LoginManager(app)
login_manager.login_view = 'login'  # If not logged in, redirect to login page


# ==================== DATABASE MODELS ====================
# Models define the STRUCTURE of our database tables.
# Think of each class as a table, and each variable as a column.


class User(UserMixin, db.Model):
    """
    USER TABLE: Stores login credentials for both teachers and students.
    - Teachers log in with a username (e.g., 'admin')
    - Students log in with their Register Number (e.g., 'TL23BTCS0218')
    """
    id = db.Column(db.Integer, primary_key=True)              # Unique ID (auto-generated)
    username = db.Column(db.String(50), unique=True, nullable=False)  # Login username
    password = db.Column(db.String(100), nullable=False)       # Hashed password (encrypted)
    role = db.Column(db.String(20), nullable=False)            # Either 'teacher' or 'student'


class StudentProfile(db.Model):
    """
    STUDENT PROFILE TABLE: Stores personal information about each student.
    Each student has ONE profile with their details.
    """
    id = db.Column(db.Integer, primary_key=True)               # Unique ID
    reg_no = db.Column(db.String(20), unique=True, nullable=False)   # Admission No (e.g., TL23BTCS0218)
    univ_no = db.Column(db.String(20), unique=True)            # University No (e.g., VAS23CS0115)
    name = db.Column(db.String(100), nullable=False)           # Student's full name

    # Personal details (can be updated by the student)
    father_name = db.Column(db.String(100))
    mother_name = db.Column(db.String(100))
    gender = db.Column(db.String(10))
    blood_group = db.Column(db.String(5))
    email = db.Column(db.String(100))
    phone = db.Column(db.String(15))
    address = db.Column(db.Text)
    date_of_birth = db.Column(db.String(20))
    admission_year = db.Column(db.Integer)

    # Relationship: One student can have MANY academic records (one per semester)
    # cascade="all, delete-orphan" means: if we delete a student, delete their records too
    records = db.relationship('AcademicRecord', backref='student', lazy=True, cascade="all, delete-orphan")


class AcademicRecord(db.Model):
    """
    ACADEMIC RECORD TABLE: Stores marks and attendance for each semester.
    Each student has one record PER semester (S1, S2, ... S8).
    """
    id = db.Column(db.Integer, primary_key=True)               # Unique ID
    student_id = db.Column(db.Integer, db.ForeignKey('student_profile.id'), nullable=False)  # Links to student
    semester = db.Column(db.String(5), nullable=False)         # Semester code: S1, S2, ... S8
    attendance_percentage = db.Column(db.Float, default=0.0)   # Overall attendance %
    internal_marks_json = db.Column(db.JSON)                   # Subject marks stored as JSON
    # Example JSON: {"CS301": {"_label": "Data Structures", "series_1": "45", "end_sem": "A+"}}
    result_status = db.Column(db.String(20))                   # Pass / Fail / Pending
    sgpa = db.Column(db.Float)                                 # Semester GPA
    cgpa = db.Column(db.Float)                                 # Cumulative GPA


# This function tells Flask-Login how to load a user from the database
@login_manager.user_loader
def load_user(user_id):
    """Load a user by their ID (called automatically by Flask-Login)."""
    return User.query.get(int(user_id))


# ==================== PARSER IMPORT ====================
# Import our data processing function from utils/parsers.py
# This function reads Excel/PDF files and saves data to the database
from utils.parsers import process_academic_files


# ==================== ROUTES (WEB PAGES) ====================
# Each @app.route defines a URL and what happens when someone visits it.


# ---------- HOME PAGE ----------
@app.route('/')
def index():
    """
    Home page: If user is logged in, go to dashboard. Otherwise, go to login.
    """
    if not current_user.is_authenticated:
        return redirect(url_for('login'))
    return redirect(url_for('dashboard'))


# ---------- LOGIN PAGE ----------
@app.route('/login', methods=['GET', 'POST'])
def login():
    """
    Login page: Shows login form (GET) or processes login attempt (POST).
    - Teachers log in with username + password
    - Students log in with their Register Number as both username and password
    """
    if request.method == 'POST':
        # Get the username and password from the form
        username = request.form.get('username')
        password = request.form.get('password')

        # Look up the user in the database
        user = User.query.filter_by(username=username).first()

        # Check if user exists AND password is correct
        if user and check_password_hash(user.password, password):
            login_user(user)                        # Log them in (creates a session)
            return redirect(url_for('dashboard'))    # Send to dashboard

        flash('Invalid Credentials', 'danger')       # Show error message

    return render_template('login.html')


# ---------- LOGOUT ----------
@app.route('/logout')
@login_required  # Must be logged in to log out
def logout():
    """Log the user out and redirect to login page."""
    logout_user()
    return redirect(url_for('login'))


# ---------- DASHBOARD ----------
@app.route('/dashboard')
@login_required
def dashboard():
    """
    Dashboard page: Shows different views based on role.
    - Teacher: See list of all students with their data for a selected semester
    - Student: See their own profile with personal details
    """
    if current_user.role == 'teacher':
        # ---------- TEACHER DASHBOARD ----------
        # Get the selected semester from URL (default: S3)
        sem = request.args.get('semester', 'S3')

        # Get ALL students from the database
        students = StudentProfile.query.all()

        # Build display data: For each student, get their record for the selected semester
        display_data = []
        for s in students:
            # Get the academic record for this specific semester
            rec = AcademicRecord.query.filter_by(student_id=s.id, semester=sem).first()

            # Find the latest CGPA (from the most recent semester that has one)
            latest_cgpa = None
            for r in sorted(s.records, key=lambda x: x.semester, reverse=True):
                if r.cgpa:
                    latest_cgpa = r.cgpa
                    break

            display_data.append({
                'profile': s,
                'record': rec,
                'latest_cgpa': latest_cgpa
            })

        return render_template('teacher/dashboard.html', data=display_data, current_sem=sem)

    else:
        # ---------- STUDENT DASHBOARD ----------
        # Find the student profile matching their login username (register number)
        student = StudentProfile.query.filter_by(reg_no=current_user.username).first()

        # Find their latest CGPA
        latest_cgpa = None
        for r in sorted(student.records, key=lambda x: x.semester, reverse=True):
            if r.cgpa:
                latest_cgpa = r.cgpa
                break

        # Extract department name from Register Number
        # Example: TL23BTCS0218 → "BTCS" → "Computer Science (Batch)"
        dept_name = "N/A"
        m = re.search(r'\d{2}([A-Z]{2,4})\d+', student.reg_no.upper())
        if m:
            # Map department codes to full names
            dept_map = {
                'CS': 'Computer Science & Engineering',
                'BTCS': 'Computer Science (Batch)',
                'EC': 'Electronics & Communication',
                'ME': 'Mechanical Engineering',
                'CE': 'Civil Engineering',
                'EE': 'Electrical & Electronics'
            }
            code = m.group(1)
            dept_name = dept_map.get(code, code)  # If code not found, show the code itself

        return render_template('student/profile.html', student=student,
                               latest_cgpa=latest_cgpa, dept_name=dept_name)


# ---------- FILE UPLOAD (Teacher Only) ----------
@app.route('/teacher/upload', methods=['GET', 'POST'])
@login_required
def upload_files():
    """
    Upload page: Teachers can upload Excel/PDF files containing student data.
    Supports different data types:
      - 'end_sem'   : End semester marks & grades
      - 'series_1'  : Series 1 internal marks
      - 'series_2'  : Series 2 internal marks
      - 'internals' : Internal assessment marks with attendance
    """
    # Only teachers can access this page
    if current_user.role != 'teacher':
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        # Get form data
        semester = request.form.get('semester')             # Which semester (S1-S8)
        data_type = request.form.get('data_type', 'end_sem')  # What type of marks
        files = request.files.getlist('files')               # The uploaded files

        # Check if any files were selected
        if not files or files[0].filename == '':
            flash('No files selected', 'warning')
            return redirect(request.url)

        # Save each uploaded file to the uploads folder
        file_paths = []
        for file in files:
            filename = secure_filename(file.filename)  # Remove dangerous characters from filename
            path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(path)
            file_paths.append(path)

        # Process the files using our parser
        # This reads the Excel/PDF files and saves data to the database
        success, message = process_academic_files(
            file_paths, semester, data_type,
            db, StudentProfile, AcademicRecord,
            User, generate_password_hash
        )

        # Show result message to the teacher
        if success:
            flash(f'{semester}: {message}', 'success')
        else:
            flash(f'Error processing: {message}', 'danger')

        return redirect(url_for('upload_files'))

    # GET request: Just show the upload form
    return render_template('teacher/upload.html')


# ---------- ATTENDANCE VIEW (Teacher Only) ----------
@app.route('/teacher/attendance')
@login_required
def attendance_view():
    """Show attendance records for all students in a selected semester."""
    if current_user.role != 'teacher':
        return redirect(url_for('dashboard'))

    sem = request.args.get('semester', 'S3')
    records = AcademicRecord.query.filter_by(semester=sem).all()
    return render_template('teacher/attendance.html', records=records, current_sem=sem)


# ---------- BIODATA LIST (Teacher Only) ----------
@app.route('/teacher/biodata')
@login_required
def biodata_list():
    """Show a list of all students with their personal details."""
    if current_user.role != 'teacher':
        return redirect(url_for('dashboard'))

    students = StudentProfile.query.all()
    return render_template('teacher/biodata.html', students=students)


# ---------- DELETE STUDENT (Teacher Only) ----------
@app.route('/teacher/delete_student/<int:id>', methods=['POST'])
@login_required
def delete_student(id):
    """
    Delete a student and their associated records.
    Also deletes their login account.
    """
    if current_user.role != 'teacher':
        return redirect(url_for('dashboard'))

    # Find the student (or show 404 error if not found)
    student = StudentProfile.query.get_or_404(id)
    reg_no = student.reg_no

    # Delete the student's login account too
    user = User.query.filter_by(username=reg_no).first()
    if user:
        db.session.delete(user)

    # Delete the student (this also deletes their academic records due to cascade)
    db.session.delete(student)
    db.session.commit()

    flash(f'Student {reg_no} and associated records deleted.', 'success')
    return redirect(url_for('biodata_list'))


# ---------- STUDENT DETAIL VIEW (Teacher Only) ----------
@app.route('/teacher/student/<int:id>')
@login_required
def student_detail(id):
    """Show detailed academic records for a specific student (all 8 semesters)."""
    if current_user.role != 'teacher':
        return redirect(url_for('dashboard'))

    student = StudentProfile.query.get_or_404(id)

    # Create a dictionary of records keyed by semester: {"S1": record, "S2": record, ...}
    records = {r.semester: r for r in student.records}

    # List of all 8 semesters
    semesters = [f'S{i}' for i in range(1, 9)]

    return render_template('teacher/student_detail.html', student=student,
                           records=records, semesters=semesters)


# ---------- STUDENT RECORDS VIEW (Student Only) ----------
@app.route('/student/records')
@login_required
def student_records():
    """Show academic records for the logged-in student (all 8 semesters)."""
    if current_user.role != 'student':
        return redirect(url_for('dashboard'))

    # Find the student profile using their login username (register number)
    student = StudentProfile.query.filter_by(reg_no=current_user.username).first()

    # Create a dictionary of records keyed by semester
    records = {r.semester: r for r in student.records}
    semesters = [f'S{i}' for i in range(1, 9)]

    return render_template('student/records.html', student=student,
                           records=records, semesters=semesters)


# ---------- UPDATE PROFILE (Student Only) ----------
@app.route('/student/update_profile', methods=['POST'])
@login_required
def update_profile():
    """Allow students to update their personal details (phone, email, etc.)."""
    if current_user.role != 'student':
        return jsonify({'status': 'error'}), 403

    # Find the student and update their details from the form
    student = StudentProfile.query.filter_by(reg_no=current_user.username).first()
    student.phone = request.form.get('phone')
    student.address = request.form.get('address')
    student.email = request.form.get('email')
    student.father_name = request.form.get('father_name')
    student.mother_name = request.form.get('mother_name')
    student.gender = request.form.get('gender')
    student.blood_group = request.form.get('blood_group')
    student.date_of_birth = request.form.get('date_of_birth')

    db.session.commit()  # Save changes to database
    flash('Profile updated successfully', 'success')
    return redirect(url_for('dashboard'))


# ==================== DATABASE INITIALIZATION ====================

def init_db():
    """
    Initialize the database when the app starts for the first time.
    - Creates all tables if they don't exist
    - Creates a default admin/teacher account (username: admin, password: admin123)
    """
    with app.app_context():
        db.create_all()  # Create all tables defined by our Models above

        # Create default admin account if it doesn't exist yet
        if not User.query.filter_by(username='admin').first():
            admin = User(
                username='admin',
                password=generate_password_hash('admin123'),
                role='teacher'
            )
            db.session.add(admin)
            db.session.commit()


# ==================== RUN THE APP ====================

if __name__ == '__main__':
    init_db()                        # Set up the database
    app.run(debug=True, port=5000)   # Start the web server on http://localhost:5000
