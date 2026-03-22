"""
=============================================================
 ACADEMIC DATA PARSER (parsers.py)
=============================================================
 This file handles reading uploaded Excel and PDF files,
 extracting student data (marks, attendance, grades),
 and saving it into the database.

 HOW IT WORKS (Big Picture):
   1. Teacher uploads Excel/PDF files via the web interface
   2. This parser reads each file and identifies:
      - Student IDs (TL..., VAS..., LVAS... numbers)
      - Subject codes (like CS301, EC201, etc.)
      - Marks, grades, attendance values
   3. It matches students to their database profiles
   4. It saves/updates the academic records

 SUPPORTED FILE TYPES:
   - Excel (.xlsx, .xls, .csv)
   - PDF (.pdf)

 SUPPORTED DATA TYPES:
   - 'end_sem'   : End semester exam results (grades, SGPA, CGPA)
   - 'series_1'  : Series 1 internal exam marks
   - 'series_2'  : Series 2 internal exam marks
   - 'internals' : Internal marks with attendance data
=============================================================
"""

import os
import copy
import pandas as pd        # For reading Excel/CSV files into tables
import pdfplumber           # For extracting text and tables from PDFs
import re                   # For pattern matching (regex)
import json


# ===================== HELPER FUNCTION =====================

def extract_text_from_pdf(file_path):
    """
    Reads ALL text from a PDF file (page by page).
    Used as a fallback when table extraction doesn't work well.

    Args:
        file_path: Path to the PDF file

    Returns:
        A single string containing all the text from the PDF
    """
    text = ""
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text


# ===================== MAIN FUNCTION =====================

def process_academic_files(file_paths, semester, data_type, db,
                           StudentProfile, AcademicRecord,
                           User=None, generate_password_hash=None):
    """
    MAIN ENTRY POINT: Processes all uploaded files and saves data to the database.

    Args:
        file_paths:              List of file paths to process
        semester:                Selected semester (e.g., 'S3')
        data_type:               Type of data ('end_sem', 'series_1', 'series_2', 'internals')
        db:                      Database object (for saving data)
        StudentProfile:          StudentProfile model class
        AcademicRecord:          AcademicRecord model class
        User:                    User model class (for creating student login accounts)
        generate_password_hash:  Function to hash passwords

    Returns:
        Tuple of (success: bool, message: str)
    """
    print(f"DEBUG: Processing {len(file_paths)} files for {semester} ({data_type})")

    try:
        # ========== STEP 1: READ ALL FILES AND EXTRACT DATA ==========
        #
        # We store extracted data in a nested dictionary:
        #   data_map_by_sem = {
        #       "S3": {
        #           "TL23BTCS0218": {
        #               "marks": {"CS301": {"series_1": "45", "_label": "Data Structures"}},
        #               "attendance": 85.5,
        #               "name": "John Doe",
        #               "admin_no": "TL23BTCS0218",
        #               "univ_no": "VAS23CS0115"
        #           },
        #           "TL23BTCS0219": { ... },
        #       },
        #       "S4": { ... }
        #   }
        data_map_by_sem = {}

        for path in file_paths:
            print(f"DEBUG: Reading file: {os.path.basename(path)}")

            # --- Auto-detect semester from filename ---
            # If the filename contains "S3", "S5", etc., use that instead of the dropdown value
            file_sem = semester  # Default: use what the teacher selected
            basename = os.path.basename(path).upper()
            sem_match = re.search(r'\b(S[1-8])\b', basename)
            if sem_match:
                file_sem = sem_match.group(1)

            # Initialize the semester bucket if it doesn't exist
            if file_sem not in data_map_by_sem:
                data_map_by_sem[file_sem] = {}
            data_map = data_map_by_sem[file_sem]

            # --- Process EXCEL / CSV files ---
            if path.endswith(('.xlsx', '.xls', '.csv', '.xsl')):
                try:
                    # Read the file into a pandas DataFrame (a table)
                    if path.endswith('.csv'):
                        df = pd.read_csv(path, header=None)
                    else:
                        df = pd.read_excel(path, header=None)

                    print(f"DEBUG: Excel shape: {df.shape}")

                    # Extract student data from the table
                    success, _ = process_dataframe(df, data_map, data_type)
                    print(f"DEBUG: process_dataframe success: {success}")

                except Exception as e:
                    print(f"Error reading {path}: {e}")

            # --- Process PDF files ---
            elif path.endswith('.pdf'):
                print(f"DEBUG: Processing PDF...")
                with pdfplumber.open(path) as pdf:
                    last_context = None  # Used to carry over column info between pages

                    for page in pdf.pages:
                        # Try to extract structured tables first
                        tables = page.extract_tables()
                        success = False

                        for table in tables:
                            if table and len(table) > 1:
                                # Convert extracted table to a DataFrame
                                df = pd.DataFrame(table)
                                ok, ctx = process_dataframe(df, data_map, data_type, last_context)
                                if ok:
                                    success = True
                                    last_context = ctx  # Remember column positions for next page

                        # --- Fallback: If table extraction failed, try raw text ---
                        if not success:
                            text = page.extract_text() or ""

                            # Look for patterns like: TL23BTCS0218 CS301(A+) EC201(B)
                            pattern = r'\b((?:TL|VAS|LVAS|AJC)\d{2}[A-Z]{2,4}\d{3,4})\b\s+([\w\d\(\)\s,]+)'
                            matches = re.findall(pattern, text)

                            for found_id, results in matches:
                                fid = found_id.strip()

                                # Skip header text that looks like IDs
                                if any(x in fid.upper() for x in ['RESULT', 'SEMESTER', 'COURSE', 'TOTAL', 'NAME']):
                                    continue
                                # Skip students from other batches (2021, 2022)
                                if any(x in fid.upper() for x in ['VAS22', 'LVAS22', 'TL22', 'VAS21', 'LVAS21', 'TL21']):
                                    continue

                                # Create entry for this student if not already present
                                if fid not in data_map:
                                    data_map[fid] = {
                                        'marks': {}, 'attendance': 0,
                                        'name': None, 'admin_no': None, 'univ_no': None
                                    }

                                # Identify which type of ID this is
                                if fid.upper().startswith(('VAS', 'LVAS')):
                                    data_map[fid]['univ_no'] = fid      # University number
                                else:
                                    data_map[fid]['admin_no'] = fid      # Admission number

                                # Extract subject grades from the text after the ID
                                # Pattern: CS301(A+) or EC201(B)
                                mark_matches = re.findall(r'([A-Z]{2,3}\d{3})\s*[\(]?([A-Z\d\.+]+)[\)]?', results)
                                for subj_code, grade in mark_matches:
                                    if subj_code not in data_map[fid]['marks']:
                                        data_map[fid]['marks'][subj_code] = {'_label': subj_code}

                                    grade_clean = str(grade).strip()
                                    if grade_clean.upper() not in ['NAN', 'NONE', '']:
                                        data_map[fid]['marks'][subj_code][data_type] = grade_clean

        # ========== STEP 2: SAVE EXTRACTED DATA TO DATABASE ==========
        print(f"DEBUG: Multi-semester Sync Started")
        updated_count = 0

        try:
            for sem_key, sem_data_map in data_map_by_sem.items():
                print(f"DEBUG: Syncing {len(sem_data_map)} IDs into semester {sem_key}")

                for found_id, data in sem_data_map.items():
                    admin_no = data.get('admin_no')   # TL... number
                    univ_no = data.get('univ_no')     # VAS... number

                    # Skip if we have no usable ID
                    if not admin_no and not univ_no:
                        print(f"DEBUG: Skipping ID {found_id} - no admin/univ number")
                        continue

                    # --- 2a. Find or Create the Student Profile ---
                    student = _find_student(StudentProfile, admin_no, univ_no)

                    if not student:
                        # For internals upload, don't create new students
                        # (they should already exist from an earlier upload)
                        if data_type == 'internals':
                            print(f"DEBUG: Skipping ID {found_id} - not in database for internals upload")
                            continue

                        # Auto-create a new student profile
                        reg_to_use = admin_no or univ_no
                        student = StudentProfile(
                            reg_no=reg_to_use,
                            univ_no=univ_no,
                            name=data.get('name') or reg_to_use
                        )
                        db.session.add(student)
                        db.session.flush()  # Get the auto-generated ID without committing yet

                    else:
                        # Update existing student's IDs if we found better ones
                        _update_student_ids(student, admin_no, univ_no, data)
                        db.session.flush()

                    # --- 2b. Create Login Account for Student ---
                    if User and generate_password_hash:
                        user = User.query.filter_by(username=student.reg_no).first()
                        if not user:
                            # Default password = register number in lowercase
                            user = User(
                                username=student.reg_no,
                                password=generate_password_hash(student.reg_no.lower()),
                                role='student'
                            )
                            db.session.add(user)
                            db.session.flush()

                    # --- 2c. Create or Update Academic Record ---
                    record = AcademicRecord.query.filter_by(
                        student_id=student.id, semester=sem_key
                    ).first()

                    if not record:
                        record = AcademicRecord(student_id=student.id, semester=sem_key)
                        db.session.add(record)

                    # Calculate average attendance from internal marks
                    if data_type == 'internals' and 'marks' in data:
                        att_values = []
                        for code, marks_dict in data['marks'].items():
                            if 'attendance' in marks_dict:
                                try:
                                    att_values.append(float(marks_dict['attendance']))
                                except (ValueError, TypeError):
                                    pass
                        if att_values:
                            avg_att = sum(att_values) / len(att_values)
                            if avg_att > 0:
                                data['attendance'] = round(avg_att, 2)

                    # --- 2d. Merge Marks Into the Record ---
                    has_actual_data = False

                    # Update attendance if we have a valid value
                    if 'attendance' in data and data['attendance'] > 0:
                        record.attendance_percentage = float(data['attendance'])
                        has_actual_data = True

                    # Merge new marks into existing marks (don't overwrite old data)
                    current_marks = copy.deepcopy(record.internal_marks_json or {})

                    if 'marks' in data:
                        for code, metric_dict in data['marks'].items():
                            # Create subject entry if it doesn't exist
                            if code not in current_marks or not isinstance(current_marks[code], dict):
                                current_marks[code] = {}

                            # Update label (keep the longer/better one)
                            if '_label' in metric_dict:
                                old_label = current_marks[code].get('_label', '')
                                new_label = metric_dict.pop('_label')
                                if len(new_label) >= len(old_label):
                                    current_marks[code]['_label'] = new_label

                            # Add each mark type (series_1, end_sem, etc.)
                            for mark_type, mark_value in metric_dict.items():
                                val_clean = str(mark_value).strip()
                                if val_clean.upper() not in ['NAN', 'NONE', '']:
                                    current_marks[code][mark_type] = val_clean
                                    has_actual_data = True

                    # Only save if we actually found real data
                    if has_actual_data:
                        # Save SGPA and CGPA if available
                        if 'sgpa' in data:
                            record.sgpa = data['sgpa']
                        if 'cgpa' in data:
                            record.cgpa = data['cgpa']

                        # Tell SQLAlchemy that the JSON field changed
                        # (needed because it doesn't auto-detect changes inside JSON)
                        from sqlalchemy.orm.attributes import flag_modified
                        record.internal_marks_json = current_marks
                        flag_modified(record, "internal_marks_json")
                        updated_count += 1

            # Save everything to the database
            db.session.commit()
            return True, "Data integrated successfully."

        except Exception as e:
            db.session.rollback()  # Undo all changes if something went wrong
            return False, f"Database Error: {str(e)}"

    except Exception as e:
        return False, str(e)


# ===================== HELPER: FIND STUDENT =====================

def _find_student(StudentProfile, admin_no, univ_no):
    """
    Search for a student in the database using their admin number or university number.
    Tries both IDs to find a match.

    Args:
        StudentProfile: The StudentProfile model class
        admin_no: Admission number (TL...)
        univ_no: University number (VAS... / LVAS...)

    Returns:
        StudentProfile object if found, None otherwise
    """
    student = None
    search_ids = [admin_no, univ_no]

    for sid in search_ids:
        if sid:
            # Search by both reg_no and univ_no fields
            student = StudentProfile.query.filter(
                (StudentProfile.reg_no == sid) | (StudentProfile.univ_no == sid)
            ).first()
            if student:
                break

    return student


# ===================== HELPER: UPDATE STUDENT IDs =====================

def _update_student_ids(student, admin_no, univ_no, data):
    """
    Update a student's ID numbers and name if we found better information.

    For example:
    - A student might initially be stored with VAS number as reg_no.
      When we find their TL number, we move VAS to univ_no and set TL as reg_no.
    - A student named 'TL23BTCS0218' gets renamed to their real name.

    Args:
        student: The StudentProfile object to update
        admin_no: New admission number (may be None)
        univ_no: New university number (may be None)
        data: The extracted data dictionary
    """
    # Update admission number (TL number)
    if admin_no and student.reg_no != admin_no:
        # If reg_no is currently a VAS number and we found a TL number, swap them
        if student.reg_no.startswith('VAS') and admin_no.startswith('TL'):
            if not student.univ_no:
                student.univ_no = student.reg_no  # Move VAS to univ_no
            student.reg_no = admin_no              # Set TL as reg_no
        elif not student.reg_no:
            student.reg_no = admin_no

    # Update university number (VAS/LVAS number)
    if univ_no and student.univ_no != univ_no:
        if not student.univ_no:
            student.univ_no = univ_no

    # Update name if current name is just an ID or missing
    name_is_placeholder = (
        student.name == student.reg_no or
        not student.name or
        student.name.startswith('TL') or
        student.name.startswith('VAS')
    )
    if name_is_placeholder and data.get('name'):
        student.name = data.get('name')


# ===================== DATAFRAME PROCESSOR =====================

def process_dataframe(df, data_map, data_type, last_context=None):
    """
    Process a single DataFrame (table) to extract student marks and attendance.

    This is where the core parsing logic lives. It:
    1. Finds the header row (which row contains column names)
    2. Identifies which columns contain student IDs, names, and subjects
    3. Reads each data row and maps marks to the correct student and subject

    Args:
        df:            A pandas DataFrame containing the table data
        data_map:      Dictionary to store extracted data (modified in place)
        data_type:     Type of data being processed
        last_context:  Column info from previous table (for multi-page PDFs)

    Returns:
        Tuple of (success: bool, context: dict)
        The context contains column positions so the next table can continue
    """

    # --- Column position trackers ---
    header_idx = -1    # Which row is the header row?
    admin_col = -1     # Which column has TL... numbers?
    univ_col = -1      # Which column has VAS... numbers?
    name_col = -1      # Which column has student names?
    subject_map = {}   # Maps column index → subject name (e.g., {5: "CS301 Data Structures"})
    headers = []       # List of header values from the header row
    is_continuation = False  # Is this a continuation of a previous table?

    # ========== HANDLE CONTINUATION TABLES ==========
    # In multi-page PDFs, a table might continue from the previous page.
    # If the current table has no subject headers but we have context from before,
    # we reuse the previous column positions.

    if last_context:
        has_subjects = False
        # Check first 3 rows for subject codes
        for i in range(min(3, len(df))):
            row_str = " ".join([str(v) for v in df.iloc[i].values]).upper()
            # Remove student IDs from the text so they don't match as subjects
            row_str_no_id = re.sub(
                r'((?:TL|VAS|LVAS|AJC)\d{2}[A-Z]{2,4}\d{2,4})', '', row_str, flags=re.IGNORECASE
            )
            # Look for subject code patterns (e.g., CS301, EC201)
            matches = re.findall(r'\b([A-Z]{2,4}\s?\d{3,4})\b', row_str_no_id)
            if matches:
                has_subjects = True
                break

        if not has_subjects:
            # No subject headers found → this is a continuation table
            print("DEBUG: Processing as continuation table!")
            admin_col = last_context['admin_col']
            univ_col = last_context['univ_col']
            name_col = last_context['name_col']
            subject_map = last_context['subject_map']
            headers = last_context['headers']
            header_idx = -1
            is_continuation = True

    # ========== FIND COLUMN POSITIONS (Normal Tables) ==========

    if not is_continuation:
        # Regex patterns to identify ID columns
        tl_pattern = re.compile(r'TL\d{2}[A-Z]{2,4}\d{2,4}', re.I)      # Matches TL23BTCS0218
        vas_pattern = re.compile(r'(?:VAS|LVAS)\d{2}[A-Z]{2,4}\d{2,4}', re.I)  # Matches VAS23CS0115

        # --- Scan rows to find header row and ID columns ---
        print("DEBUG: Scanning for headers and ID columns...")
        for i in range(min(100, len(df))):
            row = [str(v).strip() for v in df.iloc[i].values]
            row_str = " ".join(row).upper()

            # Look for ID values in each cell to identify ID columns
            for j, val in enumerate(row):
                v_upper = val.upper()
                if (tl_pattern.search(v_upper) or 'TL23' in v_upper) and admin_col == -1:
                    admin_col = j
                    print(f"DEBUG: Found likely Admin No column at index {j} (sample: {val})")
                if (vas_pattern.search(v_upper) or 'VAS23' in v_upper or 'LVAS' in v_upper) and univ_col == -1:
                    univ_col = j
                    print(f"DEBUG: Found likely Univ No column at index {j} (sample: {val})")

            # Detect header row by looking for keywords
            strong_match = any(kw in row_str for kw in [
                'STUDENT ID', 'UNIVERSITY REG NO', 'UNIVERSITY REG',
                'REG NO', 'ROLL NO', 'ADMISSION NO', 'NAME OF STUDENT'
            ])
            weak_match = ('NAME' in row_str and ('REG' in row_str or 'ID' in row_str or 'NO' in row_str))

            if strong_match or weak_match:
                if header_idx == -1:
                    header_idx = i
                    print(f"DEBUG: Found header row at index {i}")
                # Find the Name column
                for j, val in enumerate(row):
                    if 'NAME' in val.upper() and name_col == -1:
                        name_col = j
                        print(f"DEBUG: Found Name column at index {j}")

        # If no header row was found, assume row 0 is the header
        if header_idx == -1:
            header_idx = 0

        headers = [str(v).strip().upper() for v in df.iloc[header_idx].values]

        # --- Fallback: Try to find ID columns from header text ---
        if admin_col == -1:
            for j, h in enumerate(headers):
                if 'ADMISSION' in h or 'STUDENT ID' in h:
                    admin_col = j
                    break
        if univ_col == -1:
            for j, h in enumerate(headers):
                if 'UNIVERSITY' in h or 'REG' in h:
                    univ_col = j
                    break
        if name_col == -1:
            for j, h in enumerate(headers):
                if 'NAME' in h:
                    name_col = j
                    break

        # If we couldn't find ANY ID column, this table can't be processed
        if admin_col == -1 and univ_col == -1:
            print("DEBUG: Could not find ID column (TL/VAS/LVAS). Integration failed for this file.")
            return False, None

        print(f"DEBUG: Mapping subjects. AdminCol={admin_col}, UnivCol={univ_col}, NameCol={name_col}")

        # --- Map Subject Columns ---
        # Look in header rows for subject codes (e.g., "CS301 Data Structures")
        for i in range(min(header_idx + 1, len(df))):
            row = [str(v).strip() for v in df.iloc[i].values]
            for j, val in enumerate(row):
                v_upper = val.upper()
                # Remove student IDs so they don't confuse subject detection
                v_no_id = re.sub(
                    r'((?:TL|VAS|LVAS|AJC)\d{2}[A-Z]{2,4}\d{2,4})', '', v_upper, flags=re.IGNORECASE
                )
                match = re.search(r'\b([A-Z]{2,4}\s?\d{3,4})\b', v_no_id)
                if match:
                    val_clean = match.group(1).replace(" ", "")
                    # Keep the longer label (more descriptive)
                    prev_label = subject_map.get(j, "")
                    if len(val) >= len(prev_label):
                        subject_map[j] = val

        # --- Fill gaps: If a subject spans multiple columns (merged headers) ---
        # For example: "CS301 Data Structures" might cover columns 5, 6, 7
        # Column 5 has the subject, 6 and 7 are sub-columns (marks, grade, etc.)
        last_subj = None
        for j in range(len(df.columns)):
            if j in subject_map:
                last_subj = subject_map[j]
            elif last_subj:
                # Stop extending if we hit a special column
                curr_h = headers[j] if j < len(headers) else ""
                if any(kw in curr_h for kw in ['TOTAL', 'FAILED', 'RESULT', 'SGPA', 'CGPA', 'REMARK', 'COURSE']):
                    last_subj = None
                    continue
                # Don't assign subjects to ID or name columns
                if j != admin_col and j != univ_col and j != name_col:
                    subject_map[j] = last_subj

    # Save column positions so the next table can use them
    ctx = {
        'admin_col': admin_col,
        'univ_col': univ_col,
        'name_col': name_col,
        'subject_map': subject_map,
        'headers': headers
    }

    # ========== PROCESS DATA ROWS ==========
    # Now read each row (after the header) and extract student data

    for i in range(header_idx + 1, len(df)):
        row = df.iloc[i]

        # --- Extract and validate Student IDs ---
        adm_val_raw = str(row[admin_col]).strip().replace('\n', '').replace('\r', '') if admin_col != -1 else None
        univ_val_raw = str(row[univ_col]).strip().replace('\n', '').replace('\r', '') if univ_col != -1 else None

        # Extract the actual ID using regex (handles messy data)
        adm_val = None
        univ_val = None

        if adm_val_raw:
            m = re.search(r'((?:TL|VAS|LVAS|AJC)\d{2}[A-Z]{2,4}\d{2,4})', adm_val_raw, re.IGNORECASE)
            if m:
                adm_val = m.group(1).upper()

        if univ_val_raw:
            m = re.search(r'((?:TL|VAS|LVAS|AJC)\d{2}[A-Z]{2,4}\d{2,4})', univ_val_raw, re.IGNORECASE)
            if m:
                univ_val = m.group(1).upper()

        # Skip empty or too-short IDs
        if adm_val and (adm_val.upper() in ['NAN', 'NONE', ''] or len(adm_val) < 5):
            adm_val = None
        if univ_val and (univ_val.upper() in ['NAN', 'NONE', ''] or len(univ_val) < 5):
            univ_val = None

        # Need at least one valid ID to continue
        key_id = adm_val or univ_val
        if not key_id:
            continue

        # Skip students from other batches (2021, 2022)
        if any(x in str(key_id).upper() for x in ['VAS22', 'LVAS22', 'TL22', 'VAS21', 'LVAS21', 'TL21']):
            continue

        # Skip rows that look like headers (sometimes headers repeat in the data)
        row_full_text = " ".join([str(v) for v in row.values]).upper()
        if any(kw in row_full_text for kw in ['STUDENT ID', 'NAME OF STUDENT', 'REGISTER NO', 'ROLL NO']):
            continue

        # --- Create entry for this student ---
        if key_id not in data_map:
            data_map[key_id] = {
                'marks': {}, 'attendance': 0,
                'name': None, 'admin_no': adm_val, 'univ_no': univ_val
            }

        if adm_val:
            data_map[key_id]['admin_no'] = adm_val
        if univ_val:
            data_map[key_id]['univ_no'] = univ_val

        # Extract student name
        if name_col != -1:
            name_val = str(row[name_col]).strip().replace('\n', ' ').replace('\r', ' ')
            if name_val and name_val.upper() not in ['NAN', 'NONE'] and len(name_val) > 2:
                data_map[key_id]['name'] = name_val

        print(f"  > Found ID: {key_id}")

        # --- Process each column for marks ---
        for j, h in enumerate(headers):
            val = row[j]
            v_str = str(val).strip()

            # Skip empty values
            if not v_str or v_str.upper() in ['NAN', 'NONE', '']:
                continue

            h_clean = h.upper() if h else ""

            # --- Special handling for SERIES marks ---
            if data_type in ['series_1', 'series_2']:
                # Skip attendance columns in series reports
                if 'ATTENDANCE' in h_clean or 'ATTN' in h_clean:
                    continue
                # Skip columns marked as 0-mark subjects
                if '(0)' in h_clean:
                    continue

                # For Series 1: skip columns that are clearly for Series 2
                if data_type == 'series_1':
                    if '2' in h_clean and '1' not in h_clean and 'II' not in h_clean:
                        continue
                # For Series 2: skip columns that are clearly for Series 1
                elif data_type == 'series_2':
                    if '1' in h_clean and '2' not in h_clean and ' I' not in h_clean:
                        continue

            # --- Capture attendance for INTERNALS reports ---
            if data_type == 'internals' and ('ATTENDANCE' in h_clean or 'ATTN' in h_clean):
                try:
                    att = float(val)
                    if att > data_map[key_id]['attendance']:
                        data_map[key_id]['attendance'] = att
                except (ValueError, TypeError):
                    pass

            # --- Map value to subject ---
            subj_full = subject_map.get(j)
            if subj_full:
                s_upper = subj_full.upper()

                # Skip zero-mark and batch-specific subjects for series
                if data_type in ['series_1', 'series_2']:
                    if '(0)' in s_upper or 'BATCH1' in s_upper or 'EET283' in s_upper:
                        continue

                # Extract the subject code from the full subject name
                # e.g., "CS301 Data Structures" → "CS301"
                s_clean = re.sub(
                    r'((?:TL|VAS|LVAS|AJC)\d{2}[A-Z]{2,4}\d{2,4})', '',
                    subj_full.upper(), flags=re.IGNORECASE
                )
                code_match = re.search(r'\b([A-Z]{2,3}\d{3})\b', s_clean.replace(' ', ''))

                if code_match:
                    code = code_match.group(1)

                    # Create subject entry if it doesn't exist
                    if code not in data_map[key_id]['marks']:
                        data_map[key_id]['marks'][code] = {'_label': subj_full}

                    # Determine what type of mark this is
                    metric_key = data_type
                    if data_type == 'internals' and ('ATTENDANCE' in h_clean or 'ATTN' in h_clean):
                        metric_key = 'attendance'

                    # Handle special PDF format: "A: 100\nI: 40\nE: Yes"
                    # (combines attendance, internals, and eligibility in one cell)
                    if data_type == 'internals' and 'A:' in v_str and 'I:' in v_str:
                        a_match = re.search(r'A:\s*([\d\.]+)', v_str)
                        if a_match:
                            data_map[key_id]['marks'][code]['attendance'] = a_match.group(1)
                        i_match = re.search(r'I:\s*([\d\.]+)', v_str)
                        if i_match:
                            data_map[key_id]['marks'][code]['internals'] = i_match.group(1)
                    else:
                        data_map[key_id]['marks'][code][metric_key] = v_str

            # --- Extract SGPA / CGPA for end-semester results ---
            if data_type == 'end_sem':
                if 'SGPA' in h_clean:
                    try:
                        data_map[key_id]['sgpa'] = float(val)
                    except (ValueError, TypeError):
                        pass
                elif 'CGPA' in h_clean:
                    try:
                        data_map[key_id]['cgpa'] = float(val)
                    except (ValueError, TypeError):
                        pass

    return True, ctx
