# ==============================================================================
# core/database.py  —  Student Result Analyzer (v3 — PDF-Accurate Schema)
#
# Architecture: 3 normalized SQL tables
#   1. import_sessions  — har PDF upload ka record (course, semester, year)
#   2. students         — har student ka ek row (name, sp_id, gender, sgpa...)
#   3. subject_marks    — har student ke har subject ke marks (ext, int, grade)
#
# UI ke liye backward-compatible view:
#   get_all_data()  → roll_no, name, sub_1..sub_7, total, sgpa, status, grade,
#                     atkt_count, college, percentage  (same columns as before)
#   get_stats()     → same dict keys as before (+ new keys like avg_sgpa, etc.)
# ==============================================================================

import sqlite3
import os
import traceback
import pandas as pd


SCHEMA_VERSION = 3

# ------------------------------------------------------------------------------
# SQL Definitions
# ------------------------------------------------------------------------------

SQL_CREATE_SESSIONS = """
CREATE TABLE IF NOT EXISTS import_sessions (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    pdf_filename    TEXT    DEFAULT '',
    course          TEXT    DEFAULT 'Unknown Course',
    semester        TEXT    DEFAULT 'Unknown Semester',
    academic_year   TEXT    DEFAULT '',
    college_name    TEXT    DEFAULT '',
    max_marks       INTEGER DEFAULT 700,
    subject_count   INTEGER DEFAULT 7,
    imported_at     TEXT    DEFAULT (datetime('now','localtime')),
    total_students  INTEGER DEFAULT 0
)
"""

SQL_CREATE_STUDENTS = """
CREATE TABLE IF NOT EXISTS students (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id      INTEGER REFERENCES import_sessions(id) ON DELETE CASCADE,
    seat_no         TEXT    NOT NULL,
    sp_id           TEXT    DEFAULT '',
    name            TEXT    NOT NULL,
    gender          TEXT    DEFAULT '',
    total_marks     INTEGER DEFAULT 0,
    percentage      REAL    DEFAULT 0.0,
    sgpa            TEXT    DEFAULT '--',
    overall_grade   TEXT    DEFAULT '--',
    overall_status  TEXT    DEFAULT 'UNKNOWN',
    atkt_count      INTEGER DEFAULT 0,
    college         TEXT    DEFAULT ''
)
"""

SQL_CREATE_SUBJECT_MARKS = """
CREATE TABLE IF NOT EXISTS subject_marks (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id      INTEGER REFERENCES students(id) ON DELETE CASCADE,
    subject_index   INTEGER NOT NULL,
    subject_name    TEXT    DEFAULT '',
    ext_mark        TEXT    DEFAULT '--',
    int_mark        TEXT    DEFAULT '--',
    total_mark      INTEGER DEFAULT 0,
    grade           TEXT    DEFAULT '--',
    status          TEXT    DEFAULT 'UNKNOWN'
)
"""

SQL_CREATE_META = """
CREATE TABLE IF NOT EXISTS meta (
    key     TEXT PRIMARY KEY,
    value   TEXT
)
"""

# View — UI ke liye backward compatible flat row
# roll_no = seat_no, sub_1..7 = subject totals, total = total_marks
SQL_CREATE_VIEW = """
CREATE VIEW IF NOT EXISTS v_students_flat AS
SELECT
    s.id,
    s.session_id,
    s.seat_no           AS roll_no,
    s.sp_id,
    s.name,
    s.gender,
    s.college,
    s.total_marks       AS total,
    s.percentage,
    s.sgpa,
    s.overall_grade     AS grade,
    s.overall_status    AS status,
    s.atkt_count,
    -- subject totals (sub_1 to sub_7)
    COALESCE((SELECT total_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=0),0) AS sub_1,
    COALESCE((SELECT total_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=1),0) AS sub_2,
    COALESCE((SELECT total_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=2),0) AS sub_3,
    COALESCE((SELECT total_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=3),0) AS sub_4,
    COALESCE((SELECT total_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=4),0) AS sub_5,
    COALESCE((SELECT total_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=5),0) AS sub_6,
    COALESCE((SELECT total_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=6),0) AS sub_7,
    -- subject EXT marks
    COALESCE((SELECT ext_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=0),'--') AS sub_1_ext,
    COALESCE((SELECT ext_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=1),'--') AS sub_2_ext,
    COALESCE((SELECT ext_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=2),'--') AS sub_3_ext,
    COALESCE((SELECT ext_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=3),'--') AS sub_4_ext,
    COALESCE((SELECT ext_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=4),'--') AS sub_5_ext,
    COALESCE((SELECT ext_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=5),'--') AS sub_6_ext,
    COALESCE((SELECT ext_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=6),'--') AS sub_7_ext,
    -- subject INT marks
    COALESCE((SELECT int_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=0),'--') AS sub_1_int,
    COALESCE((SELECT int_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=1),'--') AS sub_2_int,
    COALESCE((SELECT int_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=2),'--') AS sub_3_int,
    COALESCE((SELECT int_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=3),'--') AS sub_4_int,
    COALESCE((SELECT int_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=4),'--') AS sub_5_int,
    COALESCE((SELECT int_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=5),'--') AS sub_6_int,
    COALESCE((SELECT int_mark FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=6),'--') AS sub_7_int,
    -- subject grades
    COALESCE((SELECT grade FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=0),'--') AS sub_1_grade,
    COALESCE((SELECT grade FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=1),'--') AS sub_2_grade,
    COALESCE((SELECT grade FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=2),'--') AS sub_3_grade,
    COALESCE((SELECT grade FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=3),'--') AS sub_4_grade,
    COALESCE((SELECT grade FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=4),'--') AS sub_5_grade,
    COALESCE((SELECT grade FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=5),'--') AS sub_6_grade,
    COALESCE((SELECT grade FROM subject_marks sm WHERE sm.student_id=s.id AND sm.subject_index=6),'--') AS sub_7_grade,
    -- failed subjects as text (for export / display)
    COALESCE((
        SELECT GROUP_CONCAT(subject_name || ' (EXT ' || ext_mark || ', INT ' || int_mark || ', TOT ' || total_mark || ')', char(10))
        FROM subject_marks sm2
        WHERE sm2.student_id = s.id AND sm2.status = 'FAIL'
    ), '') AS failed_subjects
FROM students s
"""


# ==============================================================================
# Helper — safe integer conversion
# ==============================================================================
def _safe_int(val, default=0):
    try:
        s = str(val).strip()
        if s in ('-', '--', 'AB', 'ABSENT', 'ZR', ''):
            return default
        return int(float(s))
    except Exception:
        return default


def _safe_float(val, default=0.0):
    try:
        return float(str(val).strip())
    except Exception:
        return default


def _overall_grade_from_sgpa(sgpa_str: str) -> str:
    """SGPA string se overall grade determine karo (VNSGU scale)."""
    try:
        v = float(sgpa_str)
        if v >= 9.0:  return "O"
        if v >= 8.0:  return "A+"
        if v >= 7.0:  return "A"
        if v >= 6.0:  return "B+"
        if v >= 5.5:  return "B"
        if v >= 5.0:  return "C"
        return "F"
    except Exception:
        return "--"


def _subject_status(ext_mark_str: str, grade: str, subject_index: int,
                    ext_passing_min: list) -> str:
    """Ek subject PASS hai ya FAIL — grade aur EXT threshold se decide karo."""
    g = str(grade).upper().strip() if grade else "--"
    if g in ("F", "AB", "ABSENT"):
        return "FAIL"
    ext_val = _safe_int(ext_mark_str, default=-1)
    if ext_val == -1:
        return "FAIL"   # AB / missing
    try:
        min_ext = ext_passing_min[subject_index]
    except (IndexError, TypeError):
        min_ext = 18
    return "FAIL" if ext_val < min_ext else "PASS"


# ==============================================================================
# DatabaseManager
# ==============================================================================
class DatabaseManager:
    """
    Student Result Analyzer — PDF-accurate database layer.

    Tables:
        import_sessions  — ek row per PDF upload
        students         — ek row per student
        subject_marks    — 7 rows per student (one per subject)

    UI Compatibility:
        get_all_data()  returns a DataFrame with the same column names
        as the old flat schema (roll_no, sub_1..7, total, status, grade,
        sgpa, atkt_count, college, percentage) so existing UI code works
        without any changes.
    """

    def __init__(self, db_path: str = ":memory:"):
        self.db_path   = db_path
        self._session_id: int | None = None   # current import session
        self._connect()
        self._init_schema()

    # ------------------------------------------------------------------
    # Connection & Schema
    # ------------------------------------------------------------------

    def _connect(self):
        self.conn = sqlite3.connect(self.db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("PRAGMA foreign_keys=ON")
        self.cursor = self.conn.cursor()
        print(f"[DB] Connected -> {self.db_path}")

    def _init_schema(self):
        self.cursor.execute(SQL_CREATE_SESSIONS)
        self.cursor.execute(SQL_CREATE_STUDENTS)
        self.cursor.execute(SQL_CREATE_SUBJECT_MARKS)
        self.cursor.execute(SQL_CREATE_META)
        # View — drop & recreate so definition stays fresh
        self.cursor.execute("DROP VIEW IF EXISTS v_students_flat")
        self.cursor.execute(SQL_CREATE_VIEW)
        self.cursor.execute(
            "INSERT OR REPLACE INTO meta(key,value) VALUES('schema_version',?)",
            (str(SCHEMA_VERSION),)
        )
        self.conn.commit()
        print(f"[DB] Schema v{SCHEMA_VERSION} ready.")

    def close(self):
        try:
            self.conn.close()
            print("[DB] Connection closed.")
        except Exception:
            pass

    def __repr__(self):
        return f"<DatabaseManager path='{self.db_path}' students={self.get_total_count()}>"

    # ------------------------------------------------------------------
    # IMPORT — Session Start
    # ------------------------------------------------------------------

    def start_import_session(self, pdf_filename: str = "",
                             course: str = "Unknown Course",
                             semester: str = "Unknown Semester",
                             academic_year: str = "",
                             college_name: str = "",
                             max_marks: int = 700,
                             subject_count: int = 7) -> int:
        """
        Naya import session shuru karo.
        Pehle saara purana data delete karo (fresh import policy).
        Returns the new session_id.
        """
        # Fresh import — purna data clear
        self._clear_all_data()

        self.cursor.execute("""
            INSERT INTO import_sessions
                (pdf_filename, course, semester, academic_year,
                 college_name, max_marks, subject_count)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (pdf_filename, course, semester, academic_year,
              college_name, max_marks, subject_count))
        self.conn.commit()
        self._session_id = self.cursor.lastrowid
        print(f"[DB] Import session #{self._session_id} started "
              f"({course} / {semester} / {academic_year})")
        return self._session_id

    def _clear_all_data(self):
        """Saare import sessions, students, subject_marks delete karo."""
        self.cursor.execute("DELETE FROM subject_marks")
        self.cursor.execute("DELETE FROM students")
        self.cursor.execute("DELETE FROM import_sessions")
        self.conn.commit()
        self._session_id = None

    # ------------------------------------------------------------------
    # IMPORT — Students Insert (New Primary Method)
    # ------------------------------------------------------------------

    def insert_students(self,
                        students_list: list,
                        session_id: int,
                        subjects_map: dict,
                        ext_passing_min: list,
                        max_marks: int = 700) -> bool | str:
        """
        PDF parser ke raw student dicts ko database mein store karo.

        Args:
            students_list  : list of student dicts from pdf_parser.load_students()
            session_id     : ID from start_import_session()
            subjects_map   : {0: "Mathematics", 1: "Physics", ...} from SUBJECTS
            ext_passing_min: [18, 18, 18, 18, 9, 9, 9] from EXT_PASSING_MIN
            max_marks      : Maximum total marks (default 700)

        Returns:
            True on success, error string on failure.
        """
        if not students_list:
            return "No students to insert."

        try:
            subject_count = len(subjects_map) or 7

            for s in students_list:
                # ---- Student fields ----
                seat_no  = str(s.get('seat_no', 'N/A')).strip()
                sp_id    = str(s.get('sp_id',   '')).strip()
                name     = str(s.get('name',    'Unknown')).strip()
                gender   = str(s.get('gender',  '')).strip()
                college  = str(s.get('college', '')).strip()
                total    = _safe_int(s.get('total_marks', 0))

                raw_sgpa = s.get('sgpa', '--')
                try:
                    sgpa_val = float(raw_sgpa)
                    sgpa = str(raw_sgpa) if sgpa_val > 0.0 else '--'
                except Exception:
                    sgpa = '--'

                raw_status = str(s.get('overall_status', 'UNKNOWN')).upper()
                status = raw_status if raw_status in ('PASS', 'FAIL') else 'UNKNOWN'

                # Overall grade
                if status == 'FAIL':
                    overall_grade = 'F'
                elif sgpa != '--':
                    overall_grade = _overall_grade_from_sgpa(sgpa)
                else:
                    overall_grade = '--'

                # Percentage
                perc = round((total / max_marks) * 100, 2) if max_marks > 0 and total > 0 else 0.0

                # Per-subject data
                marks_map  = s.get('marks',  {})
                grades_map = s.get('grades', {})

                atkt_count = 0
                sub_rows   = []

                for idx in range(subject_count):
                    sub_name = str(subjects_map.get(idx, f"Subject {idx+1}")).strip()
                    m = marks_map.get(idx, ['-', '-'])
                    ext_str  = str(m[0]).strip() if len(m) > 0 else '-'
                    int_str  = str(m[1]).strip() if len(m) > 1 else '-'

                    ext_val  = _safe_int(ext_str)
                    int_val  = _safe_int(int_str)
                    tot_val  = ext_val + int_val

                    grade    = str(grades_map.get(idx, '--') or '--').strip()
                    sub_stat = _subject_status(ext_str, grade, idx, ext_passing_min)

                    if sub_stat == 'FAIL':
                        atkt_count += 1

                    sub_rows.append((idx, sub_name, ext_str, int_str, tot_val, grade, sub_stat))

                # ---- Insert student ----
                self.cursor.execute("""
                    INSERT INTO students
                        (session_id, seat_no, sp_id, name, gender,
                         total_marks, percentage, sgpa, overall_grade,
                         overall_status, atkt_count, college)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
                """, (session_id, seat_no, sp_id, name, gender,
                      total, perc, sgpa, overall_grade,
                      status, atkt_count, college))
                student_id = self.cursor.lastrowid

                # ---- Insert subject marks ----
                self.cursor.executemany("""
                    INSERT INTO subject_marks
                        (student_id, subject_index, subject_name,
                         ext_mark, int_mark, total_mark, grade, status)
                    VALUES (?,?,?,?,?,?,?,?)
                """, [(student_id, *row) for row in sub_rows])

            # Update total_students count in session
            self.cursor.execute("""
                UPDATE import_sessions
                SET total_students = (SELECT COUNT(*) FROM students WHERE session_id=?)
                WHERE id=?
            """, (session_id, session_id))
            self.conn.commit()
            print(f"[DB] {len(students_list)} students inserted for session #{session_id}.")
            return True

        except sqlite3.Error as e:
            traceback.print_exc()
            return str(e)

    # ------------------------------------------------------------------
    # LEGACY insert_data() — main.py ke purane PDFWorker ke liye
    # ------------------------------------------------------------------

    def insert_data(self, data: list) -> bool | str:
        """
        Legacy method — main.py ke old PDFWorker tuple-list format ko support karta hai.
        Tuple format:
            (name, roll_no, sub1..7, total, percentage, grade, status,
             sgpa, atkt_count, failed_subjects, college)
        """
        if not data:
            return "No data to insert."
        try:
            # Legacy session (no course info)
            self._clear_all_data()
            self.cursor.execute("""
                INSERT INTO import_sessions (pdf_filename, course, semester)
                VALUES ('legacy_import', 'Unknown', 'Unknown')
            """)
            self.conn.commit()
            session_id = self.cursor.lastrowid
            self._session_id = session_id

            for row in data:
                # Unpack legacy tuple
                (name, roll_no,
                 s1, s2, s3, s4, s5, s6, s7,
                 total, percentage, grade, status, sgpa,
                 atkt_count, failed_subjects, college) = row

                self.cursor.execute("""
                    INSERT INTO students
                        (session_id, seat_no, name, total_marks, percentage,
                         sgpa, overall_grade, overall_status, atkt_count, college)
                    VALUES (?,?,?,?,?,?,?,?,?,?)
                """, (session_id, str(roll_no), str(name), int(total or 0),
                      float(percentage or 0), str(sgpa), str(grade),
                      str(status), int(atkt_count or 0), str(college)))
                student_id = self.cursor.lastrowid

                # Insert subject marks from sub_1..7 (only total known here)
                for idx, tot in enumerate([s1,s2,s3,s4,s5,s6,s7]):
                    self.cursor.execute("""
                        INSERT INTO subject_marks
                            (student_id, subject_index, subject_name,
                             ext_mark, int_mark, total_mark, grade, status)
                        VALUES (?,?,?,?,?,?,?,?)
                    """, (student_id, idx, f"Subject {idx+1}",
                          '--', '--', int(tot or 0), '--', 'UNKNOWN'))

            self.cursor.execute("""
                UPDATE import_sessions
                SET total_students=(SELECT COUNT(*) FROM students WHERE session_id=?)
                WHERE id=?
            """, (session_id, session_id))
            self.conn.commit()
            print(f"[DB] Legacy insert: {len(data)} records.")
            return True
        except Exception as e:
            traceback.print_exc()
            return str(e)

    # ------------------------------------------------------------------
    # READ — UI Compatible
    # ------------------------------------------------------------------

    def get_all_data(self) -> pd.DataFrame:
        """
        Saare students ka flat DataFrame return karo.
        Columns include: roll_no, name, sub_1..7, total, sgpa,
                         status, grade, atkt_count, college, percentage,
                         sp_id, gender, sub_1_ext..7_ext, sub_1_int..7_int,
                         sub_1_grade..7_grade, failed_subjects
        """
        try:
            return pd.read_sql_query(
                "SELECT * FROM v_students_flat ORDER BY roll_no ASC",
                self.conn
            )
        except Exception as e:
            print(f"[DB] get_all_data Error: {e}")
            return pd.DataFrame()

    def get_student_by_seat(self, seat_no: str) -> pd.DataFrame:
        """Seat number se student dhundho."""
        try:
            return pd.read_sql_query(
                "SELECT * FROM v_students_flat WHERE roll_no=?",
                self.conn, params=(seat_no,)
            )
        except Exception as e:
            print(f"[DB] get_student_by_seat Error: {e}")
            return pd.DataFrame()

    def search_student(self, query: str) -> pd.DataFrame:
        """Naam, roll_no ya sp_id se search karo (partial match)."""
        try:
            q = f"%{query.strip()}%"
            return pd.read_sql_query(
                """SELECT * FROM v_students_flat
                   WHERE name LIKE ? OR roll_no LIKE ? OR sp_id LIKE ?
                   ORDER BY roll_no ASC""",
                self.conn, params=(q, q, q)
            )
        except Exception as e:
            print(f"[DB] search_student Error: {e}")
            return pd.DataFrame()

    def filter_by_status(self, status: str) -> pd.DataFrame:
        """PASS / FAIL / ALL filter."""
        try:
            if status.upper() == "ALL":
                return self.get_all_data()
            return pd.read_sql_query(
                "SELECT * FROM v_students_flat WHERE status=? ORDER BY roll_no ASC",
                self.conn, params=(status.upper(),)
            )
        except Exception as e:
            print(f"[DB] filter_by_status Error: {e}")
            return pd.DataFrame()

    def filter_by_college(self, college: str) -> pd.DataFrame:
        """College ke naam se filter karo."""
        try:
            return pd.read_sql_query(
                "SELECT * FROM v_students_flat WHERE college=? ORDER BY roll_no ASC",
                self.conn, params=(college,)
            )
        except Exception as e:
            print(f"[DB] filter_by_college Error: {e}")
            return pd.DataFrame()

    def get_atkt_students(self) -> pd.DataFrame:
        """ATKT / backlog wale students (atkt_count > 0)."""
        try:
            return pd.read_sql_query(
                """SELECT * FROM v_students_flat
                   WHERE atkt_count > 0
                   ORDER BY atkt_count DESC, roll_no ASC""",
                self.conn
            )
        except Exception as e:
            print(f"[DB] get_atkt_students Error: {e}")
            return pd.DataFrame()

    def get_top_students(self, n: int = 5) -> pd.DataFrame:
        """Top N students by total marks."""
        try:
            return pd.read_sql_query(
                "SELECT * FROM v_students_flat ORDER BY total DESC LIMIT ?",
                self.conn, params=(n,)
            )
        except Exception as e:
            return pd.DataFrame()

    def get_bottom_students(self, n: int = 5) -> pd.DataFrame:
        """Bottom N students by total marks."""
        try:
            return pd.read_sql_query(
                "SELECT * FROM v_students_flat WHERE status='FAIL' ORDER BY total ASC LIMIT ?",
                self.conn, params=(n,)
            )
        except Exception as e:
            return pd.DataFrame()

    def get_topper(self) -> dict | None:
        """Sabse zyada marks wala student."""
        try:
            self.cursor.execute(
                "SELECT * FROM v_students_flat ORDER BY total DESC LIMIT 1"
            )
            row = self.cursor.fetchone()
            return dict(row) if row else None
        except Exception as e:
            return None

    def get_subject_marks(self, student_id: int) -> pd.DataFrame:
        """Ek student ke saare subject marks (detailed)."""
        try:
            return pd.read_sql_query(
                "SELECT * FROM subject_marks WHERE student_id=? ORDER BY subject_index",
                self.conn, params=(student_id,)
            )
        except Exception as e:
            return pd.DataFrame()

    def get_subject_analysis(self, subject_index: int) -> dict:
        """
        Ek subject ka poora analysis:
            total, pass_count, fail_count, avg_ext, avg_int, avg_total
        """
        try:
            self.cursor.execute("""
                SELECT
                    subject_name,
                    COUNT(*) as total,
                    SUM(CASE WHEN status='PASS' THEN 1 ELSE 0 END) as pass_count,
                    SUM(CASE WHEN status='FAIL' THEN 1 ELSE 0 END) as fail_count,
                    ROUND(AVG(CASE WHEN ext_mark NOT IN ('-','--','AB') THEN CAST(ext_mark AS REAL) END), 1) as avg_ext,
                    ROUND(AVG(CASE WHEN int_mark NOT IN ('-','--','AB') THEN CAST(int_mark AS REAL) END), 1) as avg_int,
                    ROUND(AVG(total_mark), 1) as avg_total
                FROM subject_marks
                WHERE subject_index=?
                GROUP BY subject_index
            """, (subject_index,))
            row = self.cursor.fetchone()
            return dict(row) if row else {}
        except Exception as e:
            return {}

    # ------------------------------------------------------------------
    # STATISTICS
    # ------------------------------------------------------------------

    def get_stats(self) -> dict:
        """
        Summary statistics — UI ke liye.
        Returns dict with keys (backward compatible + extended):
            total_students, pass_count, fail_count, pass_perc, fail_perc,
            avg_percentage, topper, topper_name, topper_roll, topper_total,
            avg_sgpa, atkt_students, course, semester, academic_year
        """
        try:
            df = pd.read_sql_query(
                "SELECT * FROM v_students_flat", self.conn
            )
            if df.empty:
                return {}

            total      = len(df)
            pass_count = len(df[df['status'] == 'PASS'])
            fail_count = len(df[df['status'] == 'FAIL'])
            atkt_count = len(df[df['atkt_count'] > 0])
            avg_perc   = round(df['percentage'].mean(), 1)

            sgpa_vals  = pd.to_numeric(df['sgpa'], errors='coerce').dropna()
            avg_sgpa   = round(sgpa_vals.mean(), 2) if not sgpa_vals.empty else 0.0

            topper_idx  = df['total'].idxmax()
            topper_row  = df.loc[topper_idx]

            # Session info
            session_info = self.get_session_info()

            return {
                # --- backward compatible keys ---
                "total_students": total,
                "pass_count":     pass_count,
                "fail_count":     fail_count,
                "pass_perc":      round((pass_count / total) * 100, 1) if total > 0 else 0,
                "fail_perc":      round((fail_count / total) * 100, 1) if total > 0 else 0,
                "avg_percentage": avg_perc,
                "topper":         topper_row.get('name', 'N/A'),   # backward compat
                # --- extended keys ---
                "topper_name":    topper_row.get('name', 'N/A'),
                "topper_roll":    topper_row.get('roll_no', 'N/A'),
                "topper_total":   int(topper_row.get('total', 0)),
                "avg_sgpa":       avg_sgpa,
                "atkt_students":  atkt_count,
                "course":         session_info.get("course", ""),
                "semester":       session_info.get("semester", ""),
                "academic_year":  session_info.get("academic_year", ""),
            }
        except Exception as e:
            traceback.print_exc()
            return {}

    def get_grade_distribution(self) -> dict:
        """Grade-wise student count (overall_grade)."""
        try:
            self.cursor.execute("""
                SELECT overall_grade as grade, COUNT(*) as cnt
                FROM students
                GROUP BY overall_grade
                ORDER BY cnt DESC
            """)
            return {r['grade']: r['cnt'] for r in self.cursor.fetchall()}
        except Exception as e:
            return {}

    def get_subject_averages(self) -> dict:
        """Har subject ka average total marks."""
        try:
            self.cursor.execute("""
                SELECT subject_index, subject_name,
                       ROUND(AVG(total_mark), 1) as avg_total,
                       ROUND(AVG(CASE WHEN ext_mark NOT IN ('-','--','AB')
                                 THEN CAST(ext_mark AS REAL) END), 1) as avg_ext,
                       ROUND(AVG(CASE WHEN int_mark NOT IN ('-','--','AB')
                                 THEN CAST(int_mark AS REAL) END), 1) as avg_int
                FROM subject_marks
                GROUP BY subject_index
                ORDER BY subject_index
            """)
            rows = self.cursor.fetchall()
            result = {}
            for r in rows:
                key = f"sub_{r['subject_index']+1}"
                result[key] = {
                    "name":      r['subject_name'],
                    "avg_total": r['avg_total'],
                    "avg_ext":   r['avg_ext'],
                    "avg_int":   r['avg_int'],
                }
            return result
        except Exception as e:
            return {}

    def get_college_stats(self) -> pd.DataFrame:
        """College-wise pass/fail breakdown."""
        try:
            return pd.read_sql_query("""
                SELECT
                    college,
                    COUNT(*) as total_students,
                    SUM(CASE WHEN overall_status='PASS' THEN 1 ELSE 0 END) as pass_count,
                    SUM(CASE WHEN overall_status='FAIL' THEN 1 ELSE 0 END) as fail_count,
                    ROUND(
                        SUM(CASE WHEN overall_status='PASS' THEN 1.0 ELSE 0 END)
                        / COUNT(*) * 100, 1
                    ) as pass_percentage,
                    ROUND(AVG(CASE WHEN sgpa NOT IN ('--','') 
                              THEN CAST(sgpa AS REAL) END), 2) as avg_sgpa
                FROM students
                WHERE college != ''
                GROUP BY college
                ORDER BY pass_percentage DESC
            """, self.conn)
        except Exception as e:
            return pd.DataFrame()

    def get_percentage_distribution(self) -> dict:
        """Percentage range mein students ki count."""
        try:
            df = pd.read_sql_query(
                "SELECT percentage FROM v_students_flat", self.conn
            )
            if df.empty:
                return {}
            bins   = [0, 50, 60, 70, 80, 90, 101]
            labels = ['Below 50', '50-59', '60-69', '70-79', '80-89', '90-100']
            df['range'] = pd.cut(df['percentage'], bins=bins, labels=labels, right=False)
            counts = df['range'].value_counts().reindex(labels, fill_value=0)
            return counts.to_dict()
        except Exception as e:
            return {}

    def get_subject_pass_fail(self) -> pd.DataFrame:
        """
        Har subject mein kitne PASS aur FAIL hue — comparison ke liye.
        Returns DataFrame: subject_index | subject_name | pass_count | fail_count | pass_perc
        """
        try:
            return pd.read_sql_query("""
                SELECT
                    subject_index,
                    subject_name,
                    SUM(CASE WHEN status='PASS' THEN 1 ELSE 0 END) as pass_count,
                    SUM(CASE WHEN status='FAIL' THEN 1 ELSE 0 END) as fail_count,
                    COUNT(*) as total,
                    ROUND(SUM(CASE WHEN status='PASS' THEN 1.0 ELSE 0 END)/COUNT(*)*100,1)
                        as pass_percentage
                FROM subject_marks
                GROUP BY subject_index
                ORDER BY subject_index
            """, self.conn)
        except Exception as e:
            return pd.DataFrame()

    # ------------------------------------------------------------------
    # UTILITY
    # ------------------------------------------------------------------

    def get_session_info(self) -> dict:
        """Current import session ki info return karo."""
        try:
            if self._session_id:
                self.cursor.execute(
                    "SELECT * FROM import_sessions WHERE id=?",
                    (self._session_id,)
                )
            else:
                self.cursor.execute(
                    "SELECT * FROM import_sessions ORDER BY id DESC LIMIT 1"
                )
            row = self.cursor.fetchone()
            return dict(row) if row else {}
        except Exception as e:
            return {}

    def get_college_list(self) -> list:
        """Saare unique college names."""
        try:
            self.cursor.execute(
                "SELECT DISTINCT college FROM students WHERE college!='' ORDER BY college ASC"
            )
            return [r['college'] for r in self.cursor.fetchall()]
        except Exception as e:
            return []

    def get_subject_names(self) -> dict:
        """
        Subject index → name mapping (from latest import).
        Returns: {0: "Mathematics", 1: "Physics", ...}
        """
        try:
            self.cursor.execute("""
                SELECT DISTINCT subject_index, subject_name
                FROM subject_marks
                ORDER BY subject_index
            """)
            return {r['subject_index']: r['subject_name']
                    for r in self.cursor.fetchall()}
        except Exception as e:
            return {}

    def get_total_count(self) -> int:
        try:
            self.cursor.execute("SELECT COUNT(*) FROM students")
            return self.cursor.fetchone()[0]
        except Exception:
            return 0

    def is_empty(self) -> bool:
        return self.get_total_count() == 0

    def delete_all(self) -> bool:
        try:
            self._clear_all_data()
            return True
        except Exception as e:
            print(f"[DB] delete_all Error: {e}")
            return False

    def delete_student(self, seat_no: str) -> bool:
        try:
            self.cursor.execute(
                "DELETE FROM students WHERE seat_no=?", (seat_no,)
            )
            self.conn.commit()
            return self.cursor.rowcount > 0
        except Exception as e:
            return False

    def get_summary_text(self) -> str:
        stats = self.get_stats()
        if not stats:
            return "No data loaded."
        session = self.get_session_info()
        course_info = ""
        if session:
            c = session.get("course", "")
            sem = session.get("semester", "")
            ay  = session.get("academic_year", "")
            if c or sem:
                course_info = f" | {c} {sem}"
            if ay:
                course_info += f" ({ay})"
        return (
            f"Total: {stats['total_students']} | "
            f"Pass: {stats['pass_count']} ({stats['pass_perc']}%) | "
            f"Fail: {stats['fail_count']} | "
            f"Avg: {stats['avg_percentage']}%"
            f"{course_info}"
        )


# ==============================================================================
# Module-level convenience
# ==============================================================================

def create_db(db_path: str = ":memory:") -> DatabaseManager:
    """DatabaseManager ka instance banao."""
    return DatabaseManager(db_path)


# ==============================================================================
# Self-Test
# ==============================================================================
if __name__ == "__main__":
    print("=" * 60)
    print("  DatabaseManager v3 — Self Test")
    print("=" * 60)

    db = DatabaseManager()

    # Simulate what pdf_parser gives us
    subjects_map = {
        0: "Fundamentals of Programming (C)",
        1: "Mathematics",
        2: "Communication Skills",
        3: "Computer Organization",
        4: "Practical - I",
        5: "Practical - II",
        6: "VAC",
    }
    ext_passing_min = [18, 18, 18, 18, 9, 9, 9]

    sid = db.start_import_session(
        pdf_filename  = "BSc_IT_Sem1_Nov2025.pdf",
        course        = "BSc IT",
        semester      = "Sem 1",
        academic_year = "2025-2026",
        college_name  = "SASCMA",
        max_marks     = 700,
    )
    print(f"\nSession ID: {sid}")

    students = [
        {
            "seat_no": "001", "sp_id": "1234567890", "name": "ARJUN SHARMA",
            "gender": "M", "college": "SASCMA",
            "marks": {0:["25","10"], 1:["30","12"], 2:["22","9"],
                      3:["28","11"], 4:["15","0"], 5:["12","0"], 6:["8","0"]},
            "grades": {0:"A+", 1:"O", 2:"A", 3:"A+", 4:"A", 5:"A", 6:"O"},
            "total_marks": 182, "sgpa": "8.50", "overall_status": "PASS",
        },
        {
            "seat_no": "002", "sp_id": "9876543210", "name": "PRIYA PATEL",
            "gender": "F", "college": "SASCMA",
            "marks": {0:["15","7"], 1:["12","5"], 2:["18","8"],
                      3:["20","9"], 4:["9","0"], 5:["7","0"], 6:["6","0"]},
            "grades": {0:"F", 1:"F", 2:"C", 3:"B", 4:"C", 5:"F", 6:"C"},
            "total_marks": 116, "sgpa": "0.00", "overall_status": "FAIL",
        },
        {
            "seat_no": "003", "sp_id": "5555555555", "name": "RAHUL MEHTA",
            "gender": "M", "college": "VNSGU",
            "marks": {0:["20","8"], 1:["22","9"], 2:["19","7"],
                      3:["21","9"], 4:["11","0"], 5:["10","0"], 6:["7","0"]},
            "grades": {0:"B", 1:"A", 2:"B", 3:"B+", 4:"B", 5:"B", 6:"B"},
            "total_marks": 143, "sgpa": "6.80", "overall_status": "PASS",
        },
    ]

    result = db.insert_students(students, sid, subjects_map, ext_passing_min, max_marks=700)
    print(f"\nInsert result : {result}")
    print(f"Total students: {db.get_total_count()}")

    print("\n--- get_stats() ---")
    stats = db.get_stats()
    for k, v in stats.items():
        print(f"  {k:20s}: {v}")

    print("\n--- get_all_data() columns ---")
    df = db.get_all_data()
    print("  Columns:", list(df.columns))
    print("\n--- Flat view (roll_no, name, sub_1..7, total, sgpa, status) ---")
    print(df[['roll_no','name','sub_1','sub_2','sub_3','sub_4','sub_5','sub_6','sub_7',
              'total','sgpa','status','grade','atkt_count']].to_string(index=False))

    print("\n--- EXT / INT per subject for student 001 ---")
    print(df[df['roll_no']=='001'][['sub_1_ext','sub_1_int',
                                    'sub_2_ext','sub_2_int',
                                    'sub_1_grade','sub_2_grade']].to_string(index=False))

    print("\n--- Subject Averages ---")
    for key, val in db.get_subject_averages().items():
        print(f"  {key}: {val}")

    print("\n--- Subject Pass/Fail ---")
    print(db.get_subject_pass_fail().to_string(index=False))

    print("\n--- Grade Distribution ---")
    print(db.get_grade_distribution())

    print("\n--- College Stats ---")
    print(db.get_college_stats().to_string(index=False))

    print("\n--- ATKT Students ---")
    atkt = db.get_atkt_students()
    if not atkt.empty:
        print(atkt[['roll_no','name','atkt_count','failed_subjects']].to_string(index=False))

    print("\n--- Summary Text ---")
    print(db.get_summary_text())

    print("\n--- Session Info ---")
    for k, v in db.get_session_info().items():
        print(f"  {k}: {v}")

    db.close()
    print("\n" + "=" * 60)
    print("  Self-Test Complete [OK]")
    print("=" * 60)
