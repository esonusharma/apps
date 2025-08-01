import streamlit as st
import pandas as pd
import sqlite3
import re

DB_NAME = "sal.db"
TABLE_NAME = "salTable"

# ---------------------- DB Setup ----------------------
def create_table_if_not_exists():
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute(f'''
        CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
            batch INTEGER NOT NULL,
            "academic-year" TEXT NOT NULL,
            "odd-even" TEXT NOT NULL,
            session TEXT NOT NULL,
            branch TEXT NOT NULL,
            semester INTEGER NOT NULL,
            roll INTEGER NOT NULL,
            name TEXT NOT NULL,
            "subject-code" TEXT NOT NULL,
            "subject-name" TEXT NOT NULL,
            "st1-mm" NUMERIC(10) NOT NULL,
            "st1-mo" TEXT NOT NULL,
            "st2-mm" NUMERIC(10) NOT NULL,
            "st2-mo" TEXT NOT NULL,
            "ete-grade" TEXT,
            UNIQUE(roll, "subject-code", semester)
        )
    ''')
    conn.commit()
    conn.close()

# ---------------------- Utilities ----------------------

def clean_session(session_str):
    if pd.isna(session_str):
        return ""
    session_str = str(session_str).strip()
    match = re.match(r"([A-Za-z]+)-([A-Za-z]+)\s+(\d{4})", session_str)
    if match:
        start_month, end_month, year = match.groups()
        return f"{start_month} {year} - {end_month} {year}"
    if " - " in session_str:
        return session_str
    return session_str

def parse_marks(value):
    if pd.isna(value):
        return "A"
    val = str(value).strip().lower()
    if val in ['a', 'ab', 'absent', '-', 'na', 'n/a']:
        return "A"
    if val.isdigit():
        return int(val)
    return "A"

# ---------------------- Insert Function with Duplicate Check ----------------------

def insert_to_db(df, academic_year, odd_even):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    inserted_count = 0
    skipped_count = 0

    for _, row in df.iterrows():
        try:
            roll = int(row['Roll No.'])
            subject_code = row['Subject Code']
            semester = int(row['Semester'])

            # Check if record already exists
            cur.execute(f'''
                SELECT 1 FROM {TABLE_NAME}
                WHERE roll = ? AND "subject-code" = ? AND semester = ?
            ''', (roll, subject_code, semester))
            if cur.fetchone():
                skipped_count += 1
                continue  # Skip existing row

            cur.execute(f'''
                INSERT INTO {TABLE_NAME} (
                    batch, "academic-year", "odd-even", session, branch, semester,
                    roll, name, "subject-code", "subject-name",
                    "st1-mm", "st1-mo", "st2-mm", "st2-mo", "ete-grade"
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                int(row['Batch']),
                academic_year,
                odd_even,
                clean_session(row['Session']),
                row['Branch'],
                semester,
                roll,
                row['Student Name'],
                subject_code,
                row['Subject Name'],
                parse_marks(row['ST1 Total Marks']),
                str(parse_marks(row['ST1 Marks'])),
                parse_marks(row['ST2 Total Marks']),
                str(parse_marks(row['ST2 Marks'])),
                None
            ))
            inserted_count += 1
        except Exception as e:
            st.warning(f"⚠️ Error inserting row: {e}")

    conn.commit()
    conn.close()
    st.success(f"✅ {inserted_count} new rows inserted.")
    if skipped_count > 0:
        st.info(f"ℹ️ {skipped_count} duplicate rows skipped.")

# ---------------------- Streamlit App ----------------------

def main():
    st.set_page_config(page_title="SAL DB Loader", layout="wide")
    st.title("Excel → SQLite")
    st.header("Works for only 24-25 Odd Semester Faculty filled sheets")

    create_table_if_not_exists()

    # Sidebar UI
    with st.sidebar:
        st.header("Input")
        academic_year = st.text_input("Academic Year", value="24-25")
        odd_even = st.selectbox("Term", ["Odd", "Even"])
        uploaded_files = st.file_uploader("Upload Excel Files", type=['xlsx'], accept_multiple_files=True)
        process_triggered = st.button("🚀 Process Files")

    if process_triggered and uploaded_files:
        for file in uploaded_files:
            try:
                df = pd.read_excel(file)

                required_cols = [
                    'Batch', 'Session', 'Branch', 'Semester', 'Roll No.', 'Student Name',
                    'Subject Code', 'Subject Name', 'ST1 Total Marks', 'ST1 Marks',
                    'ST2 Total Marks', 'ST2 Marks'
                ]

                if not all(col in df.columns for col in required_cols):
                    st.error(f"❌ Missing required columns in {file.name}")
                    continue

                df_clean = df[required_cols].copy()
                insert_to_db(df_clean, academic_year, odd_even)
                st.success(f"✅ {file.name} processed.")
            except Exception as e:
                st.error(f"❌ Error processing {file.name}: {e}")

    if st.checkbox("📄 Show Inserted Records"):
        conn = sqlite3.connect(DB_NAME)
        df_all = pd.read_sql_query(f"SELECT * FROM {TABLE_NAME}", conn)
        st.dataframe(df_all)
        conn.close()

if __name__ == "__main__":
    main()
