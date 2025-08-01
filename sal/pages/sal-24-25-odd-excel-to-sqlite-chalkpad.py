import streamlit as st
import pandas as pd
import sqlite3

DB_NAME = "sal.db"
TABLE_NAME = "salTable"

def create_table():
    with sqlite3.connect(DB_NAME) as conn:
        conn.execute(f'''
            CREATE TABLE IF NOT EXISTS "{TABLE_NAME}" (
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
                "st1-mm" TEXT NOT NULL,
                "st1-mo" TEXT NOT NULL,
                "st2-mm" TEXT NOT NULL,
                "st2-mo" TEXT NOT NULL,
                "ete-grade" TEXT,
                UNIQUE(roll, "subject-code", semester)
            )
        ''')

def parse_mark(val):
    if pd.isna(val): return "A"
    val = str(val).strip().lower()
    if val in ["a", "ab", "absent", "-", "n/a", "na"]:
        return "A"
    return val if val.replace('.', '', 1).isdigit() else "A"

def is_duplicate(roll, subject_code, semester):
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute(f'''
            SELECT 1 FROM "{TABLE_NAME}"
            WHERE roll = ? AND "subject-code" = ? AND semester = ?
        ''', (roll, subject_code, semester))
        return cur.fetchone() is not None

def insert_data(df, meta, columns):
    inserted, skipped = 0, 0
    with sqlite3.connect(DB_NAME) as conn:
        for _, row in df.iterrows():
            try:
                roll = int(row[columns["roll"]])
                name = str(row[columns["name"]])
                st1_mm = parse_mark(row[columns["st1-mm"]])
                st1_mo = parse_mark(row[columns["st1-mo"]])
                st2_mm = parse_mark(row[columns["st2-mm"]])
                st2_mo = parse_mark(row[columns["st2-mo"]])
            except Exception as e:
                st.warning(f"Skipping row due to error: {e}")
                continue

            if is_duplicate(roll, meta["subject_code"], meta["semester"]):
                skipped += 1
                continue

            conn.execute(f'''
                INSERT INTO "{TABLE_NAME}" (
                    batch, "academic-year", "odd-even", session, branch, semester,
                    roll, name, "subject-code", "subject-name",
                    "st1-mm", "st1-mo", "st2-mm", "st2-mo", "ete-grade"
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                meta["batch"], meta["academic_year"], meta["odd_even"],
                meta["session"], meta["branch"], meta["semester"],
                roll, name, meta["subject_code"], meta["subject_name"],
                st1_mm, st1_mo, st2_mm, st2_mo, None
            ))
            inserted += 1

    st.success(f"✅ Inserted: {inserted}")
    if skipped:
        st.info(f"⏩ Skipped duplicates: {skipped}")

def main():
    st.set_page_config(page_title="C2 Upload", layout="wide")
    st.title("Excel → SQLite")
    st.header("Works for only 24-25 Odd Semester Chalkpad sheets that are modified")

    create_table()
    file = st.file_uploader("Upload Excel (.xlsx)", type="xlsx")
    if not file:
        return

    try:
        meta = pd.read_excel(file, header=None, nrows=4)
        class_name = str(meta.iloc[0, 1])
        subject_name = str(meta.iloc[1, 1])
        subject_code = str(meta.iloc[2, 1])
        parts = class_name.split('-')
        batch = int(parts[0])
        branch = parts[2]
        semester = int(parts[3].split()[0])

        # Read with single header only (row 5)
        df = pd.read_excel(file, header=4)

        st.subheader("📊 Detected Columns")
        st.write(df.columns.tolist())

        # Map columns by position
        columns = {
            "roll": "Roll. No",
            "name": "Student Name",
            "st1-mo": df.columns[4],
            "st1-mm": df.columns[5],
            "st2-mo": df.columns[6],
            "st2-mm": df.columns[7],
        }

        missing = [k for k, v in columns.items() if v not in df.columns]
        if missing:
            st.error(f"❌ Missing columns in Excel: {missing}")
            return

    except Exception as e:
        st.error(f"❌ Failed to read Excel: {e}")
        return

    with st.sidebar:
        st.header("🛠 Input")
        academic_year = st.text_input("Academic Year", "24-25")
        odd_even = st.selectbox("Odd/Even", ["Odd", "Even"])
        session = st.text_input("Session", "July 2024 - December 2024")
        batch = st.number_input("Batch", 2000, 2100, value=batch)
        branch = st.text_input("Branch", value=branch)
        semester = st.number_input("Semester", 1, 8, value=semester)
        st.text(f"📘 Subject: {subject_code} - {subject_name}")
        if st.button("🚀 Insert into Database"):
            insert_data(df, {
                "academic_year": academic_year,
                "odd_even": odd_even,
                "session": session,
                "batch": batch,
                "branch": branch,
                "semester": semester,
                "subject_code": subject_code,
                "subject_name": subject_name,
            }, columns)

    if st.checkbox("📄 Show salTable"):
        with sqlite3.connect(DB_NAME) as conn:
            st.dataframe(pd.read_sql_query(f'SELECT * FROM "{TABLE_NAME}"', conn))

if __name__ == "__main__":
    main()
