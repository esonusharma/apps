import streamlit as st
import pandas as pd
import sqlite3
import numpy as np
import os

# Streamlit sidebar inputs
st.sidebar.title("Input")

excel_file = st.sidebar.file_uploader("Upload Excel file", type=["xlsx"])

academic_year = st.sidebar.text_input("Academic Year (e.g., 24-25)")
odd_even = st.sidebar.selectbox("Odd/Even Semester", ["Odd", "Even"])

process_btn = st.sidebar.button("🚀 Start Processing")

st.title("Excel → SQLite")
st.header("Works for only 24-25 Even Semester Faculty filled sheets all inside single Google Sheet downloaded as Excel Sheet")

DB_FILE = "sal.db"
TABLE_NAME = "salTable"

# Required columns mapping
COLUMN_MAP = {
    "Batch": "batch",
    "Session": "session",
    "Branch": "branch",
    "Semester": "semester",
    "Roll": "roll",
    "Student Name": "name",
    "Subject Code": "subject-code",
    "Subject Name": "subject-name",
    "ST1 Maximum Marks": "st1-mm",
    "ST1 Obtained Marks": "st1-mo",
    "ST2 Maximum Marks": "st2-mm",
    "ST2 Obtained Marks": "st2-mo",
}

REQUIRED_COLUMNS = list(COLUMN_MAP.values()) + ["academic-year", "odd-even", "ete-grade"]


def create_table_if_not_exists():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(f"""
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
        "st1-mo" NUMERIC(10) NOT NULL,
        "st2-mm" NUMERIC(10) NOT NULL,
        "st2-mo" NUMERIC(10) NOT NULL,
        "ete-grade" TEXT NOT NULL
    )
    """)
    conn.commit()
    conn.close()


def clean_and_map(df):
    # Normalize absent values to "A"
    df.replace(
        to_replace=r"(?i)\b(Ab|Absent|A|ab)\b",
        value="A",
        regex=True,
        inplace=True
    )

    # Rename columns
    df.rename(columns=COLUMN_MAP, inplace=True)

    # Add extra columns
    df["academic-year"] = academic_year
    df["odd-even"] = odd_even
    df["ete-grade"] = np.nan

    # Drop rows with any missing required fields
    df.dropna(subset=[
        "batch", "session", "branch", "semester", "roll", "name",
        "subject-code", "subject-name", "st1-mm", "st1-mo", "st2-mm", "st2-mo"
    ], inplace=True)

    # Ensure correct data types
    df["batch"] = df["batch"].astype(int)
    df["semester"] = df["semester"].astype(int)
    df["roll"] = df["roll"].astype(int)

    # Final clean dataframe
    return df[[*COLUMN_MAP.values(), "academic-year", "odd-even", "ete-grade"]]


def row_exists(conn, row):
    cursor = conn.cursor()
    query = f"""
        SELECT 1 FROM {TABLE_NAME}
        WHERE roll=? AND "subject-code"=? AND "academic-year"=? AND "odd-even"=?
    """
    cursor.execute(query, (row["roll"], row["subject-code"], row["academic-year"], row["odd-even"]))
    return cursor.fetchone() is not None


def insert_into_db(df):
    conn = sqlite3.connect(DB_FILE)
    inserted = 0
    skipped = 0
    for _, row in df.iterrows():
        if not row_exists(conn, row):
            row.to_frame().T.to_sql(TABLE_NAME, conn, if_exists="append", index=False)
            inserted += 1
        else:
            skipped += 1
    conn.close()
    return inserted, skipped


if process_btn:
    if not excel_file:
        st.error("⚠️ Please upload an Excel file first.")
    elif not academic_year or not odd_even:
        st.error("⚠️ Please enter academic year and select odd/even.")
    else:
        create_table_if_not_exists()

        xls = pd.ExcelFile(excel_file, engine='openpyxl')
        sheet_names = xls.sheet_names[1:]  # Skip the first sheet

        total_inserted = 0
        total_skipped = 0
        total_sheets = 0

        for sheet in sheet_names:
            df = xls.parse(sheet)
            if df.empty or "Roll" not in df.columns or df["Roll"].dropna().empty:
                continue  # Skip if no students

            try:
                df_clean = clean_and_map(df)
                inserted, skipped = insert_into_db(df_clean)
                total_inserted += inserted
                total_skipped += skipped
                total_sheets += 1
            except Exception as e:
                st.warning(f"⚠️ Skipped sheet '{sheet}': {e}")

        st.success(f"✅ Processed {total_sheets} sheets.")
        st.info(f"✅ Inserted: {total_inserted} rows. Skipped (duplicates): {total_skipped}.")
