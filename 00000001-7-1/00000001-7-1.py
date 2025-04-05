import streamlit as st
import pandas as pd

USERNAME = "admin"
PASSWORD = "r2f"

def login():
    st.title("Login")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    login_button = st.button("Login")

    if login_button:
        if username == USERNAME and password == PASSWORD:
            st.session_state["logged_in"] = True
        else:
            st.error("Invalid username or password")

def main_app():
    st.title("3 Worst Drawdowns")

    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)

            st.write("Preview of the uploaded data:")
            st.dataframe(df)

            column = st.selectbox("Select a column to analyze", df.columns)

            if column:
                numeric_col = pd.to_numeric(df[column], errors='coerce')
                top_values = numeric_col.dropna().nlargest(3)

                st.write(f"3 Worst Drawdowwns '{column}':")
                st.write(top_values)
        except Exception as e:
            st.error(f"Error reading the file: {e}")

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if st.session_state["logged_in"]:
    main_app()
else:
    login()