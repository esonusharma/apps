import streamlit as st
import pandas as pd

def check_login():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if not st.session_state["authenticated"]:
        with st.form("Login"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submit = st.form_submit_button("Login")

            if submit:
                valid_users = st.secrets["credentials"]
                if username in valid_users and password == valid_users[username]:
                    st.session_state["authenticated"] = True
                    st.session_state["user"] = username
                    st.success(f"Welcome, {username}!")
                else:
                    st.error("Invalid username or password")

    return st.session_state["authenticated"]

if check_login():
    st.title("Trade Analyzer")

    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            st.write("File Preview:")
            st.dataframe(df)

            column = st.selectbox("Pick a column", df.columns)

            if column:
                numeric_values = pd.to_numeric(df[column], errors="coerce").dropna()
                filtered_values = numeric_values.nlargest(3)
                st.write(f"Filtered Values '{column}':")
                st.write(filtered_values)
        except Exception as e:
            st.error(f"Error processing file: {e}")