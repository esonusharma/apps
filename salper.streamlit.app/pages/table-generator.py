import streamlit as st
import numpy as np

def table_generator():
    """
    This function generates tables.
    """
    with st.sidebar:
        st.title("Table Generator")
        # Get user input
        st.subheader("Enter the following parameters:")
    # Display results
    st.subheader("Results:")

# Main app
st.set_page_config(layout="wide")
table_generator()
st.text('This code/service/product is subject to the terms of the MIT License')