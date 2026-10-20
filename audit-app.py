"""Streamlit Community Cloud entry point: a standalone "moved" notice.

The application now runs on Render with individual accounts (`app.py` is the router there). This
file deliberately imports nothing from the app and creates no API client, so the public Community
Cloud copy serves no content and records nothing. Delete it with the Community Cloud app.
"""
import streamlit as st

NEW_URL = "https://aarp-automatic-audit.onrender.com/"

st.set_page_config(page_title="Automatic Audit has moved")
st.title("Automatic Audit has moved")
st.markdown(
    f"The tool now runs at **[{NEW_URL}]({NEW_URL})** and requires an individual "
    "account. Use the username and password you already use for the Model Rules Explorer "
    "or Themes Explorer; if you do not have one, contact J.T. Fales at Austin Data Solutions."
)
st.info("This page will be taken down 30 days after the move. Please update your bookmark.")
st.caption("Nothing is uploaded, audited, or recorded on this page.")
