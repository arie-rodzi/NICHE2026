
import streamlit as st
st.set_page_config(page_title="NICHE 2026", layout="wide")

st.markdown("<h1 style='text-align:center'>NICHE 2026 Conference Management System</h1>", unsafe_allow_html=True)

menu = ["Home","Register","Academic Conference","Industrial Conference","Gala Dinner","Abstract Book","My Registration","Admin"]
tabs = st.tabs(menu)

with tabs[0]:
    st.header("Home")
    st.write("Upload posters via Admin panel.")

with tabs[1]:
    st.header("Register")
    email = st.text_input("Email")
    st.button("Continue")

with tabs[2]:
    st.header("Academic Conference")

with tabs[3]:
    st.header("Industrial Conference")

with tabs[4]:
    st.header("Gala Dinner")

with tabs[5]:
    st.header("Abstract Book")

with tabs[6]:
    st.header("My Registration")

with tabs[7]:
    st.header("Admin")
    pwd = st.text_input("Admin Password", type="password")
    if pwd == "NICHE2026admin":
        st.success("Admin Access Granted")
        st.file_uploader("Upload NICHE2026_MASTER.xlsx")
        st.file_uploader("Upload Conference Poster")
        st.file_uploader("Upload Dinner Poster")
