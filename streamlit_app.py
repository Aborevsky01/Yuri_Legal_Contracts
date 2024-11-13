import streamlit as st
from datetime import datetime
from src.jura import Jura

jura = Jura()

st.set_page_config(
    page_title="Юрист AI-Powered",
    page_icon="⚖️",
    layout="centered",
    initial_sidebar_state="expanded",
)

st.title('⚖️ Юрист AI-Powered')

if 'generated' not in st.session_state:
    st.session_state['generated'] = False
    st.session_state['filename'] = ''
    st.session_state['question'] = ''

question = st.text_area(
    "Напиши, какой договор тебе необходимо составить",
    placeholder=st.session_state["question"] if st.session_state["question"] != '' else "Хочу продать дом компании с ИНН 000...",
    height=200
)

if question and not st.session_state['generated']:
    time_now = str(datetime.now())
    st.session_state['filename'] = time_now + 'doc.docx'
    with st.spinner('Генерация...'):
        jura.launch(question, time_now)
    st.session_state['generated'] = True

if st.session_state['generated']:
    with open(st.session_state['filename'], "rb") as file:
        btn = st.download_button(
            label="Скачай сгенерированный файл",
            data=file,
            file_name=st.session_state['filename'],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

with st.sidebar:
    "Sidebar"
