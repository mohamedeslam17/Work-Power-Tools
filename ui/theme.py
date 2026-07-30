"""Minimal presentation shell. Colours live in .streamlit/config.toml — this
is layout-only CSS, kept deliberately small (see the redesign brief's 40-line
CSS budget)."""
import streamlit as st

_CSS = """
<style>
.block-container { padding-top: 2rem; max-width: 1100px; }
[data-testid="stDecoration"] { display: none; }
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] { padding: 0 0 .5rem .1rem; }
</style>
"""


def inject():
    st.markdown(_CSS, unsafe_allow_html=True)
