"""Shared presentation shell for the materials tools."""
import streamlit as st

_CSS = """
<style>
.block-container { padding-top: 2rem; max-width: 1100px; }
[data-testid="stDecoration"] { display: none; }
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] { padding: 0 0 .5rem .1rem; }
.aeg-page-kicker {
  color: #687386; font-size: .72rem; font-weight: 700; letter-spacing: .12em;
  margin: 0 0 .45rem .15rem;
}
.aeg-issue-card, .aeg-clear-card {
  border: 1px solid #d9dee7; border-left-width: 4px; border-radius: 10px;
  background: #fff; padding: .78rem .82rem; margin: 0 0 .7rem;
  box-shadow: 0 1px 2px rgba(21, 31, 48, .05);
}
.aeg-critical { border-left-color: #d62d38; background: #fff8f8; }
.aeg-warning { border-left-color: #ec8612; background: #fffaf3; }
.aeg-info { border-left-color: #186ed6; background: #f7faff; }
.aeg-pass { border-left-color: #20a050; background: #f6fcf8; }
.aeg-issue-head { display: flex; align-items: center; gap: .45rem; margin-bottom: .35rem; }
.aeg-issue-number {
  display: inline-flex; width: 1.45rem; height: 1.45rem; border-radius: 999px;
  align-items: center; justify-content: center; background: #253047; color: #fff;
  font-size: .72rem; font-weight: 800;
}
.aeg-issue-label { color: #283348; font-size: .76rem; font-weight: 800; text-transform: uppercase; }
.aeg-issue-meta { color: #687386; font-size: .72rem; font-weight: 650; margin-bottom: .32rem; }
.aeg-issue-copy { color: #30394a; font-size: .86rem; line-height: 1.42; }
.aeg-clear-card { border-left-color: #35a864; background: #f6fcf8; }
.aeg-clear-title { color: #207a43; font-weight: 750; margin-bottom: .2rem; }
.aeg-clear-copy { color: #5f6b7b; font-size: .84rem; }
</style>
"""


def inject():
    st.markdown(_CSS, unsafe_allow_html=True)
