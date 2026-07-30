import streamlit as st

from ui import theme
from ui import lab_tool, iir_tool, sem_tool, photo_tool

TOOLS = [
    ("🧪", "Lab Report Review",
     "Is this lab report wrong anywhere, and where exactly?", lab_tool.render),
    ("🛠️", "IIR Review",
     "Is this incoming-inspection report wrong anywhere, and where exactly?", iir_tool.render),
    ("🔬", "SEM Converter",
     "Turn a vendor SEM PDF into our Word report format.", sem_tool.render),
    ("🖼️", "Photo Library",
     "Browse stored micrographs by alloy.", photo_tool.render),
]


def main():
    st.set_page_config(page_title="AEG Materials Tools", page_icon="🔬",
                       layout="wide", initial_sidebar_state="expanded")
    theme.inject()

    with st.sidebar:
        st.caption("AEG Materials Tools")
        labels = [f"{icon}  {name}" for icon, name, _, _ in TOOLS]
        choice = st.radio("Tool", labels, label_visibility="collapsed")
    icon, name, purpose, run = TOOLS[labels.index(choice)]

    st.title(name)
    st.caption(purpose)
    try:
        run()
    except Exception as e:
        st.error("This tool hit an error and couldn't finish.")
        with st.expander("Error details"):
            st.exception(e)


if __name__ == "__main__":
    main()
