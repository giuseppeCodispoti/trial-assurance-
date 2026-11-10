import streamlit as st

st.set_page_config(page_title="Programmazione e Consuntivo Trial", layout="wide")

st.title("📊 PROGRAMMAZIONE ASSURANCE – TRIAL")
scelta = st.selectbox(
    "Scegli l'app da eseguire:",
    ["programmazione trial", "consuntivo trial"]
)

if scelta == "programmazione trial":
    import trial
    trial.run()
else:
    import trialChiusi
    trialChiusi.run()
