import streamlit as st

scelta = st.selectbox("Scegli l'app da eseguire:", ["programmazione trial ", "consuntivo trial"])
if scelta == "programmazione trial ":
    import trial
    trial.run()
elif scelta == "consuntivo trial":
    import trialChiusi
    trialChiusi.run()
