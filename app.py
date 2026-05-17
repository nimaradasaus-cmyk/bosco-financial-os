import streamlit as st
import pandas as pd
import anthropic
import sqlite3
import json
import os
from dotenv import load_dotenv
from datetime import date
load_dotenv()
client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
st.title("Bosco Financial OS")

# --- UPLOAD FICHIER EXCEL ---
fichier = st.file_uploader("📂 Charge ton fichier Excel (DATA - EXCEL .xlsx)", type=["xlsx"])

if fichier is not None:
    # --- SNAPSHOT ---
    st.header("Snapshot")
    df = pd.read_excel(fichier, sheet_name="Money Flow", header=552, nrows=20)
    df = df.dropna(how="all")
    df_reel = df[df["Income"] > 0].copy()
    derniere = df_reel.iloc[-1]
    cash_actuel = 4015
    objectif_sydney = 8000
    semaines_restantes = 5
    besoin_semaine = round((objectif_sydney - cash_actuel) / semaines_restantes, 2)
    col1, col2, col3 = st.columns(3)
    col1.metric("Cash Sydney", f"{cash_actuel} AUD", f"Objectif {objectif_sydney}")
    col2.metric("Surplus dernière semaine", f"{round(derniere['True Surplus W/O Savings'], 2)} AUD")
    col3.metric("À mettre de côté", f"{besoin_semaine} AUD/sem", f"{semaines_restantes} semaines")

    # --- AGENT ANALYSTE ---
    st.header("Agent Analyste")
    if st.button("Analyser ma situation"):
        with st.spinner("Analyse en cours..."):
            snapshot = f"""
Cash Sydney : {cash_actuel} AUD
Objectif : {objectif_sydney} AUD
Semaines restantes : {semaines_restantes}
Besoin/semaine : {besoin_semaine} AUD
Revenu dernière semaine : {round(derniere['Income'], 2)} AUD
Surplus W/O Savings : {round(derniere['True Surplus W/O Savings'], 2)} AUD
"""
            response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=1024,
                system="Tu es l'analyste financier personnel de Nicolas, 25 ans, casual worker à Melbourne, déménagement Sydney imminent. Direct, structuré, prudent.",
                messages=[{"role": "user", "content": f"Snapshot :\n{snapshot}\nDiagnostic et recommandation concrète."}]
            )
            st.markdown(response.content[0].text)

else:
    st.info("⬆️ Charge ton fichier Excel pour accéder au Snapshot, à l'Analyste et au Tracker Sydney.")

# --- DECISION JOURNAL ---
st.header("Decision Journal")
nouvelle = st.text_area("Décris une décision en langage naturel")
if st.button("Archiver cette décision"):
    with st.spinner("Archivage..."):
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=500,
            system="""Transforme en JSON avec 4 champs : decision, contexte, raisonnement, invalidation. JSON uniquement, rien d'autre.""",
            messages=[{"role": "user", "content": nouvelle}]
        )
        raw = response.content[0].text.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        data = json.loads(raw.strip())
        conn = sqlite3.connect("bosco_financial_os.db")
        cursor = conn.cursor()
        cursor.execute("""
CREATE TABLE IF NOT EXISTS decisions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT, decision TEXT, contexte TEXT, raisonnement TEXT, invalidation TEXT
)""")
        cursor.execute("INSERT INTO decisions VALUES (?, ?, ?, ?, ?, ?)",
            (None, str(date.today()), data["decision"], data["contexte"], data["raisonnement"], data["invalidation"]))
        conn.commit()
        conn.close()
        st.success("Décision archivée.")
        st.markdown(f"**Décision :** {data['decision']}")
        st.markdown(f"**Invalidation :** {data['invalidation']}")

# --- HISTORIQUE ---
st.header("Historique des décisions")
conn = sqlite3.connect("bosco_financial_os.db")
cursor = conn.cursor()
cursor.execute("""
CREATE TABLE IF NOT EXISTS decisions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT, decision TEXT, contexte TEXT, raisonnement TEXT, invalidation TEXT
)""")
conn.commit()
decisions = pd.read_sql("SELECT date, decision, invalidation FROM decisions ORDER BY id DESC", conn)
conn.close()
st.dataframe(decisions)

# --- AGENT CHALLENGER ---
st.header("Agent Challenger")
decision_a_challenger = st.text_area("Quelle décision veux-tu que je challenge ?", key="challenger")
if st.button("Challenger cette décision"):
    with st.spinner("Recherche des failles..."):
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1024,
            system="""Tu es le Challenger financier de Nicolas. Ton rôle est de trouver les failles dans son raisonnement. Tu ne valides pas, tu questionnes. Tu cherches : les biais émotionnels, les rationalisations, les informations manquantes, les risques ignorés. Tu es direct et sans complaisance. Nicolas a 25 ans, casual worker, déménagement Sydney imminent, cash limité.""",
            messages=[{"role": "user", "content": f"Voici ma décision : {decision_a_challenger}\n\nChallenge-moi."}]
        )
        st.markdown(response.content[0].text)

# --- TRACKER SYDNEY ---
st.header("Tracker Sydney")
if fichier is not None:
    hq = pd.read_excel(fichier, sheet_name="Headquarters", header=48, nrows=24)
    hq = hq.dropna(how="all")
    hq = hq[hq["Unnamed: 1"].astype(str).str.contains(r"\d+/\d+/\d+", na=False)]
    hq_clean = hq[["Unnamed: 1", "Planned Expenses Acc"]].copy()
    hq_clean.columns = ["Date", "Cash Sydney"]
    hq_clean = hq_clean[hq_clean["Cash Sydney"] > 0]
    hq_clean["Date"] = hq_clean["Date"].astype(str).str.extract(r"(\d+/\d+/\d+)")
    hq_clean["Date_tri"] = pd.to_datetime(hq_clean["Date"], dayfirst=True)
    hq_clean = hq_clean.sort_values("Date_tri")
    cash_actuel_sydney = hq_clean["Cash Sydney"].iloc[-1]
    objectif = 8000
    semaines_restantes = 5
    besoin = round((objectif - cash_actuel_sydney) / semaines_restantes, 2)
    st.line_chart(hq_clean.set_index("Date_tri")["Cash Sydney"])
    col1, col2, col3 = st.columns(3)
    col1.metric("Cash Sydney actuel", f"${round(cash_actuel_sydney, 2)}")
    col2.metric("Objectif", f"${objectif}")
    col3.metric("À mettre de côté", f"${besoin}/sem", f"{semaines_restantes} semaines")
else:
    st.warning("Charge d'abord ton fichier Excel.")

# --- AGENT FORMATEUR ---
st.header("Agent Formateur")
question_fiscale = st.text_input("Pose une question sur la fiscalité australienne, ta super, ton visa...")
if st.button("Poser la question"):
    with st.spinner("Recherche en cours..."):
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1024,
            system="""Tu es le conseiller financier et fiscal de Nicolas. Il est français, 25 ans, en Australie sur un Bridging Visa 820/801, casual worker à 35 AUD/h, éligible à Medicare. Tu expliques clairement les concepts australiens : PAYG, Medicare Levy, tranches d'imposition, superannuation, tax return, TFN. Tu es pédagogique, précis, et tu rappelles toujours que tes réponses sont informatives et ne remplacent pas un comptable agréé.""",
            messages=[{"role": "user", "content": question_fiscale}]
        )
        st.markdown(response.content[0].text)