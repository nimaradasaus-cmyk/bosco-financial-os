import streamlit as st
import pandas as pd
import anthropic
import sqlite3
import json
import os
from dotenv import load_dotenv
from datetime import date, datetime
load_dotenv()
client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
st.title("Bosco Financial OS")

# --- UPLOAD FICHIER EXCEL ---
fichier = st.file_uploader("📂 Charge ton fichier Excel (DATA - EXCEL .xlsx)", type=["xlsx"])

# --- SEMAINES RESTANTES AUTOMATIQUES ---
date_derniere_paye = datetime(2026, 6, 16)
aujourd_hui = date.today()
date_derniere_paye = date(2026, 6, 16)
jours_restants = (date_derniere_paye - aujourd_hui).days
import math
semaines_restantes = max(1, math.ceil(jours_restants / 7))

# --- CONTEXTE GLOBAL (calculé une seule fois si Excel chargé) ---
contexte_global = ""
derniere = None
cash_actuel = 0
objectif_sydney = 8000
besoin_semaine = 0

if fichier is not None:
    # Lecture Money Flow
    df = pd.read_excel(fichier, sheet_name="Money Flow", header=552, nrows=20)
    df = df.dropna(how="all")
    df_reel = df[df["Income"] > 0].copy()
    derniere = df_reel.iloc[-1]

    # Lecture Headquarters
    hq = pd.read_excel(fichier, sheet_name="Headquarters", header=48, nrows=24)
    hq = hq.dropna(how="all")
    hq = hq[hq["Unnamed: 1"].astype(str).str.contains(r"\d+/\d+/\d+", na=False)]
    hq_clean = hq[["Unnamed: 1", "Planned Expenses Acc"]].copy()
    hq_clean.columns = ["Date", "Cash Sydney"]
    hq_clean = hq_clean[hq_clean["Cash Sydney"] > 0]
    hq_clean["Date"] = hq_clean["Date"].astype(str).str.extract(r"(\d+/\d+/\d+)")
    hq_clean["Date_tri"] = pd.to_datetime(hq_clean["Date"], dayfirst=True)
    hq_clean = hq_clean.sort_values("Date_tri")
    cash_actuel = round(hq_clean["Cash Sydney"].iloc[-1], 2)
    besoin_semaine = round((objectif_sydney - cash_actuel) / semaines_restantes, 2)

    # Statistiques historiques
    revenu_moyen = round(df_reel["Income"].mean(), 2)
    revenu_min = round(df_reel["Income"].min(), 2)
    revenu_max = round(df_reel["Income"].max(), 2)
    revenu_std = round(df_reel["Income"].std(), 2)
    surplus_moyen = round(df_reel["True Surplus W/O Savings"].mean(), 2)
    semaines_negatives = int((df_reel["True Surplus W/O Savings"] < 0).sum())
    tendance_4sem = round(df_reel["Income"].tail(4).mean(), 2)
    nb_semaines = len(df_reel)

    # Historique semaine par semaine
    historique = ""
    for _, row in df_reel.iterrows():
        historique += f"  {row['Week']} | Revenu: {round(row['Income'],2)} | Surplus: {round(row['True Surplus W/O Savings'],2)}\n"

    # Bloc contexte global partagé par tous les agents
    contexte_global = f"""
=== CONTEXTE NICOLAS — MIS À JOUR DEPUIS EXCEL ===

PROFIL :
- Nicolas, 25 ans, français, Melbourne, Bridging Visa 820/801
- Casual worker chez Sapien à 35 AUD/h, revenus variables
- Dernière paye Sapien : 16/06/2026 — date non négociable
- DCA hebdomadaire : 15 AUD BTC + 35 AUD ETF (50 AUD/sem)
- À Sydney : ~2 semaines chez les parents de Lily, puis bond + loyer 50/50

SITUATION ACTUELLE :
- Cash Sydney : {cash_actuel} AUD
- Objectif Sydney : {objectif_sydney} AUD
- Écart : {objectif_sydney - cash_actuel} AUD
- Semaines de revenu restantes : {semaines_restantes} (calculé automatiquement)
- Besoin par semaine : {besoin_semaine} AUD
- Revenu dernière semaine : {round(derniere['Income'], 2)} AUD
- Surplus dernière semaine : {round(derniere['True Surplus W/O Savings'], 2)} AUD

HISTORIQUE ({nb_semaines} semaines — jan à mai 2026) :
{historique}
STATISTIQUES :
- Revenu moyen : {revenu_moyen} AUD/sem
- Revenu min : {revenu_min} AUD | max : {revenu_max} AUD
- Volatilité revenus : {revenu_std} AUD (écart-type)
- Surplus moyen W/O Savings : {surplus_moyen} AUD/sem
- Semaines à surplus négatif : {semaines_negatives} sur {nb_semaines}
- Tendance revenus 4 dernières semaines : {tendance_4sem} AUD/sem

CONTEXTE SAPIEN :
- Collègue Tam absente fin mai (Europe 1 mois)
- Collègue Abdul en départ imminent
- Risque de réduction d'heures dans les dernières semaines
=== FIN CONTEXTE ===
"""

# --- SNAPSHOT ---
if fichier is not None:
    st.header("Snapshot")
    col1, col2, col3 = st.columns(3)
    col1.metric("Cash Sydney", f"{cash_actuel} AUD", f"Objectif {objectif_sydney}")
    col2.metric("Surplus dernière semaine", f"{round(derniere['True Surplus W/O Savings'], 2)} AUD")
    col3.metric("À mettre de côté", f"{besoin_semaine} AUD/sem", f"{semaines_restantes} semaines restantes")
else:
    st.info("⬆️ Charge ton fichier Excel pour accéder au Snapshot, à l'Analyste et au Tracker Sydney.")

# --- AGENT ANALYSTE ---
st.header("Agent Analyste")
if st.button("Analyser ma situation"):
    with st.spinner("Analyse en cours..."):
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1024,
            system=f"""Tu es l'analyste financier personnel de Nicolas. Direct, structuré, prudent. Tu analyses sa situation à partir de données réelles.

{contexte_global}""",
            messages=[{"role": "user", "content": "Donne-moi un diagnostic complet de ma situation financière et une recommandation concrète pour les semaines restantes."}]
        )
        st.markdown(response.content[0].text)

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
            system=f"""Tu es le Challenger financier de Nicolas. Ton rôle est de trouver les failles dans son raisonnement. Tu ne valides pas, tu questionnes. Tu cherches : les biais émotionnels, les rationalisations, les informations manquantes, les risques ignorés. Tu es direct et sans complaisance.

{contexte_global}""",
            messages=[{"role": "user", "content": f"Voici ma décision : {decision_a_challenger}\n\nChallenge-moi."}]
        )
        st.markdown(response.content[0].text)

# --- TRACKER SYDNEY ---
st.header("Tracker Sydney")
if fichier is not None:
    st.line_chart(hq_clean.set_index("Date_tri")["Cash Sydney"])
    col1, col2, col3 = st.columns(3)
    col1.metric("Cash Sydney actuel", f"${cash_actuel}")
    col2.metric("Objectif", f"${objectif_sydney}")
    col3.metric("À mettre de côté", f"${besoin_semaine}/sem", f"{semaines_restantes} semaines restantes")
else:
    st.warning("Charge d'abord ton fichier Excel.")

# --- AGENT RISK MANAGER ---
st.header("Agent Risk Manager")
if st.button("Évaluer mes risques"):
    with st.spinner("Évaluation en cours..."):
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1500,
            system=f"""Tu es le Risk Manager personnel de Nicolas. Ton rôle est d'identifier et prioriser les risques financiers actifs à partir de données réelles.

{contexte_global}

Pour chaque risque identifié, donne :
- Niveau : 🔴 Critique / 🟠 Élevé / 🟡 Modéré / 🟢 Faible
- Description courte basée sur les données réelles
- Action concrète recommandée

Termine par une recommandation DCA : maintenir / réduire / couper temporairement.
Sois direct, sans complaisance, sans optimisme excessif.""",
            messages=[{"role": "user", "content": "Identifie et priorise mes risques financiers actuels."}]
        )
        st.markdown(response.content[0].text)

# --- AGENT FORMATEUR ---
st.header("Agent Formateur")
question_fiscale = st.text_input("Pose une question sur la fiscalité australienne, ta super, ton visa...")
if st.button("Poser la question"):
    with st.spinner("Recherche en cours..."):
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1024,
            system=f"""Tu es le conseiller financier et fiscal de Nicolas. Tu expliques clairement les concepts australiens : PAYG, Medicare Levy, tranches d'imposition, superannuation, tax return, TFN. Tu es pédagogique, précis, et tu rappelles toujours que tes réponses sont informatives et ne remplacent pas un comptable agréé.

{contexte_global}""",
            messages=[{"role": "user", "content": question_fiscale}]
        )
        st.markdown(response.content[0].text)