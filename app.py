import streamlit as st
import pandas as pd
import anthropic
import sqlite3
import json
import os
import math
import re
from dotenv import load_dotenv
from datetime import date, datetime
load_dotenv()
client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
st.title("Bosco Financial OS")

# --- UPLOAD FICHIER EXCEL ---
fichier = st.file_uploader("📂 Charge ton fichier Excel (DATA - EXCEL .xlsx)", type=["xlsx"])

# --- SEMAINES RESTANTES AUTOMATIQUES ---
aujourd_hui = date.today()
date_derniere_paye = date(2026, 6, 16)
jours_restants = (date_derniere_paye - aujourd_hui).days
semaines_restantes = max(1, math.ceil(jours_restants / 7))

# --- CONTEXTE GLOBAL ---
contexte_global = ""
transactions_recentes = ""
derniere = None
cash_actuel = 0
objectif_sydney = 8000
besoin_semaine = 0

if fichier is not None:
    # Lecture Money Flow summary
    df = pd.read_excel(fichier, sheet_name="Money Flow", header=552, nrows=20)
    df = df.dropna(how="all")
    df_reel = df[df["Income"] > 0].copy()
    derniere = df_reel.iloc[-1]

    # Lecture transactions détaillées
    df_raw = pd.read_excel(fichier, sheet_name="Money Flow", header=None)
    df_raw.columns = range(df_raw.shape[1])
    df_t = df_raw.iloc[3:].copy()
    df_t = df_t[[0, 2, 3, 4, 7]].copy()
    df_t.columns = ['Week', 'Date', 'Description', 'Flow_Type', 'Amount']
    df_t['Week'] = df_t['Week'].ffill()
    df_t = df_t[df_t['Description'].notna()]
    df_t = df_t[df_t['Amount'].notna()]
    df_t = df_t[pd.to_numeric(df_t['Amount'], errors='coerce').notna()]
    df_t['Amount'] = pd.to_numeric(df_t['Amount'])
    df_t = df_t[df_t['Week'].astype(str).str.match(r'\d+/\d+/\d+.*->')].copy()

    def parse_week_start(w):
        try:
            start = str(w).split('->')[0].strip()
            parts = start.split('/')
            return date(int('20'+parts[2]), int(parts[1]), int(parts[0]))
        except:
            return None

    df_t['week_start'] = df_t['Week'].apply(parse_week_start)
    df_t = df_t[df_t['week_start'].notna()]
    df_t = df_t[df_t['week_start'] <= aujourd_hui]
    weeks_valid = list(dict.fromkeys(df_t['Week'].tolist()))
    last_4_weeks = weeks_valid[-4:]
    df_recent = df_t[df_t['Week'].isin(last_4_weeks)]

    transactions_recentes = "DÉTAIL TRANSACTIONS — 4 DERNIÈRES SEMAINES :\n"
    for week in last_4_weeks:
        week_data = df_recent[df_recent['Week'] == week]
        transactions_recentes += f"\nSemaine {week}:\n"
        for _, row in week_data.iterrows():
            transactions_recentes += f"  {str(row['Flow_Type']):25} | {str(row['Description']):40} | {row['Amount']:>8.2f} AUD\n"

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

    historique = ""
    for _, row in df_reel.iterrows():
        historique += f"  {row['Week']} | Revenu: {round(row['Income'],2)} | Surplus: {round(row['True Surplus W/O Savings'],2)}\n"

    contexte_global = f"""
=== CONTEXTE NICOLAS — MIS À JOUR DEPUIS EXCEL ===

PROFIL :
- Nicolas, 25 ans, français, Melbourne, Bridging Visa 820/801
- Casual worker chez Sapien à 35 AUD/h, revenus variables
- Dernière paye Sapien : 16/06/2026 — date non négociable
- DCA hebdomadaire : 15 AUD BTC + 35 AUD ETF (50 AUD/sem)
- À Sydney : ~2 semaines chez les parents de Lily, puis bond + loyer 50/50

PRIORITÉ FINANCIÈRE :
- À chaque paye : virement Sydney buffer EN PREMIER, avant toute autre dépense
- DCA (50 AUD/sem) = ligne rouge — discipline non négociable, pas une variable d'ajustement
- L'objectif Sydney prime sur tout sauf le DCA

NOTE IMPORTANTE SUR LE SURPLUS HEBDOMADAIRE :
- Le "True Surplus W/O Savings" peut sembler bas certaines semaines
- C'est parce qu'il inclut les virements volontaires vers le Sydney buffer (Saving Allocation)
- La vraie mesure de performance est la progression du cash Sydney dans Headquarters
- Ne pas interpréter un surplus bas comme un relâchement — vérifier les mouvements réels

SITUATION ACTUELLE :
- Cash Sydney : {cash_actuel} AUD
- Objectif Sydney : {objectif_sydney} AUD
- Écart restant : {objectif_sydney - cash_actuel} AUD
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
- Boss : Arky — relation de confiance réelle, discussions fréquentes et profondes
- Les heures sont auto-gérées via Xero.me — Arky ne les contrôle pas à la loupe
- Au workshop : heures flexibles, toujours du travail disponible
- Nicolas contrôle lui-même son volume d'heures
- Arky valorise la vision business de Nicolas (systèmes, standards, scaling)
- L'annonce du départ Sydney peut changer la dynamique mais pas de manière certaine
- Tam absente 1 mois fin mai — vide opérationnel réel
- Abdul : bon travailleur avec de vrais skills, fiable en exécution, mais pas encore un leader autonome
- Business Sapien actuellement dans le rouge
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
            max_tokens=2000,
            system=f"""Tu es l'analyste financier personnel de Nicolas. Direct, structuré, prudent. Tu analyses sa situation à partir de données réelles. Tu connais Nicolas personnellement — utilise son prénom.

RÈGLE CRITIQUE : Ne juge jamais une semaine sur son surplus net seul. Lis les transactions détaillées pour comprendre les vrais mouvements — en particulier les "Saving Allocation" qui représentent des virements réels vers Sydney, et les "Joint Account" liés au déménagement.

{contexte_global}

{transactions_recentes}""",
            messages=[{"role": "user", "content": "En te basant sur mes transactions réelles et la progression de mon cash Sydney, suis-je sur la bonne trajectoire pour atteindre 8000 AUD d'ici le 16 juin ? Qu'est-ce qui menace cette trajectoire et quelle est l'action prioritaire ?"}]
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
            max_tokens=2000,
            system=f"""Tu es le Challenger financier de Nicolas. Ton rôle est de trouver les failles dans son raisonnement. Tu ne valides pas, tu questionnes. Tu cherches : les biais émotionnels, les rationalisations, les informations manquantes, les risques ignorés. Tu es direct et sans complaisance. Tu connais Nicolas personnellement — utilise son prénom.

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
            max_tokens=2000,
            system=f"""Tu es le Risk Manager personnel de Nicolas. Ton rôle est d'identifier et prioriser les risques financiers actifs à partir de données réelles. Tu connais Nicolas personnellement — utilise son prénom.

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
            max_tokens=2000,
            system=f"""Tu es le conseiller financier et fiscal de Nicolas. Tu expliques clairement les concepts australiens : PAYG, Medicare Levy, tranches d'imposition, superannuation, tax return, TFN. Tu es pédagogique, précis, et tu rappelles toujours que tes réponses sont informatives et ne remplacent pas un comptable agréé. Tu connais Nicolas personnellement — utilise son prénom.

{contexte_global}""",
            messages=[{"role": "user", "content": question_fiscale}]
        )
        st.markdown(response.content[0].text)