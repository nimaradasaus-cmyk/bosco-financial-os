import streamlit as st
import pandas as pd
import anthropic
import sqlite3
import json
import io
import os
import math
import re
from dotenv import load_dotenv
from datetime import date

load_dotenv()
client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

st.set_page_config(page_title="Bosco Financial OS", page_icon="💼", layout="wide")
st.title("💼 Bosco Financial OS")

# ─────────────────────────────────────────────────────────────────
# CONSTANTS & DATE
# ─────────────────────────────────────────────────────────────────
OBJECTIF_SYDNEY = 8000
DATE_DERNIERE_PAYE = date(2026, 6, 16)
aujourd_hui = date.today()
jours_restants = (DATE_DERNIERE_PAYE - aujourd_hui).days
SEMAINES_RESTANTES = max(1, math.ceil(jours_restants / 7))

# ─────────────────────────────────────────────────────────────────
# UPLOAD
# ─────────────────────────────────────────────────────────────────
fichier = st.file_uploader("📂 Charge ton fichier Excel (DATA - EXCEL .xlsx)", type=["xlsx"])

if fichier is None:
    st.info("⬆️ Charge ton fichier Excel pour démarrer.")
    st.stop()

# ─────────────────────────────────────────────────────────────────
# FONCTIONS LECTURE EXCEL
# ─────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_context_tab(file_bytes):
    """Lit l'onglet _Context (clé/valeur). Retourne dict ou {} si absent."""
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="_Context", header=1, usecols=[0, 1, 2])
        df.columns = ["section", "cle", "valeur"]
        df = df.dropna(subset=["cle", "valeur"])
        df = df[~df["cle"].isin(["CLÉ", "SECTION", "section", "cle"])]
        df = df[df["cle"].astype(str).str.strip() != ""]
        return dict(zip(df["cle"].astype(str).str.strip(), df["valeur"].astype(str).str.strip()))
    except Exception:
        return {}


@st.cache_data(show_spinner=False)
def load_money_flow(file_bytes, aujourd_hui):
    """Retourne les stats Money Flow. Auto-détecte le header row."""
    # Scan pour trouver le tableau résumé hebdo (cherche la ligne avec "Week" ET "Income")
    df_scan = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Money Flow", header=None)
    header_row = None
    for i, row in df_scan.iterrows():
        values = [str(v).strip() for v in row.values]
        if "Income" in values and "Week" in values:
            header_row = i
            break
    if header_row is None:
        raise ValueError("Tableau résumé introuvable dans Money Flow. Vérifie que les colonnes 'Week' et 'Income' existent.")

    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Money Flow", header=header_row, nrows=25)
    df = df.dropna(how="all")
    df["Income"] = pd.to_numeric(df["Income"], errors="coerce")
    df["True Surplus W/O Savings"] = pd.to_numeric(df["True Surplus W/O Savings"], errors="coerce")
    # Garde uniquement les lignes avec une vraie période semaine (ex: "27/1/26 -> 2/2/26")
    df_reel = df[
        df["Income"].notna() &
        (df["Income"] > 0) &
        df["Week"].astype(str).str.contains(r"\d+/\d+/\d+.*->", na=False)
    ].copy()
    if df_reel.empty:
        raise ValueError("Aucune ligne semaine valide trouvée dans le tableau résumé Money Flow.")
    derniere = df_reel.iloc[-1]

    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Money Flow", header=None)
    df_raw.columns = range(df_raw.shape[1])
    df_t = df_raw.iloc[3:, [0, 2, 3, 4, 7]].copy()
    df_t.columns = ["Week", "Date", "Description", "Flow_Type", "Amount"]
    df_t["Week"] = df_t["Week"].ffill()
    df_t = df_t[df_t["Description"].notna() & df_t["Amount"].notna()]
    df_t = df_t[pd.to_numeric(df_t["Amount"], errors="coerce").notna()]
    df_t["Amount"] = pd.to_numeric(df_t["Amount"])
    df_t = df_t[df_t["Week"].astype(str).str.match(r"\d+/\d+/\d+.*->")].copy()

    def parse_week_start(w):
        try:
            s = str(w).split("->")[0].strip().split("/")
            return date(int("20" + s[2]), int(s[1]), int(s[0]))
        except:
            return None

    df_t["week_start"] = df_t["Week"].apply(parse_week_start)
    df_t = df_t[df_t["week_start"].notna() & (df_t["week_start"] <= aujourd_hui)]
    weeks_valid = list(dict.fromkeys(df_t["Week"].tolist()))
    last_4 = weeks_valid[-4:]
    df_recent = df_t[df_t["Week"].isin(last_4)]

    txn_txt = "DÉTAIL TRANSACTIONS — 4 DERNIÈRES SEMAINES :\n"
    for week in last_4:
        wdata = df_recent[df_recent["Week"] == week]
        txn_txt += f"\nSemaine {week}:\n"
        for _, r in wdata.iterrows():
            txn_txt += f"  {str(r['Flow_Type']):25} | {str(r['Description']):38} | {r['Amount']:>8.2f} AUD\n"

    historique_txt = ""
    for _, r in df_reel.iterrows():
        historique_txt += f"  {r['Week']} | Revenu: {round(r['Income'],2):>8} | Surplus: {round(r['True Surplus W/O Savings'],2):>8}\n"

    return {
        "derniere": derniere,
        "revenu_moyen": round(df_reel["Income"].mean(), 2),
        "revenu_min": round(df_reel["Income"].min(), 2),
        "revenu_max": round(df_reel["Income"].max(), 2),
        "revenu_std": round(df_reel["Income"].std(), 2),
        "surplus_moyen": round(df_reel["True Surplus W/O Savings"].mean(), 2),
        "semaines_negatives": int((df_reel["True Surplus W/O Savings"] < 0).sum()),
        "tendance_4sem": round(df_reel["Income"].tail(4).mean(), 2),
        "nb_semaines": len(df_reel),
        "historique_txt": historique_txt,
        "txn_txt": txn_txt,
    }


@st.cache_data(show_spinner=False)
def load_headquarters(file_bytes):
    """Retourne (cash_actuel float, hq_clean DataFrame)."""
    hq = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Headquarters", header=48, nrows=24)
    hq = hq.dropna(how="all")
    hq = hq[hq["Unnamed: 1"].astype(str).str.contains(r"\d+/\d+/\d+", na=False)]
    hq_clean = hq[["Unnamed: 1", "Planned Expenses Acc"]].copy()
    hq_clean.columns = ["Date", "Cash Sydney"]
    hq_clean = hq_clean[hq_clean["Cash Sydney"] > 0]
    hq_clean["Date"] = hq_clean["Date"].astype(str).str.extract(r"(\d+/\d+/\d+)")
    hq_clean["Date_tri"] = pd.to_datetime(hq_clean["Date"], dayfirst=True)
    hq_clean = hq_clean.sort_values("Date_tri")
    cash_actuel = round(hq_clean["Cash Sydney"].iloc[-1], 2)
    return cash_actuel, hq_clean


@st.cache_data(show_spinner=False)
def load_investing(file_bytes):
    """Retourne le snapshot P/L de l'onglet Investing sous forme de DataFrame + texte."""
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Investing", header=None)
        rows_idx = [3, 4, 5, 6, 7, 9]
        labels = ["BTC", "BGBL", "VGE", "CSL", "VOO", "TOTAL"]
        records = []
        for i, r in enumerate(rows_idx):
            if r >= len(df):
                continue
            row = df.iloc[r]
            actif = str(row[37]).strip() if pd.notna(row[37]) else labels[i]
            try:
                investi = float(row[38])
                valeur  = float(row[39])
                pl      = float(row[40])
                roi     = float(row[41]) * 100
                records.append({"Actif": actif, "Investi (AUD)": round(investi, 2),
                                 "Valeur actuelle": round(valeur, 2),
                                 "P/L (AUD)": round(pl, 2), "ROI (%)": round(roi, 2)})
            except:
                pass
        df_snap = pd.DataFrame(records)

        txt = "PORTEFEUILLE INVESTISSEMENT :\n"
        for rec in records:
            sign = "+" if rec["P/L (AUD)"] >= 0 else ""
            txt += (f"  {rec['Actif']:22} | Investi: {rec['Investi (AUD)']:>8.2f} AUD"
                    f" | Valeur: {rec['Valeur actuelle']:>8.2f}"
                    f" | P/L: {sign}{rec['P/L (AUD)']:>7.2f}"
                    f" | ROI: {sign}{rec['ROI (%)']:.1f}%\n")
        return df_snap, txt
    except Exception as e:
        return pd.DataFrame(), f"[Données investing non disponibles : {e}]"


# ─────────────────────────────────────────────────────────────────
# CONSTRUCTION DU CONTEXTE GLOBAL
# ─────────────────────────────────────────────────────────────────

def build_contexte(ctx, stats, cash_actuel, besoin_semaine, investing_txt=""):
    """Contexte complet injecté dans tous les agents."""

    # Bloc statique : depuis _Context si disponible, sinon fallback
    if ctx:
        bloc_statique = "CONTEXTE PERSONNEL (depuis _Context) :\n"
        for k, v in ctx.items():
            bloc_statique += f"  {k} : {v}\n"
    else:
        bloc_statique = (
            "CONTEXTE PERSONNEL (fallback — ajoute l'onglet _Context à ton Excel) :\n"
            "  Prénom : Nicolas | Âge : 25 | Nationalité : Française\n"
            "  Visa : Bridging 820/801 | Partenaire : Lily\n"
            "  Employeur : Sapien casual 35 AUD/h | Départ : 16/06/2026\n"
            "  DCA : 15 AUD BTC + 35 AUD ETF/sem — ligne rouge absolue\n"
        )

    return f"""
=== CONTEXTE NICOLAS — {aujourd_hui} ===

{bloc_statique}

SITUATION FINANCIÈRE :
  Cash Sydney      : {cash_actuel} AUD / {OBJECTIF_SYDNEY} AUD ({round(cash_actuel/OBJECTIF_SYDNEY*100,1)}%)
  Écart restant    : {OBJECTIF_SYDNEY - cash_actuel} AUD
  Semaines restantes : {SEMAINES_RESTANTES}
  Besoin/semaine   : {besoin_semaine} AUD
  Revenu sem. passée   : {round(stats['derniere']['Income'], 2)} AUD
  Surplus sem. passée  : {round(stats['derniere']['True Surplus W/O Savings'], 2)} AUD

STATISTIQUES ({stats['nb_semaines']} semaines) :
  Revenu moyen : {stats['revenu_moyen']} AUD | Min : {stats['revenu_min']} | Max : {stats['revenu_max']}
  Volatilité   : {stats['revenu_std']} AUD (écart-type)
  Surplus moyen : {stats['surplus_moyen']} AUD/sem
  Semaines négatives : {stats['semaines_negatives']} / {stats['nb_semaines']}
  Tendance 4 sem : {stats['tendance_4sem']} AUD/sem

HISTORIQUE SEMAINES :
{stats['historique_txt']}
{investing_txt}
NOTE MÉTRIQUE : "True Surplus W/O Savings" inclut les virements Sydney volontaires.
Vraie mesure de performance = progression Cash Sydney dans Headquarters.

RÈGLES ABSOLUES :
  1. Virement Sydney en PREMIER à chaque paie
  2. DCA 50 AUD/sem = non négociable
=== FIN CONTEXTE ===
"""


# ─────────────────────────────────────────────────────────────────
# CHARGEMENT DES DONNÉES
# ─────────────────────────────────────────────────────────────────

file_bytes = fichier.read()

with st.spinner("Lecture du fichier Excel..."):
    ctx           = load_context_tab(file_bytes)
    stats         = load_money_flow(file_bytes, aujourd_hui)
    cash_actuel, hq_clean = load_headquarters(file_bytes)
    df_investing, investing_txt = load_investing(file_bytes)

besoin_semaine  = round((OBJECTIF_SYDNEY - cash_actuel) / SEMAINES_RESTANTES, 2)
contexte_global = build_contexte(ctx, stats, cash_actuel, besoin_semaine, investing_txt)

# Badge _Context
if ctx:
    st.success(f"✅ _Context chargé — {len(ctx)} entrées de contexte")
else:
    st.warning("⚠️ Onglet _Context non trouvé. Ajoute-le à ton Excel pour des analyses plus riches.")

# ─────────────────────────────────────────────────────────────────
# NAVIGATION PAR ONGLETS
# ─────────────────────────────────────────────────────────────────

tab1, tab2, tab3, tab4 = st.tabs(["💰 Finances", "📈 Investissement", "📓 Journal", "🎓 Formation"])


# ══════════════════════════════════════════════════════════════════
# TAB 1 — FINANCES
# ══════════════════════════════════════════════════════════════════
with tab1:

    # SNAPSHOT
    st.subheader("Snapshot")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Cash Sydney", f"{cash_actuel} AUD", f"Objectif {OBJECTIF_SYDNEY} AUD")
    c2.metric("Progression", f"{round(cash_actuel/OBJECTIF_SYDNEY*100,1)}%",
              f"{OBJECTIF_SYDNEY - cash_actuel} AUD restants")
    c3.metric("À mettre de côté", f"{besoin_semaine} AUD/sem",
              f"{SEMAINES_RESTANTES} semaines restantes")
    c4.metric("Revenu sem. passée",
              f"{round(stats['derniere']['Income'], 2)} AUD",
              f"Surplus : {round(stats['derniere']['True Surplus W/O Savings'], 2)} AUD")

    st.divider()

    # TRACKER
    st.subheader("Tracker Sydney")
    st.line_chart(hq_clean.set_index("Date_tri")["Cash Sydney"])

    st.divider()

    # AGENT STRATÈGE
    st.subheader("Agent Stratège")
    st.caption("Analyse complète en une passe : trajectoire, risques, biais, action prioritaire.")

    if st.button("🎯 Lancer l'analyse stratégique", type="primary"):
        with st.spinner("Analyse en cours..."):
            response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=2500,
                system=f"""Tu es le conseiller financier personnel de Nicolas. Direct, structuré, sans complaisance, sans optimisme excessif. Tu analyses uniquement à partir des données réelles fournies. Tu connais Nicolas personnellement — utilise son prénom.

RÈGLE CRITIQUE : Ne juge jamais une semaine sur son surplus net seul. Les "Saving Allocation" dans les transactions = virements volontaires vers Sydney. La vraie mesure = progression Cash Sydney.

{contexte_global}

{stats['txn_txt']}""",
                messages=[{"role": "user", "content": """Donne-moi une analyse complète structurée en 4 sections exactes :

### 1. TRAJECTOIRE
Analyse chiffrée. Suis-je sur la bonne trajectoire vers 8000 AUD ? Si non, de combien je suis en retard ou en avance, et pourquoi.

### 2. RISQUES
Liste mes 3 risques actifs par ordre de criticité :
🔴 Critique | 🟠 Élevé | 🟡 Modéré | 🟢 Faible
Pour chaque : description courte basée sur les données + action concrète.

### 3. ANGLE MORT
1 biais émotionnel ou rationalisation que je pourrais être en train de faire. Appuie-toi sur mon état d'esprit si disponible dans le contexte.

### 4. ACTION PRIORITAIRE
1 seule action concrète à faire dans les 7 prochains jours. Pas de liste. Une action."""}]
            )
            st.markdown(response.content[0].text)


# ══════════════════════════════════════════════════════════════════
# TAB 2 — INVESTISSEMENT
# ══════════════════════════════════════════════════════════════════
with tab2:

    st.subheader("Snapshot Portefeuille")

    if not df_investing.empty:
        # Affichage du tableau avec couleurs P/L
        def color_pl(val):
            if isinstance(val, (int, float)):
                color = "#d4edda" if val > 0 else ("#f8d7da" if val < 0 else "")
                return f"background-color: {color}" if color else ""
            return ""

        st.dataframe(
            df_investing.style.map(color_pl, subset=["P/L (AUD)", "ROI (%)"]),
            use_container_width=True, hide_index=True
        )
    else:
        st.warning("Données investing non disponibles.")

    st.divider()

    st.subheader("Agent Investissement")
    st.caption("Analyse de cohérence portfolio avec ta philosophie DCA long terme.")

    if st.button("📊 Analyser mon portfolio", type="primary"):
        with st.spinner("Analyse en cours..."):
            response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=1500,
                system=f"""Tu es l'analyste investissement de Nicolas. Direct, pédagogique, sans optimisme excessif.

RÈGLE ABSOLUE : Nicolas pratique le DCA long terme. Il ne fait pas de trading. Toute recommandation doit être compatible avec cette philosophie. Ne recommande jamais de vendre sur une simple réaction à une perte court terme.

{contexte_global}""",
                messages=[{"role": "user", "content": """Analyse mon portfolio en 3 sections exactes :

### 1. PERFORMANCE ACTUELLE
Commente chaque position. Ce qui performe bien et pourquoi. Ce qui sous-performe et pourquoi (si identifiable).

### 2. COHÉRENCE DU PORTFOLIO
Mon portfolio est-il cohérent avec une stratégie DCA long terme ?
Signale : concentration excessive, position orpheline (un seul achat sans suivi), actif sans thèse claire.

### 3. DÉCISION EN ATTENTE
Y a-t-il une décision à prendre sur une position ? Formule-la avec le pour et le contre. Si aucune décision urgente, dis-le clairement."""}]
            )
            st.markdown(response.content[0].text)


# ══════════════════════════════════════════════════════════════════
# TAB 3 — JOURNAL DE DÉCISIONS
# ══════════════════════════════════════════════════════════════════
with tab3:

    st.subheader("Archiver une décision")
    nouvelle = st.text_area(
        "Décris ta décision en langage naturel",
        placeholder="Ex : J'ai décidé de ne pas dépenser sur X ce mois-ci parce que..."
    )

    if st.button("💾 Archiver", type="primary"):
        if not nouvelle.strip():
            st.warning("Écris d'abord une décision.")
        else:
            with st.spinner("Archivage..."):
                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=600,
                    system="""Transforme le texte en JSON avec exactement 4 champs :
- decision : en une phrase, ce qui a été décidé
- contexte : la situation qui a mené à cette décision
- raisonnement : pourquoi cette décision plutôt qu'une autre
- invalidation : quelle condition ou information remettrait cette décision en question

Réponds UNIQUEMENT avec le JSON valide, sans backticks, sans commentaire.""",
                    messages=[{"role": "user", "content": nouvelle}]
                )
                raw = response.content[0].text.strip()
                raw = re.sub(r"^```json?\s*", "", raw)
                raw = re.sub(r"\s*```$", "", raw)
                try:
                    data = json.loads(raw)
                    conn = sqlite3.connect("bosco_financial_os.db")
                    cursor = conn.cursor()
                    cursor.execute("""CREATE TABLE IF NOT EXISTS decisions (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        date TEXT, decision TEXT, contexte TEXT,
                        raisonnement TEXT, invalidation TEXT)""")
                    cursor.execute(
                        "INSERT INTO decisions VALUES (?, ?, ?, ?, ?, ?)",
                        (None, str(date.today()), data["decision"],
                         data["contexte"], data["raisonnement"], data["invalidation"])
                    )
                    conn.commit()
                    conn.close()
                    st.success("Décision archivée.")
                    col1, col2 = st.columns(2)
                    col1.info(f"**Décision :** {data['decision']}")
                    col2.warning(f"**Invalidation :** {data['invalidation']}")
                except json.JSONDecodeError as e:
                    st.error(f"Erreur parsing JSON : {e}")
                    st.code(raw)

    st.divider()
    st.subheader("Historique des décisions")

    conn = sqlite3.connect("bosco_financial_os.db")
    cursor = conn.cursor()
    cursor.execute("""CREATE TABLE IF NOT EXISTS decisions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT, decision TEXT, contexte TEXT,
        raisonnement TEXT, invalidation TEXT)""")
    conn.commit()
    decisions = pd.read_sql(
        "SELECT date, decision, invalidation FROM decisions ORDER BY id DESC", conn
    )
    conn.close()

    if decisions.empty:
        st.info("Aucune décision archivée pour l'instant.")
    else:
        st.dataframe(decisions, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════
# TAB 4 — FORMATION
# ══════════════════════════════════════════════════════════════════
with tab4:

    st.subheader("Agent Formateur")
    st.caption("Fiscalité australienne, superannuation, visa, stratégie financière.")

    question = st.text_input(
        "Ta question",
        placeholder="Ex : Comment fonctionne la tax return en Australie ? Qu'est-ce que la super ?"
    )

    if st.button("🎓 Poser la question", type="primary"):
        if not question.strip():
            st.warning("Pose une question d'abord.")
        else:
            with st.spinner("Recherche en cours..."):
                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=2000,
                    system=f"""Tu es le conseiller financier et fiscal de Nicolas. Tu expliques clairement les concepts australiens : PAYG withholding, Medicare Levy, tranches d'imposition, superannuation (taux, accès, fonds), tax return, TFN, ABN, Bridging Visa droits travail. Tu es pédagogique, précis, et tu rappelles que tes réponses sont informatives et ne remplacent pas un comptable agréé. Tu connais Nicolas personnellement — utilise son prénom.

{contexte_global}""",
                    messages=[{"role": "user", "content": question}]
                )
                st.markdown(response.content[0].text)

