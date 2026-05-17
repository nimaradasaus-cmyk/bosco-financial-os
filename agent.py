import anthropic
import pandas as pd
from dotenv import load_dotenv
import os

load_dotenv()
client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

# Lit ton Excel
fichier = "DATA - EXCEL .xlsx"
df = pd.read_excel(fichier, sheet_name="Money Flow", header=552, nrows=20)
df = df.dropna(how="all")
df_reel = df[df["Income"] > 0].copy()

revenu_moyen = df_reel["Income"].mean()
derniere = df_reel.iloc[-1]

cash_actuel = 4015
objectif_sydney = 8000
semaines_restantes = 5
besoin_semaine = round((objectif_sydney - cash_actuel) / semaines_restantes, 2)

# Construit le snapshot
snapshot = f"""
Dernière semaine    : {derniere['Week']}
Revenu              : {round(derniere['Income'], 2)} AUD
Dépenses réelles    : {round(derniere['Real Expenses'], 2)} AUD
Surplus W/O Savings : {round(derniere['True Surplus W/O Savings'], 2)} AUD
Revenu moyen        : {round(revenu_moyen, 2)} AUD
Cash Sydney actuel  : {cash_actuel} AUD
Objectif Sydney     : {objectif_sydney} AUD
Semaines restantes  : {semaines_restantes}
À mettre de côté    : {besoin_semaine} AUD/semaine
"""

# Envoie à Claude
message = client.messages.create(
    model="claude-sonnet-4-6",
    max_tokens=1024,
    system="Tu es l'analyste financier personnel de Nicolas, 25 ans, casual worker à Melbourne, qui prépare un déménagement à Sydney. Sois direct, structuré et prudent. Pas de blabla.",
    messages=[
        {"role": "user", "content": f"Voici mon snapshot financier cette semaine :\n{snapshot}\n\nDonne-moi un diagnostic court et une recommandation concrète."}
    ]
)

print("=== AGENT ANALYSTE ===")
print(message.content[0].text)