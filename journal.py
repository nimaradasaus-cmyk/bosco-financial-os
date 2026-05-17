import sqlite3
import anthropic
import json
import os
from dotenv import load_dotenv
from datetime import date

load_dotenv()
client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

conn = sqlite3.connect("bosco_financial_os.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS decisions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT,
    decision TEXT,
    contexte TEXT,
    raisonnement TEXT,
    invalidation TEXT
)
""")

print("=== ARCHIVISTE — Nouvelle décision ===")
print("")
texte = input("Décris ta décision en quelques phrases : ")

response = client.messages.create(
    model="claude-sonnet-4-6",
    max_tokens=500,
    system="""Tu es l'Archiviste financier de Nicolas. 
Transforme sa description en JSON structuré avec exactement ces 4 champs :
- decision : la décision en une phrase
- contexte : la situation au moment de la décision
- raisonnement : pourquoi cette décision
- invalidation : quel signal indiquerait que c'était une mauvaise décision
Réponds UNIQUEMENT avec le JSON, rien d'autre.""",
    messages=[{"role": "user", "content": texte}]
)

raw = response.content[0].text
raw = raw.strip()
if raw.startswith("```"):
    raw = raw.split("```")[1]
    if raw.startswith("json"):
        raw = raw[4:]
data = json.loads(raw.strip())

cursor.execute("""
INSERT INTO decisions (date, decision, contexte, raisonnement, invalidation)
VALUES (?, ?, ?, ?, ?)
""", (str(date.today()), data["decision"], data["contexte"], data["raisonnement"], data["invalidation"]))

conn.commit()

print("")
print("Décision archivée :")
print(f"Décision    : {data['decision']}")
print(f"Invalidation: {data['invalidation']}")

conn.close()