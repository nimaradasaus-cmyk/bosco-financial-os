import pandas as pd

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

print("=== BOSCO FINANCIAL SNAPSHOT ===")
print("")
print(f"Dernière semaine    : {derniere['Week']}")
print(f"Revenu              : {round(derniere['Income'], 2)} AUD")
print(f"Dépenses réelles    : {round(derniere['Real Expenses'], 2)} AUD")
print(f"Surplus W/O Savings : {round(derniere['True Surplus W/O Savings'], 2)} AUD")
print("")
print(f"Revenu moyen        : {round(revenu_moyen, 2)} AUD")
print("")
print(f"Cash Sydney actuel  : {cash_actuel} AUD")
print(f"Objectif Sydney     : {objectif_sydney} AUD")
print(f"Semaines restantes  : {semaines_restantes}")
print(f"À mettre de côté    : {besoin_semaine} AUD/semaine")