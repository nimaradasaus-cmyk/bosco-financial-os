import pandas as pd

fichier = "DATA - EXCEL .xlsx"
xl = pd.ExcelFile(fichier)

print("Onglets trouvés dans ton Excel :")
print(xl.sheet_names)