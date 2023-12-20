import openpyxl

# Charger le fichier Excel
excel_file = openpyxl.load_workbook('fraude_vBmh.xlsx')

# Sélectionner la feuille de calcul (par exemple, la feuille qui se nomme Feuil1)
test1_sheet = excel_file["Feuil1"]

# Ajouter une formule dans une cellule (par exemple, addition de deux cellules)
test1_sheet['C1'] = '=A1 + B1'

# Sauvegarder les modifications dans un nouveau fichier Excel sans modifier le précedant
excel_file.save('resultat_fraude.xlsx')
print("Le fichier a bien été modifiés")
