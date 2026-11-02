from openpyxl import load_workbook

SOURCE = r"S:\Publics\DISC_KB\SOS\@Validation_APC\2026_Validation_APC\2026_Validation_APC.xlsx"
DEST = r"S:\Publics\DISC_KB\SOS\SOS_HAL_FullText\HAL_API_APC.xlsx"
SORTIE = DEST.replace(".xlsx", "_maj.xlsx")

# Numéros de colonnes (à partir de 1) : A=1, J=10, Y=25, Z=26
COL_CLE_DEST = 1
COL_CIBLE_DEST = 10
COL_VALEUR_SRC = 25
COL_CLE_SRC = 26

def norm(v):
    """Uniformise une clé : texte sans espaces superflus (123.0 -> '123')."""
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()

# 1) Lecture de la 4e feuille de la source (valeurs, pas formules)
wb_src = load_workbook(SOURCE, read_only=True, data_only=True)
ws_src = wb_src.worksheets[3]

correspondance = {}
for row in ws_src.iter_rows(min_row=2, values_only=True):
    cle = norm(row[COL_CLE_SRC - 1]) if len(row) >= COL_CLE_SRC else None
    if cle and cle not in correspondance:  # première occurrence, comme RECHERCHEX
        correspondance[cle] = row[COL_VALEUR_SRC - 1] if len(row) >= COL_VALEUR_SRC else None
wb_src.close()

# 2) Ouverture du destinataire (mise en forme conservée)
wb = load_workbook(DEST)
ws = wb.worksheets[0]

non_trouves = 0
for r in range(2, ws.max_row + 1):
    cle = norm(ws.cell(row=r, column=COL_CLE_DEST).value)
    if not cle:
        continue
    if cle in correspondance:
        ws.cell(row=r, column=COL_CIBLE_DEST).value = correspondance[cle]
    else:
        ws.cell(row=r, column=COL_CIBLE_DEST).value = "Non trouvé"
        non_trouves += 1

wb.save(SORTIE)
print(f"Terminé : {SORTIE}")
print(f"Identifiants non trouvés : {non_trouves}")