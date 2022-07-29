# Modules / Dépendances
from tools.date_info import get_year


def fill_dates_table(sheet, year, ligne):
    # Date
    for colonne in range(3, 15):
        sheet.cell(row=ligne, column=colonne).value = (sheet.cell(row=ligne, column=colonne).value).replace("X", year)


def fill_suivi_bp_sheets(sheet,year):
    fill_dates_table(sheet, year, 2)
    fill_dates_table(sheet, year, 17)


def fill_titles_table(sheet, year, name, ligne):
    # Name
    sheet.cell(row=ligne, column=2).value = (sheet.cell(row=ligne, column=2).value).replace("X", name)
    # Date
    for colonne in range(3, 15):
        sheet.cell(row=ligne, column=colonne).value = (sheet.cell(row=ligne, column=colonne).value).replace("X", year)


def fill_bp_sheets(database, workbook, template_bp, year):

    for manager in database:
        # Duplicate template
        new_sheet = workbook.copy_worksheet(template_bp)

        # Title
        new_sheet.title = (new_sheet.title).replace("X  Copy",  manager["manager"]["nom"])

        # First table
        fill_titles_table(new_sheet, year, manager["manager"]["nom"], 2)

        # Second table
        fill_titles_table(new_sheet, year, manager["manager"]["nom"], 14)


def creation_bp_of_entite(entity_database, template_workbook, save_path):
    # On récupère le template du BP, 2 feuilles différentes: suivi | BP
    template_suivi_bp = template_workbook.worksheets[0]
    template_bp = template_workbook.worksheets[1]

    # Remplissage BP
    fill_bp_sheets(entity_database, template_workbook, template_bp, get_year())

    # Suppression de la feuille BP vide du template
    template_workbook.remove(template_bp)

    # Remplissage Suivi BP (bp récap)
    fill_suivi_bp_sheets(template_suivi_bp, get_year())

    # Sauvegarde du BP
    template_workbook.save(filename=save_path)
