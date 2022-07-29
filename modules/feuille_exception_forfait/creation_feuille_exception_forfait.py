# Librairies
import os.path
import openpyxl

# Modules / Dépendances
# Tools
from tools.find_path import get_general_path, get_template_path
from tools.format_cell_tools import get_table_coords, display_border_table
from tools.read_json import get_dict_from_json_file
from tools.safe_actions import dprint, get_tableau_size


def update_exception_sheet(workbook, sheet, all_ressources):
    def get_sheet_ressources_list(sheet):
        # On récupère les dimensions du tableau
        height = get_tableau_size(sheet, "ligne")

        sheet_ressources_list = list()
        for ligne in range(height[0] + 1, height[1] + 1):
            try:
                sheet_ressources_list.append(" ".join([sheet[f"B{ligne}"].value, sheet[f"C{ligne}"].value]))
            except:
                pass
        return sheet_ressources_list

    # On récupère sous forme de liste toutes les ressources du  tableau
    sheet_ressources_list = get_sheet_ressources_list(sheet)

    # On vérifie si de nouvelles ressources sont présentes
    for ressource in all_ressources:
        if ressource not in sheet_ressources_list:
            sheet.insert_rows(3)
            sheet[f"B3"].value = ressource.split(" ")[0]
            sheet[f"C3"].value = " ".join(ressource.split(" ")[1:])

    # Style -> contours du tableau
    display_border_table(get_table_coords("B", 2, sheet), sheet)

    # Sauvegarde
    workbook.save(filename=get_general_path("consultant_forfait"))


def fill_exception_sheet(sheet, all_ressources):
    # On commence à la ligne 3, colonne B (Nom), C (Prénom)
    for index, ressource in enumerate(all_ressources):
        sheet[f"B{index + 3}"].value = ressource.split(" ")[0]
        sheet[f"C{index + 3}"].value = " ".join(ressource.split(" ")[1:])


def creation_feuille_exception_forfait():
    # On charge la liste des ressources
    all_ressources = get_dict_from_json_file(get_general_path("ressources"))
    all_ressources = all_ressources["ressources"]

    # 2 possibilités: feuille déjà existante | feuille non existante
    dprint("Vérification de la présence ou non de la feuille d'exception des consultants payés au forfait", priority_level=3)
    # Option 1:
    if os.path.isfile(get_general_path("consultant_forfait")):
        dprint("Présente, on vérifie si un nouveau consultant a été recruté, si oui on l'ajoute au tableau", priority_level=4)
        # On met à jour le tableau si de nouvelles ressources on été ajouté
        workbook = openpyxl.open(filename=get_general_path("consultant_forfait"))
        sheet = workbook.worksheets[0]
        update_exception_sheet(workbook, sheet, all_ressources)
    # Option 2:
    else:
        dprint("Non présente, on crée la feuille", priority_level=4)
        # On crée le fichier
        # On récupère le template de la feuille d'exception des consultant payés au forfait
        template_workbook = openpyxl.open(filename=get_template_path("consultant_forfait"))
        template_sheet = template_workbook.worksheets[0]

        # Remplissage de la feuille des excpetions avec les managers
        fill_exception_sheet(template_sheet, all_ressources)

        # Style -> contours du tableau
        display_border_table(get_table_coords("B", 2, template_sheet), template_sheet)

        # Save
        template_workbook.save(filename=get_general_path("consultant_forfait"))
