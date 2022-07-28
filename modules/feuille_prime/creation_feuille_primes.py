# Librairies
import os.path
import openpyxl

# Modules / Dépendances
# Tools
from tools.find_path import get_general_path, get_template_path
from tools.format_cell_tools import get_table_coords, display_border_table
from tools.read_json import get_dict_from_json_file
from tools.safe_actions import dprint, get_tableau_size


def update_primes_workbook(save_path, workbook, all_ressources):
    """
    Met à jour le tableau de primes: si un consultant arrive au cours de l'année, il sera ajouté au tableau
    :param save_path:
    :param workbook:
    :param liste de toutes les ressources
    :return:
    """

    def get_sheet_ressources_list(sheet):
        height = get_tableau_size(sheet, "ligne")

        sheet_ressources_list = list()
        for ligne in range(height[0] + 1, height[1] + 1):
            try:
                sheet_ressources_list.append(" ".join([sheet[f"B{ligne}"].value, sheet[f"C{ligne}"].value]))
            except:
                pass
        return sheet_ressources_list

    # On itère chaque wheet du workbook
    for sheet in workbook.worksheets:
        # On récupère toutes les ressources de la sheet
        sheet_ressources = get_sheet_ressources_list(sheet)

        # On vérifie si de nouvelles ressources sont présentes
        for ressource in all_ressources:
            if ressource not in sheet_ressources:
                sheet.insert_rows(3)
                sheet[f"B3"].value = ressource.split(" ")[0]
                sheet[f"C3"].value = " ".join(ressource.split(" ")[1:])

            # Style -> contours du tableau
            display_border_table(get_table_coords("B", 2, sheet), sheet)

    # Sauvegarde
    workbook.save(filename=save_path)


def creation_primes_sheet(workbook, all_ressources):
    """
    Crée le fichier avec les 12 tableaux de primes
    :param workbook:
    :param liste de toutes les ressources
    :return:
    """
    for month in range(1, 13):
        sheet = workbook.worksheets[month - 1]
        # On commence à la ligne 3, colonne B (Nom), C (Prénom)
        for index, ressource in enumerate(all_ressources):
            sheet[f"B{index + 3}"].value = ressource.split(" ")[0]
            sheet[f"C{index + 3}"].value = " ".join(ressource.split(" ")[1:])

        # Style -> contours du tableau
        display_border_table(get_table_coords("B", 2, sheet), sheet)


def creation_feuilles_primes():
    """
    Création de la feuille de primes, 2 modes:
        - drag and drop: dossier dans lequel on glisse les fichiers de primes
        - tableau: on créer un fichier avec 12 feuilles, 1 par mois, il sera à compléter
    :return:
    """
    # On charge la liste des ressources
    all_ressources = get_dict_from_json_file(get_general_path("ressources"))
    all_ressources = all_ressources["ressources"]

    # 2 possibilités: feuille déjà existante | feuille non existante
    dprint("Vérification de la présence ou non de la feuille de primes", priority_level=3)
    # Option 1:
    if os.path.isfile(get_general_path("primes")):
        dprint("Présente, on vérifie si un nouveau consultant a été recruté, si oui on l'ajoute au tableau",
               priority_level=4)
        # On met à jour le tableau si de nouvelles ressources on été ajouté
        workbook = openpyxl.open(filename=get_general_path("primes"))
        update_primes_workbook(get_general_path("primes"), workbook, all_ressources)

        # Option 2:
    else:
        dprint("Non présente, on crée la feuille", priority_level=4)
        # On crée le fichier
        # On récupère le template de la feuille d'exception des consultant payés au forfait
        template_workbook = openpyxl.open(filename=get_template_path("primes"))

        # Remplissage de la feuille des excpetions avec les managers
        creation_primes_sheet(template_workbook, all_ressources)

        # Save
        template_workbook.save(filename=get_general_path("primes"))
