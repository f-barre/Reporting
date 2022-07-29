# Librairies
import openpyxl

# Modules / Dépendances
# Tools
from tools.find_path import get_general_path, get_agency_path
from tools.format_cell_tools import display_border, update_euro
from tools.read_json import get_dict_from_json_file
from tools.safe_actions import dprint, safe_dict_get, get_tableau_size


def add_ressources_exception_in_reporting_exception(exception_sheet, ressources_exception, ressources_reporting):
    """
    Permet d'ajouter une exception du tableau générale vers
    le tableau récap du mois du reporting
    :param exception_sheet:
    :param ressources_exception:
    :param ressources_reporting:
    :return:
    """
    # On récupère les dimensions du tableau d'exception
    height = get_tableau_size(exception_sheet, "ligne")

    # On itère chaque ressource de la feuille d'exception
    for ligne in range(height[0] + 1, height[1] + 1):
        # On vérifie que la ressource est associé à un motant:
        if exception_sheet[f"D{ligne + 3}"].value is not None:
            # On vérifie que la ressource appartient à l'entité / agence
            ressource = " ".join([exception_sheet[f"B{ligne + 3}"].value, exception_sheet[f"C{ligne + 3}"].value])
            if ressource in ressources_reporting:
                # On ajoute la ressource à la feuille d'exception du reporting
                dprint(f"Ajout de l'exception: {exception_sheet[f'B{ligne + 3}'].value}", priority_level=7)
                ressources_exception.insert_rows(3)
                ressources_exception[f"B{3}"].value = exception_sheet[f"B{ligne + 3}"].value
                ressources_exception[f"C{3}"].value = exception_sheet[f"C{ligne + 3}"].value
                update_euro(ressources_exception[f"D{3}"], exception_sheet[f"D{ligne + 3}"].value)


def creation_feuille_exception_forfait_du_reporting(agency_name):
    """
    Création de la feuille récap des consultants forfait
    du mois du reporting (tableau utilisé dans le calcul du reporting)
    :param agency_name:
    :return:
    """
    dprint("Récupération de la liste des ressources du reporting", priority_level=6)
    ressources_reporting = safe_dict_get(get_dict_from_json_file(get_agency_path("ressources", agency_name)),
                                         ["ressources"])

    dprint("Récupération du fichier des exceptions des consultants payés au forfait du reporting", priority_level=6)
    exception_workbook_reporting = openpyxl.open(filename=get_agency_path("consultant_forfait", agency_name))
    exception_sheet_reporting = exception_workbook_reporting.worksheets[0]

    dprint("Récupération du fichier des exceptions de consultants payés au forfait (fichier général)", priority_level=6)
    exception_workbook = openpyxl.open(filename=get_general_path("consultant_forfait"))
    exception_sheet = exception_workbook.worksheets[0]

    dprint("Ajout des ressources payés au forfait dans le fichier récap", priority_level=6)
    add_ressources_exception_in_reporting_exception(exception_sheet, exception_sheet_reporting, ressources_reporting)
    dprint("Ajout du style au tableau récap", priority_level=6)
    display_border("B", 2, exception_sheet_reporting)

    dprint("Sauvegarde fichier", priority_level=6)
    exception_workbook_reporting.save(filename=get_agency_path("consultant_forfait", agency_name))
