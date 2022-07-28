# Librairies
import openpyxl

# Modules / Dépendances
# Modules
from modules.reporting.feuille_overview import creation_overview_sheet
from modules.reporting.feuilles_managers import creation_managers_sheets
# Tools
from tools.find_path import get_template_path, get_agency_path
from tools.read_json import get_dict_from_json_file
from tools.safe_actions import dprint


def creation_reporting(agency_name):
    """
    Création du reporting, de l'entié en paramètre
    :param agency_name:
    :return:
    """
    dprint("Récupération du template de reporting", priority_level=5)
    reporting_workbook = openpyxl.open(filename=get_template_path("reporting"))
    overview_sheet = reporting_workbook.worksheets[0]
    manager_sheet = reporting_workbook.worksheets[1]

    dprint("Récupération de la feuille d'excpetion des consultants forfait", priority_level=5)
    exception_workbook = openpyxl.open(filename=get_agency_path("consultant_forfait", agency_name))
    exception_sheet = exception_workbook.worksheets[0]

    dprint("Récupération de la feuille de primes", priority_level=5)
    prime_workbook = openpyxl.open(filename=get_agency_path("primes", agency_name))
    prime_sheet = prime_workbook.worksheets[0]

    dprint("Récupération de la db", priority_level=5)
    database = get_dict_from_json_file(get_agency_path("database", agency_name))

    dprint("Remplissage des feuilles de manager", priority_level=5)
    creation_managers_sheets(database, reporting_workbook, manager_sheet, exception_sheet, prime_sheet, agency_name)
    dprint("Suppression de la feuille template excedente", priority_level=5)
    reporting_workbook.remove(manager_sheet)

    dprint("Remplissage de la feuille Overview", priority_level=5)
    creation_overview_sheet(reporting_workbook, overview_sheet)

    dprint("Sauvegarde du fichier", priority_level=5)
    reporting_workbook.save(filename=get_agency_path("reporting", agency_name))
