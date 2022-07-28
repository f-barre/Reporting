# Librairies
import openpyxl

# Modules / Dépendances
# Tools
from tools.find_path import get_agency_path
from tools.requests_tools import get_list_of_agencies
from tools.safe_actions import dprint
# Modules
from modules.calcul_avancement.calcul_avancement import calculer_avancement_reporting


def d_calcul_avancement():
    """
    Fontion d'appel pour le calcul de l'avancement de chaque reporting par rapport aux prévisions du BP
    :return:
    """
    dprint("Btn D pressed: run d_calcul_avancement", priority_level=1)
    dprint("Calcul de l'avancement des reporting pour toutes les entités", priority_level=2)
    for agency_name in get_list_of_agencies():
        dprint(f"Calcul de l'avancement du reporting de {agency_name}", priority_level=2)

        dprint("Récupération du BP", priority_level=3)
        bp_workbook = openpyxl.open(filename=get_agency_path("BP", agency_name))

        dprint("Récupération du reporting", priority_level=3)
        reporting_workbook = openpyxl.open(filename=get_agency_path("reporting", agency_name))

        dprint("Calcul de l'avancement", priority_level=3)
        try:
            calculer_avancement_reporting(bp_workbook, reporting_workbook, agency_name)
        except:
            print("BP non renseigné", agency_name)
