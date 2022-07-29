# Modules / Dépendances
# Tools
from tools.date_info import get_month
from tools.find_path import get_agency_path
from tools.safe_actions import dprint
# Modules
from modules.calcul_avancement.calcul_avancement_manager import calcul_avancement_manager
from modules.calcul_avancement.calcul_avancement_overview import calcul_avancement_overview


def calculer_avancement_reporting(bp_workbook, reporting_workbook, agency_name):
    """
    Calcul tous les avancements du reporting par rapport au BP
        - Dans les feuilles managers
        - Dans la feuille d'overview
    :param bp_workbook:
    :param reporting_workbook:
    :param agency_name:
    :return:
    """
    dprint("Calcul des avancements dans les feuilles managers", priority_level=4)
    calcul_avancement_manager(bp_workbook, reporting_workbook, int(get_month()), agency_name)
    dprint("Calcul des avancements dans la feuille d'overview", priority_level=4)
    calcul_avancement_overview(bp_workbook, reporting_workbook, int(get_month()))
    dprint("Sauvegardes des modififcations", priority_level=4)
    reporting_workbook.save(filename=get_agency_path("reporting", agency_name))
