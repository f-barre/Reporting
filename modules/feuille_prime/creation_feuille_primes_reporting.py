# Librairies
import openpyxl

# Modules / Dépendances
from configuration import APP_CONFIG
# Tools
from tools.date_info import get_month
from tools.find_path import get_general_path, get_agency_path
from tools.format_cell_tools import display_border, update_euro
from tools.read_json import get_dict_from_json_file
from tools.safe_actions import dprint, safe_dict_get, get_ligne_values, get_tableau_size


def add_ressources_prime_in_reporting_prime(prime_sheet, prime_sheet_reporting, ressources_reporting):
    """
    Permet d'ajouter les primes du tableau générale vers
    le tableau récap du mois du reporting
    :param prime_sheet:
    :param prime_sheet_reporting:
    :param ressources_reporting:
    :return:
    """
    # On récupère les dimensions du tableau de primes
    height = get_tableau_size(prime_sheet, "ligne")
    width = get_tableau_size(prime_sheet, "col")

    # On itère chaque ressource de la feuille d'exception
    # TODO: int(dims_ligne[1])-2 bizare logiquement ça dervait être +1
    for ligne in range(height[0] + 1, height[1]):
        # On vérifie que la ressource appartient à l'entité / agence
        ressource = " ".join([prime_sheet[f"B{ligne}"].value, prime_sheet[f"C{ligne}"].value])
        if ressource in ressources_reporting:

            # On vérifie que la ressource est associé à un montant:
            ligne_data = get_ligne_values(prime_sheet, ligne, 2, width[0], width[1])
            if any(list(ligne_data.values())[2:]):

                # On ajoute la ressource à la feuille de primes du reporting
                dprint(f"Ajout de(s) prime(s) de : {list(ligne_data.values())[0]}", priority_level=7)
                prime_sheet_reporting.insert_rows(3)

                for index, value in enumerate(ligne_data.values()):
                    try:
                        update_euro(prime_sheet_reporting.cell(row=3, column=index + 2), value)
                    except:
                        pass


def creation_feuille_primes_reporting(agency_name):
    """
    Création de la feuille récap des primes
    du mois du reporting (tableau utilisé dans le calcul du reporting)
    :param agency_name:
    :return:
    """
    dprint("Récupération de la liste des ressources du reporting", priority_level=6)
    ressources_reporting = safe_dict_get(get_dict_from_json_file(get_agency_path("ressources", agency_name)),
                                         ["ressources"])

    dprint("Récupération du fichier des primes (fichier général)", priority_level=6)
    prime_workbook = openpyxl.open(filename=get_general_path("primes"))
    if APP_CONFIG.MODE_PRIMES == "tableau":
        prime_sheet = prime_workbook.worksheets[int(get_month()) - 1]
    elif APP_CONFIG.MODE_PRIMES == "drag and drop":
        prime_sheet = prime_workbook.worksheets[0]

    dprint("Récupération du fichier des primes du reporting", priority_level=6)
    prime_workbook_reporting = openpyxl.open(filename=get_agency_path("primes", agency_name))
    prime_sheet_reporting = prime_workbook_reporting.worksheets[0]

    dprint("Ajout des ressources qui sont liés à de(s) prime(s) dans le fichier récap", priority_level=6)
    add_ressources_prime_in_reporting_prime(prime_sheet, prime_sheet_reporting, ressources_reporting)
    dprint("Ajout du style au tableau récap", priority_level=6)
    display_border("B", 2, prime_sheet_reporting)

    dprint("Sauvegarde fichier", priority_level=6)
    prime_workbook_reporting.save(filename=get_agency_path("primes", agency_name))
