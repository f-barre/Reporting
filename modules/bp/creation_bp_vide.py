# Librairies
import os.path
import openpyxl

# Modules / Dépendances
# Tools
from tools.find_path import get_agency_path, get_template_path
from tools.read_json import get_dict_from_json_file
from tools.requests_tools import get_list_of_agencies
from tools.safe_actions import dprint
# Modules
from modules.bp.creation_bp_vide_entite import creation_bp_of_entite


def creation_bp_vide():
    """
    Création des feuilles de BP vides pour toutes les entités
    :return:
    """
    for agency_name in get_list_of_agencies():
        dprint(f"Création des BP vides pour {agency_name}", priority_level=3)
        dprint("Vérification de la présence ou non du BP (évite de l'écraser s'il existe)", priority_level=4)
        # On vérifie que le BP n'est pas déjà crée (à remplir que 1 fois par an)
        if not os.path.isfile(get_agency_path("BP", agency_name)):
            dprint("BP non présent, création lancée", priority_level=5)
            # Get entity db
            entity_db = get_dict_from_json_file(get_agency_path("database", agency_name))

            # Get BP template
            template = openpyxl.open(filename=get_template_path("BP"))

            # Création du BP pour l'entité
            creation_bp_of_entite(entity_db, template, get_agency_path("BP", agency_name))

        else:
            dprint("Existe déjà !", priority_level=4)





