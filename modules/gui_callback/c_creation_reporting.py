# Librairies
import os.path

# Modules / Dépendances
# Tools
from tools.find_path import get_template_path, get_agency_path
from tools.requests_tools import get_list_of_agencies
from tools.safe_actions import dprint, copy_template_file
# modules
from modules.feuille_exception_forfait.creation_feuille_exception_forfait_du_reporting import \
    creation_feuille_exception_forfait_du_reporting
from modules.feuille_prime.creation_feuille_primes_reporting import creation_feuille_primes_reporting
from modules.reporting.creation_reporting import creation_reporting


def c_completer_reporting_mensuel():
    """
    Fontion d'appel pour la création des reporting
    :return:
    """

    dprint("Btn C pressed: run c_completer_reporting_mensuel", priority_level=1)
    dprint("Création des reporting pour toutes les entités", priority_level=2)
    for agency_name in get_list_of_agencies():
        dprint(f"Création du reporting de {agency_name}", priority_level=3)
        # On créer un dossier qui va contenir le reporting
        dprint("Vérification de la présence ou non du dossier du reporting", priority_level=4)
        if not os.path.isdir(get_agency_path("dossier_reporting", agency_name)):
            dprint("Non présent, création du dossier lancée", priority_level=4)
            os.makedirs(get_agency_path("dossier_reporting", agency_name))
        else:
            dprint("Déjà présent !", priority_level=4)

        # On créer la feuille d'exceptions des consultants payés au forfait qui seront dans le reporting
        dprint("Vérification de la présence ou non du fichier des consultants forfait", priority_level=4)
        if not os.path.isfile(get_agency_path("consultant_forfait", agency_name)):
            dprint("Non présent, création du fichier lancée", priority_level=4)
            copy_template_file(get_template_path("consultant_forfait"),
                               get_agency_path("consultant_forfait", agency_name))
        else:
            dprint("Déjà présent !", priority_level=4)

        # On complète la feuille d'exceptions
        dprint("Remplissage de la feuille des consultants forfait", priority_level=5)
        creation_feuille_exception_forfait_du_reporting(agency_name)

        # On créer la feuille des primes des consultants dans le reporting
        dprint("Vérification de la présence ou non du fichier des primes", priority_level=4)
        if not os.path.isfile(get_agency_path("primes", agency_name)):
            dprint("Non présent, création du fichier lancée", priority_level=4)
            copy_template_file(get_template_path("primes_reporting"), get_agency_path("primes", agency_name))
        else:
            dprint("Déjà présent !", priority_level=4)

        # On complète la feuille de primes
        dprint("Remplissage de la feuille des primes", priority_level=5)
        creation_feuille_primes_reporting(agency_name)

        # Création du reporting
        dprint("Vérification de la présence ou non du fichier de reporting", priority_level=4)
        if not os.path.isfile(get_agency_path("reporting", agency_name)):
            dprint("Non présent, création du fichier lancée", priority_level=4)
            creation_reporting(agency_name)
        else:
            dprint("Déjà présent !", priority_level=4)
