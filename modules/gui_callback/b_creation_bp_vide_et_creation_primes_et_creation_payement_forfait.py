# Modules / Dépendances
from configuration import APP_CONFIG
# Tools
from tools.safe_actions import dprint
# Modules
from modules.bp.creation_bp_vide import creation_bp_vide
from modules.feuille_exception_forfait.creation_feuille_exception_forfait import creation_feuille_exception_forfait
from modules.feuille_prime.creation_feuille_primes import creation_feuilles_primes

def b_creation_bp_vide_et_creation_primes_et_creation_payement_forfait():
    """
    Fontion d'appel de la création
    - des BP vides
    - Feuilles de primes vides
    - Feuilles des consultants payés au forfait vides
    Fonctionnalités:
    - Si la feuille BP est déjà créée -> pas d'écrasement, la feuille sera conservée
    - Pareil pour la feuille de primes
    - Pareil pour la feuille des consultants payés au forfait

    - Si des nouveaux consultants apparaissent sur le mois,
    les feuilles de primes et de consultants forfaits sont mises à jour

    - Deux modes pour les primes:
        - 'drag and drop': un dossier sera crée dans lequell il suffit de déposer le fichier des primes du mois
        - 'tableau': un tableau est crée, il suffit de le compléter
    :return:
    """

    dprint("Btn B pressed: run b_creation_bp_vide_et_creation_primes_et_creation_payement_forfait", priority_level=1)
    dprint("Création des BP vides", priority_level=2)
    creation_bp_vide()
    dprint("Création des feuilles d'exceptions des consultants payés au forfait", priority_level=2)
    creation_feuille_exception_forfait()
    if APP_CONFIG.MODE_PRIMES == "tableau":
        dprint("Primes en mode 'tableau', création de la feuille des primes vide", priority_level=2)
        creation_feuilles_primes()
