# Librairies
import os.path

# Modules / Dépendances
from configuration import APP_CONFIG


def get_template_path(type):
    """
    Permet de récupérer les path de tous les templates
    :param type:
    :return: path
    """
    path = None
    if type == "primes":
        if APP_CONFIG.MODE_PRIMES == "tableau":
            path = os.path.join(APP_CONFIG.BASE_DIR, "templates", APP_CONFIG.NOMS_TEMPLATES["primes_tableau"])
        elif APP_CONFIG.MODE_PRIMES == "drag and drop":
            path = os.path.join(APP_CONFIG.BASE_DIR, "templates", APP_CONFIG.NOMS_TEMPLATES["primes_drag_and_drop"])

    elif type == "primes_reporting":
        path = os.path.join(APP_CONFIG.BASE_DIR, "templates", APP_CONFIG.NOMS_TEMPLATES["primes_reporting"])

    elif type == "BP":
        path = os.path.join(APP_CONFIG.BASE_DIR, "templates", APP_CONFIG.NOMS_TEMPLATES["BP"])

    elif type == "noms_modeles":
        path = os.path.join(APP_CONFIG.BASE_DIR, "templates", APP_CONFIG.NOMS_TEMPLATES["noms_modeles"])

    elif type == "consultant_forfait":
        path = os.path.join(APP_CONFIG.BASE_DIR, "templates", APP_CONFIG.NOMS_TEMPLATES["consultant_forfait"])

    elif type == "reporting":
        path = os.path.join(APP_CONFIG.BASE_DIR, "templates", APP_CONFIG.NOMS_TEMPLATES["reporting"])

    return path


def get_general_path(type):
    """
    Permet de récupérer les path de tous les fichiers/dossiers généraux
    :param type:
    :return: path
    """
    from tools.date_info import get_year, get_month
    path = None
    nom_dossier_general = (APP_CONFIG.NOM_DOSSIER).replace("YYYY", get_year())

    if type == "primes":
        if APP_CONFIG.MODE_PRIMES == "tableau":
            path = os.path.join(
                APP_CONFIG.BASE_DIR,
                nom_dossier_general,
                APP_CONFIG.NOMS_GENERAUX["primes_tableau"]
            )
        elif APP_CONFIG.MODE_PRIMES == "drag and drop":
            path = os.path.join(
                APP_CONFIG.BASE_DIR,
                nom_dossier_general,
                APP_CONFIG.NOMS_GENERAUX["dossier_primes_drag_and_drop"],
                APP_CONFIG.NOMS_GENERAUX["primes_drag_and_drop"].replace("MM", get_month())
            )

    elif type == "noms_modeles":
        path = os.path.join(
            APP_CONFIG.BASE_DIR,
            nom_dossier_general,
            APP_CONFIG.NOMS_GENERAUX["dossier_primes_drag_and_drop"],
            APP_CONFIG.NOMS_GENERAUX["noms_modeles_drag_and_drop"]
        )

    elif type == "database":
        path = os.path.join(
            APP_CONFIG.BASE_DIR,
            nom_dossier_general,
            APP_CONFIG.NOMS_GENERAUX["dossier_stockage"],
            (APP_CONFIG.NOMS_GENERAUX["database"]).replace("MM", get_month())
        )

    elif type == "ressources":
        path = os.path.join(
            APP_CONFIG.BASE_DIR,
            nom_dossier_general,
            APP_CONFIG.NOMS_GENERAUX["dossier_stockage"],
            (APP_CONFIG.NOMS_GENERAUX["ressources"]).replace("MM", get_month()),
        )

    elif type == "consultant_forfait":
        path = os.path.join(
            APP_CONFIG.BASE_DIR,
            nom_dossier_general,
            APP_CONFIG.NOMS_GENERAUX["consultant_forfait"]
        )

    elif type == "dossier":
        path = os.path.join(
            APP_CONFIG.BASE_DIR,
            nom_dossier_general
        )
    return path


def get_agency_path(type, agency_name):
    """
    Permet de récupérer les path de tous les fichiers liés à l'agence entrée en paramètre
    :param type:
    :param agency_name:
    :return: path
    """
    from tools.date_info import get_year, get_month
    path = None
    nom_dossier_general = (APP_CONFIG.NOM_DOSSIER).replace("YYYY", get_year())

    if type == "primes":
        path = os.path.join(
            APP_CONFIG.BASE_DIR,
            nom_dossier_general,
            agency_name,
            APP_CONFIG.NOMS_AGENCES["dossier_reporting"].replace("MM", get_month()),
            APP_CONFIG.NOMS_AGENCES["primes"]
        )

    elif type == "consultant_forfait":
        path = os.path.join(
            APP_CONFIG.BASE_DIR,
            nom_dossier_general,
            agency_name,
            APP_CONFIG.NOMS_AGENCES["dossier_reporting"].replace("MM", get_month()),
            APP_CONFIG.NOMS_AGENCES["consultant_forfait"]
        )

    elif type == "database":
        path = os.path.join(
            APP_CONFIG.BASE_DIR,
            nom_dossier_general,
            agency_name,
            APP_CONFIG.NOMS_AGENCES["dossier_stockage"],
            APP_CONFIG.NOMS_AGENCES["database"].replace("MM", get_month())
        )

    elif type == "ressources":
        path = os.path.join(
            APP_CONFIG.BASE_DIR,
            nom_dossier_general,
            agency_name,
            APP_CONFIG.NOMS_AGENCES["dossier_stockage"],
            APP_CONFIG.NOMS_AGENCES["ressources"].replace("MM", get_month())
        )

    elif type == "reporting":
        path = os.path.join(
            APP_CONFIG.BASE_DIR,
            nom_dossier_general,
            agency_name,
            APP_CONFIG.NOMS_AGENCES["dossier_reporting"].replace("MM", get_month()),
            APP_CONFIG.NOMS_AGENCES["reporting"]
        )

    elif type == "previous_reporting":
        if get_month() != "01":
            previous_month = int(get_month()) - 1
            if previous_month < 10:
                previous_month = f"0{previous_month}"
            else:
                previous_month = str(previous_month)

            path = os.path.join(
                APP_CONFIG.BASE_DIR,
                nom_dossier_general,
                agency_name,
                APP_CONFIG.NOMS_AGENCES["dossier_reporting"].replace("MM", previous_month),
                APP_CONFIG.NOMS_AGENCES["reporting"]
            )

    elif type == "BP":
        path = os.path.join(
            APP_CONFIG.BASE_DIR,
            nom_dossier_general,
            agency_name,
            APP_CONFIG.NOMS_AGENCES["BP"]
        )

    elif type == "dossier":
        path = os.path.join(
            APP_CONFIG.BASE_DIR,
            nom_dossier_general,
            agency_name
        )

    elif type == "dossier_reporting":
        path = os.path.join(
            APP_CONFIG.BASE_DIR,
            nom_dossier_general,
            agency_name,
            APP_CONFIG.NOMS_AGENCES["dossier_reporting"].replace("MM", get_month())
        )

    return path
