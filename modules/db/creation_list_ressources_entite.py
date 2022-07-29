# Modules / Dépendances
# Tools
from tools.find_path import get_agency_path
from tools.read_json import get_dict_from_json_file
from tools.requests_tools import get_list_of_agencies
from tools.safe_actions import dprint, safe_dict_get, save_file


def get_ressources_of_entite(database):
    """
    Création des listes de toutes les ressoures pour une entité,
    à itérer autant de fois qu'il y a d'entités
    :param db générale
    :return: liste des ressources
    """
    ressoures = list()
    for manager in database:
        ressoures.append(
            " ".join(safe_dict_get(manager, ["manager"]).values())
        )

        for interne in safe_dict_get(manager, ["internes"]):
            ressoures.append(
                " ".join([safe_dict_get(interne, ["nom"]), safe_dict_get(interne, ["prenom"])])
            )

        for consultant in safe_dict_get(manager, ["consultants"]):
            ressoures.append(
                " ".join([safe_dict_get(consultant, ["nom"]), safe_dict_get(consultant, ["prenom"])])
            )
    return ressoures


def creation_list_ressources_all_entite():
    """
    Création des listes de toutes les ressoures pour chaque entité
    :return:
    """
    # Finance est une fusion
    for agency_name, list_agency_ids in get_list_of_agencies().items():
        dprint(f"Création de la liste de toutes les ressources de l'entité: {agency_name}", priority_level=4)
        # On récupère la base de données de l'agence
        database = get_dict_from_json_file(get_agency_path("database", agency_name))
        ressources = {"ressources": get_ressources_of_entite(database)}

        # Sauvegarde de la liste de ressources dans le dossier de l'entité
        save_file(get_agency_path("ressources", agency_name), ressources)
