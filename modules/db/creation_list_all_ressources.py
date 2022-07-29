# Modules / Dépendances
# Tools
from tools.find_path import get_general_path
from tools.safe_actions import dprint, safe_dict_get, save_file


def get_all_ressources(database):
    """
    Récupère toutes les ressources du mois du reporting dans labase de données générale
    :param db générales:
    :return: liste de toutes les ressources
    """
    list_of_all_ressources = list()
    # On récupère tous les consultants de la db
    for page, projet in enumerate(database["database"]):
        dprint(f"Page: {page}", priority_level=4)
        for ressource in projet["data"][7:]:
            name = f"{safe_dict_get(ressource, ['attributes', 'scorecard', 'resource', 'lastName'])} " \
                   f"{safe_dict_get(ressource, ['attributes', 'scorecard', 'resource', 'firstName'])}"
            list_of_all_ressources.append(name)

    # On supprime les doublons
    list_of_all_ressources_without_double = []
    for i in list_of_all_ressources:
        if i not in list_of_all_ressources_without_double:
            list_of_all_ressources_without_double.append(i)

    return list_of_all_ressources_without_double


def creation_list_all_ressources(database):
    """
    Création de la liste générales des ressources, liste contenant toutes les ressources du mois du reporting
    :param database:
    :return: liste de toutes les ressources
    """
    dprint("Création de la liste de toutes les ressources", priority_level=3)
    all_ressources = {"ressources": get_all_ressources(database)}
    save_file(get_general_path("ressources"), all_ressources)

    return all_ressources
