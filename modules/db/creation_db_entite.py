# Modules / Dépendances
# Tools
from tools.find_path import get_agency_path
from tools.requests_tools import get_list_of_agencies
from tools.safe_actions import dprint, save_file, safe_dict_get
# Modules
from modules.db.creation_db_entite_get_complement_infos import get_general_project_info, get_ressource_info
from modules.db.creation_db_entite_delete_doublons import delete_doublons


def creation_db_of_entite(genaral_database, list_agency_ids):
    """
    Création de la db d'une entité, fonction à itérer autant de fois qu'il y a d'entités
    :param genaral_database:
    :param list_agency_ids:
    :return: db de l'entité
    """
    entity_db = []
    for project in genaral_database["database"]:
        # On vérifie que le projet appartient à l'entité pour laquelle on crée la db
        if project["manager agency id"] in list_agency_ids:
            # On récupère les infos générales du projet
            general_project_info = get_general_project_info(project)
            dprint(f"Récupération des infos générales du projet: {safe_dict_get(project, ['data', 0, 'attributes', 'value'])}", priority_level=5)

            # On itère sur chaque ressource qui travaille sur le projet:
            # Les consultants apparaissent à partir de la ligne 7 de 'data'
            for ressource in project["data"][7:]:
                # On récupère toutes les infos utiles sur la ressource
                ressource_info = get_ressource_info(general_project_info, ressource)
                dprint(f"Récupération des informations détaillées de la ressource: "
                       f"{safe_dict_get(ressource, ['attributes', 'scorecard', 'resource', 'lastName'])} "
                       f"{safe_dict_get(ressource, ['attributes', 'scorecard', 'resource', 'firstName'])}",
                       priority_level=6)

                # On trie le type de ressource: consultant (en mission) | interne (pas de mission)
                if bool(ressource_info["consultant"]):
                    general_project_info["consultants"].append(ressource_info["consultant"])

                if bool(ressource_info["interne"]):
                    general_project_info["internes"].append(ressource_info["interne"])

            entity_db.append(general_project_info)

    return entity_db


def creation_db_pour_toutes_les_entites(genaral_database):
    """
    Création de la db de chaque entité
    :param genaral_database:
    :return:
    """
    # Finance est une fusion donc on utlise une liste finance in [1, 3] les ids de FS et CS
    for agency_name, list_agency_ids in get_list_of_agencies().items():
        dprint(f"Création de la db de l'agence: {agency_name}", priority_level=4)
        entite_db = creation_db_of_entite(genaral_database, list_agency_ids)
        final_entite_db = delete_doublons(entite_db)
        save_file(get_agency_path("database", agency_name), final_entite_db)



