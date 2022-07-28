# Modules / Dépendances
# Tools
from tools.find_path import get_general_path
from tools.requests_tools import request
from tools.safe_actions import dprint, safe_dict_get, save_file


def create_database_with_parameters(**params):
    """
    Création de la db générale
    :param params:
    :return: db générale
    """
    database = {"database": []}
    page = 1

    # On récupère tous les projets du mois
    while True:
        # Params
        params["maxProjects"] = 1
        params["page"] = page
        # Request
        response_json = request("/reporting-projects", **params)
        dprint(f"Page: {page}", priority_level=4)

        if response_json and len(safe_dict_get(response_json, ["data"])) > 0:
            # On ajoute l'id de l'agence pour trier par entité les managers par la suite
            project_detail = request(f"/projects/{response_json['included'][2]['id']}")
            manager_details = request(
                f"/resources/{project_detail['data']['relationships']['mainManager']['data']['id']}")
            manager_agency_id = manager_details["data"]["relationships"]["agency"]["data"]["id"]

            response_json["manager agency id"] = manager_agency_id

            database["database"].append(response_json)
            page += 1
        else:
            break
    save_file(get_general_path("database"), database)

    return database
