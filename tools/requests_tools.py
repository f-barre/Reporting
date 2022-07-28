# Librairies
import json
from time import sleep
import requests
from requests.auth import HTTPBasicAuth

# Modules / Dépendances
from configuration import APP_CONFIG
from tools.safe_actions import dprint


def request(endpoint, **query_params):
    """
    Permet de faire une requête à l'API de BoondManager
    tout en y passant autant de paramètre que l'on souhaite
    :param endpoint:
    :param query_params:
    :return: reponse de la requête au format dict()
    """
    param = endpoint

    if bool(query_params):
        param += "?"
        for key, value in query_params.items():
            param += ("&" + key + "=" + str(value))
    dprint(f"Endpoint requested: {param}", priority_level=10)

    response = None
    while True:
        try:
            response = requests.get(APP_CONFIG.BOONDMANAGER_API_URL + param, auth=HTTPBasicAuth(APP_CONFIG.BOONDMANAGER_API_LOGIN, APP_CONFIG.BOONDMANAGER_API_PASSWORD))
            if response.ok:
                response = json.loads(response.text)
                break
            else:
                print(f"ERROR, Code erreur de la requete: {response}")
        except:
            sleep(5)
            response = requests.get(APP_CONFIG.BOONDMANAGER_API_URL + param, auth=HTTPBasicAuth(APP_CONFIG.BOONDMANAGER_API_LOGIN, APP_CONFIG.BOONDMANAGER_API_PASSWORD))
            if response.ok:
                response = json.loads(response.text)
                break


    return response


def get_list_of_agencies():
    """
    Permet de récupérer la liste des agences Lamarck
    :return: liste des agences Lamarack
    """
    from configuration import APP_CONFIG
    response_json = request("/agencies")
    # liste brute: agence["attributes"]["name"] for agence in response_json["data"]
    # PB: FS et CS à merge
    # ID des entités = ordre dans la liste +1

    agencies = dict()
    for agence_data in response_json["data"]:
        agence_name = agence_data["attributes"]["name"]
        agence_id = agence_data["id"]

        if agence_name in APP_CONFIG.AGENCES:
            # Merge CS et FC -> Finance
            if agence_name in ["Lamarck FS", "Lamarck CS"]:
                if agencies.get("Lamarck Finance", "no") == "no":
                    agencies["Lamarck Finance"] = list()
                agencies["Lamarck Finance"].append(agence_id)
            else:
                agencies[agence_name] = [agence_id]

    return agencies