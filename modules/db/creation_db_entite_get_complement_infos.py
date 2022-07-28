# Librairies
from datetime import datetime

# Modules / Dépendances
# Tools
from tools.date_info import get_previous_month_dates
from tools.safe_actions import safe_dict_get
from tools.requests_tools import request

def get_general_project_info(project):
    return {
        "extract date": datetime.today().strftime('%d/%m/%Y'),
        "client": [{"nom": safe_dict_get(project, ["data", 0, "attributes", "value"], fail_result="").replace("'", " ")}],
        "manager": {
            "nom": safe_dict_get(project, ["data", 1, "attributes", "value"]).split(" ")[0],
            "prenom": safe_dict_get(project, ["data", 1, "attributes", "value"]).split(" ")[1]
        },
        "debut": [safe_dict_get(project, ["data", 2, "attributes", "value"])],
        "fin": [safe_dict_get(project, ["data", 3, "attributes", "value"])],
        "ca facture": float(safe_dict_get(project, ["data", 4, "attributes", "value"], fail_result=float(0))),
        "cout de prod": float(safe_dict_get(project, ["data", 5, "attributes", "value"], fail_result=float(0))),
        "ca prod": float(safe_dict_get(project, ["data", 6, "attributes", "value"], fail_result=float(0))),
        "consultants": [],
        "internes": [],
        "equipes": []
    }


def get_ressource_info(projet, ressource):
    def get_temps_interne(ressource):
        previous_month_dates = get_previous_month_dates()
        # On récupère tous les temps de la ressource sur le mois précédent
        temps = request(
            f"/reporting-resources/",
            resources=safe_dict_get(ressource, ['attributes', 'scorecard', 'resource', 'id']),
            startDate=previous_month_dates[0],
            endDate=previous_month_dates[1],
            period="onePeriod"
        )

        # Calcul du nombre de jours en interne
        nb_de_jours_en_interne = float(0)
        for t in temps.get("data", []):
            if safe_dict_get(t, ["attributes", "scorecard", "project"]) is None:
                nb_de_jours_en_interne += float(safe_dict_get(t, ["attributes", "value"]))

        return nb_de_jours_en_interne

    def get_ca_total_de_prod(ressource_info, last_prestation):
        # CA total de prod (nb de jours facturés * TJM)
        # régies et hors forfait il faut prendre le forfait et le divisé par 12 pour avoir son salaire lissé
        # pour les autres aller dans prestation prendre la dernière et regarder TARIF PROJET JOURNALIER * temps consommé
        # ancien emplacement donné (pas tjrs complété) consultant_information["data"]["attributes"]["averageDailyPriceExcludingTax"]
        ca_total_de_prod = float(0)

        tpj = safe_dict_get(last_prestation, ["data", "attributes", "averageDailyPriceExcludingTax"])
        if tpj is not None:
            ca_total_de_prod = float(tpj) * float(ressource_info["nb de jours en mission"])

        return ca_total_de_prod

    def get_cout_de_prod(ressource_info, last_prestation):
        cout_de_prod = float(0)

        cjm = safe_dict_get(last_prestation, ["data", "attributes", "averageDailyContractCost"])
        if cjm is not None:
            cout_de_prod = float(cjm) * float(ressource_info["nb de jours en mission"])

        return cout_de_prod

    ressource_info = {
            "nom": safe_dict_get(ressource, ['attributes', 'scorecard', 'resource', 'lastName']),
            "prenom": safe_dict_get(ressource, ['attributes', 'scorecard', 'resource', 'firstName']),
            "nb de jours en mission": float(safe_dict_get(ressource, ['attributes', 'value'], fail_result=float(0))),
            "fin": safe_dict_get(projet, ["fin", 0]),
            "ca facture": safe_dict_get(projet, ["ca facture"])
        }

    # Pour avoir toutes les infos nécessaires, on fait plusieurs requêtes complémentaires sur la ressource
    information = request(f"/resources/{ressource['attributes']['scorecard']['resource']['id']}/information")
    administrative = request(f"/resources/{ressource['attributes']['scorecard']['resource']['id']}/administrative")
    prestations = request(f"/resources/{ressource['attributes']['scorecard']['resource']['id']}/deliveries-inactivities")
    last_prestation = request(f"/deliveries/{prestations['data'][0]['id']}")

    # Temps en internes
    ressource_info["nb de jours en interne"] = get_temps_interne(ressource)

    # Fonction
    ressource_info["fonction"] = safe_dict_get(information, ["data", "attributes", "title"], fail_result="")

    # Statut (0 = salarié / 1 = indépendant)
    statut_list = ["Salarié", "Indépendant"]
    if safe_dict_get(information, ["data", "attributes", "typeOf"]) is not None:
        ressource_info["statut"] = safe_dict_get(statut_list, [int(information["data"]["attributes"]["typeOf"])])
    else:
        ressource_info["statut"] = ""

    # CA total de prod (nb de jours facturés * TJM)
    ressource_info["ca total de prod"] = get_ca_total_de_prod(ressource_info, last_prestation)

    # Coût de prod (CJM * nb de jours)
    ressource_info["cout de prod"] = get_cout_de_prod(ressource_info, last_prestation)

    # Salaire (besoin pour les internes et le calcul du montant mensuel chargé)
    ressource_info["salaire"] = safe_dict_get(administrative, ["included", 1, "attributes", "monthlySalary"], fail_result=0)

    # TODO: à supprimer ?
    # Variable*2 où est-ce ??
    # Marge de prod (Ca total - cout de prod - variable) à calculer dans le repoting maker
    # Rentabilité (marge de prod / cout de prod) Utile ? pas dans le reporting ...

    # Trie consultant en mission ou non (en mission = consultant / pas de mission = interne)
    data_to_return = {"consultant": dict(), "interne": dict()}

    if ressource_info["nb de jours en interne"] != float(0):
        data_to_return["interne"] = ressource_info

    if ressource_info["nb de jours en mission"] != float(0):
        data_to_return["consultant"] = ressource_info

    return data_to_return
