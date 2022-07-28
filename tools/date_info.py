from calendar import monthrange
from datetime import datetime

from dateutil.relativedelta import relativedelta
from jours_feries_france import JoursFeries


def get_year():
    """
    Permet de récupérer l'année du reporting
    Si on est en décembre 2022 -> renvoie 2021
    :return: report year
    """
    year = datetime.today().strftime('%Y')
    # Reporting de décembre sera effectué en janvier de l'année suivante
    if datetime.today().strftime('%m') == "01":
        year = str(int(datetime.today().strftime('%Y')) - 1)

    return year


def get_month():
    """
    Permet de récupérer le mois du reporting
    Si on est en décembre 2022 -> renvoie janvier
    :return: mois du report
    """
    from configuration import APP_CONFIG
    current_date = datetime.today()
    # On récupère le mois précédent
    date_previous_month = current_date - relativedelta(months=APP_CONFIG.DELTA_MONTH)
    return date_previous_month.strftime('%m')


def get_max_day_of_month(year, month):
    """
    Permet de récupérer le nombre de jour max d'un mois,
    on peut alors avoir les deux jours extrêmes d'un mois: 01:max_day
    :param year:
    :param month:
    :return: nombre de jours maximun du mois
    """
    return str(monthrange(int(year), int(month))[1])


def get_previous_month_dates():
    """
    Permet de récupérer les jours extrêmes du mois du reporting
    :return: jours extrêmes du mois
    """
    # On récupère le début et la fin du mois précédent
    # Année
    year = get_year()
    # Mois précédent
    previous_month = get_month()
    # Max de jour de ce mois
    end_previous_month = get_max_day_of_month(get_year(), previous_month)
    # On assemble les infos pour avoir le debut et le fin du mois précédent
    return [
        "-".join([year, previous_month, "01"]),
        "-".join([year, previous_month, end_previous_month])
    ]


def get_nb_jours_ouvrables(year, month):
    """
    Permet de récupérer le nombre de jours ouvrables dans un mois
    calcul très précis: prend en compte les weekend ET es jours fériés
    liés aux fêtes. Ce calcul est adapté pour chaque années et prend en
    compte le changements de dates des fêtes françaises

    :param year:
    :param month:
    :return: nombre de jours ouvrables
    """
    # On compte le nombre de jours fériées lié à des fêtes / moments historiques
    nb_jours_feries = 0
    for date in JoursFeries.for_year(int(year)).values():
        if date.strftime('%m') == get_month():
            nb_jours_feries += 1

    # Récupère le nombre de jours du mois
    max_day = get_max_day_of_month(year, month)

    # On compte le nombre de jours non ouvrables (Samedi = 5 et Dimanche = 6)
    nb_jours_non_ouvrable = 0
    for day in range(1, int(max_day) + 1):
        day_id = datetime(int(year), int(month), day).weekday()
        if day_id == 5 or day_id == 6:
            nb_jours_non_ouvrable += 1

    # Nb de jours ouvrables = nb de jours dans le mois - nb de jours fériés - nb de jours non ouvrables
    return int(max_day) - nb_jours_feries - nb_jours_non_ouvrable
