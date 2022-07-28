# Modules / Dépendances
# Tools
from tools.date_info import get_year, get_month, get_max_day_of_month
from tools.safe_actions import dprint
# Modules
from modules.db.creation_db import creation_db


def a_creation_db():
    """
    Fontion d'appel de la création des bases de données
    :return:
    """
    dprint("Btn A pressed: run a_creation_db", priority_level=1)
    year = get_year()
    month = get_month()
    creation_db(
        startDate="-".join([year, month, "01"]),
        endDate="-".join([year, month, get_max_day_of_month(year, month)])
    )
