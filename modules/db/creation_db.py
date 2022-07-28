# Modules / Dépendances
# Tools
from tools.safe_actions import dprint
# Modules
from modules.db.creation_db_genarale import create_database_with_parameters
from modules.db.creation_list_all_ressources import creation_list_all_ressources
from modules.db.creation_db_entite import creation_db_pour_toutes_les_entites
from modules.db.creation_list_ressources_entite import creation_list_ressources_all_entite


def creation_db(**filters):
    """
    Création de toutes les bases de données:
    - db générale
    - liste ressources générale
    - db par entité
    - liste ressources par entité
    :param filters:
    :return:
    """

    dprint("Création des db:", priority_level=2)

    # Création de la db générales en fonctions des paramètres transmis
    dprint("Création de la db génarale", priority_level=3)
    general_database = create_database_with_parameters(**filters)

    # Création de la liste contenant toutes les ressources
    dprint("Création de la liste de toutes les ressources", priority_level=3)
    creation_list_all_ressources(general_database)

    # Création des db de chaque entite
    dprint("Création des db pour chaque entité", priority_level=3)
    creation_db_pour_toutes_les_entites(general_database)

    # Création des listes de ressources de chaque entite
    dprint("Création de la liste de toutes les ressources pour chaque entité", priority_level=3)
    creation_list_ressources_all_entite()
