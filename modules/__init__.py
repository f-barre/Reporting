# Librairies
import os.path

# Modules / Dépendances
from configuration import APP_CONFIG
from tools.date_info import get_year
from tools.requests_tools import get_list_of_agencies
from tools.safe_actions import copy_template_file
from tools.find_path import get_general_path, get_template_path

# Création de la structure du projet
# Dossier général
general_path = get_general_path("dossier")
if not os.path.isdir(general_path):
    os.makedirs(general_path)

# Dossier de stockage de général
stockage_path = os.path.join(general_path, APP_CONFIG.NOMS_GENERAUX["dossier_stockage"])
if not os.path.isdir(stockage_path):
    os.makedirs(stockage_path)

# Dossier de stockage des primes
if APP_CONFIG.MODE_PRIMES == "drag and drop":
    stockage_path = os.path.join(general_path, APP_CONFIG.NOMS_GENERAUX["dossier_primes_drag_and_drop"])
    if not os.path.isdir(stockage_path):
        os.makedirs(stockage_path)
        # On ajoute une fichier expliquant le format des fichiers à déposer
        copy_template_file(get_template_path("noms_modeles"), get_general_path("noms_modeles"))

# Un dossier par entité
for agence in get_list_of_agencies().keys():
    agence_path = os.path.join(general_path, agence)
    if not os.path.isdir(agence_path):
        os.makedirs(agence_path)

    # Dossier de stockage de dépendances
    stockage_path = os.path.join(agence_path, APP_CONFIG.NOMS_GENERAUX["dossier_stockage"])
    if not os.path.isdir(stockage_path):
        os.makedirs(stockage_path)