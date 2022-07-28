import json
import os

BASE_DIR = os.path.dirname(os.path.realpath(__file__))
SECRET_FILE_PATH = 'secrets.json'
SECRET_CONFIG_STORE = {}

try:
    with open(SECRET_FILE_PATH) as config:
        SECRET_CONFIG_STORE = json.load(config)
except FileNotFoundError:
    try:
        with open(os.path.join(BASE_DIR, 'secrets.json')) as config:
            SECRET_CONFIG_STORE = json.load(config)
    except Exception:
        raise


class BaseConfig(object):
    APPLICATION_NAME = 'Reporting'
    DEBUG = 1
    BASE_DIR = BASE_DIR


class ProductionConfig(BaseConfig):
    ENV = "production"
    LISTE_PRIMES_A_COMPTER = [

    ]
    DEBUG = 1
    PRIORITY_DEBUG_LEVEL = 100
    GUI_PROPERTIES = {
        "bg": "black",
        "size": ["800", "500"],
        "title": "Menu reporting",
        "nb btns": 5,
        "space between btn": 10
    }
    BOONDMANAGER_API_URL = SECRET_CONFIG_STORE["boondManager_api_url"]
    BOONDMANAGER_API_LOGIN = SECRET_CONFIG_STORE["boondManager_api_login"]
    BOONDMANAGER_API_PASSWORD = SECRET_CONFIG_STORE["boondManager_api_password"]


class LocalConfig(BaseConfig):
    ENV = "local"
    MODE_PRIMES = "tableau"  # "drag and drop" | "tableau"
    DELTA_MONTH = 1
    AGENCES = [
        "Keystone",
        "Lamarck FS",
        "Lamarck CS",
        "Lamarck Group",
        "Lamarck Solutions"
    ]
    NOMS_TEMPLATES = {
        "primes_tableau": "template_feuille_de_primes_tableau.xlsx",
        "primes_drag_and_drop": "template_feuille_de_primes_drag_and_drop.xlsx",
        "primes_reporting": "template_feuille_de_primes_reporting.xlsx",
        "BP": "template_BP.xlsx",
        "noms_modeles": "template_noms_modeles.txt",
        "consultant_forfait": "template_consultants_forfait.xlsx",
        "reporting": "template_reporting.xlsx"
    }
    NOM_DOSSIER = "Business_Plan_et_Reporting_YYYY"
    NOMS_GENERAUX = {
        "primes_tableau": "Primes_consultants.xlsx",
        "consultant_forfait": "Consultants_payés_au_forfait.xlsx",
        "dossier_primes_drag_and_drop": "Primes",
        "primes_drag_and_drop": "MM.xlsx",
        "noms_modeles_drag_and_drop": "Titres_fichiers_règles_à_suivre.txt",
        "dossier_stockage": "autres",
        "database": "MM_database.json",
        "ressources": "MM_all_ressources.json"
    }

    NOMS_AGENCES = {
        "BP": "BP.xlsx",
        "dossier_reporting": "MM",
        "consultant_forfait": "Consultants_payés_au_forfait.xlsx",
        "primes": "Primes_Consultants.xlsx",
        "reporting": "reporting.xlsx",
        "dossier_stockage": "autres",
        "database": "MM_database.json",
        "ressources": "MM_all_ressources.json"
    }

    LISTE_PRIMES_A_COMPTER = [
        "Prime d'intervention en propre",
        "Régularisation prime d'intervention en propre",
        "Prime de management",
        "Prime vacances"
    ]

    DEBUG = 1
    PRIORITY_DEBUG_LEVEL = 100
    FOLDER_NAME = "Business_Plan_et_Reporting"
    GUI_PROPERTIES = {
        "bg": "black",
        "size": ["800", "500"],
        "title": "Menu reporting",
        "nb btns": 4,
        "space between btn": 10
    }
    BOONDMANAGER_API_URL = SECRET_CONFIG_STORE["boondManager_api_url"]
    BOONDMANAGER_API_LOGIN = SECRET_CONFIG_STORE["boondManager_api_login"]
    BOONDMANAGER_API_PASSWORD = SECRET_CONFIG_STORE["boondManager_api_password"]


class ConfigurationException(Exception):
    pass


class Configuration(dict):
    def __init__(self, *args, **kwargs):
        if not os.environ.get('ENV'):
            raise ConfigurationException(
                "Please set 'ENV' environment variable"
            )

        super(Configuration, self).__init__(*args, **kwargs)

        self["ENV"] = os.environ['ENV']
        self.__dict__ = self

    def from_object(self, obj):
        for attr in dir(obj):

            if not attr.isupper():
                continue

            self[attr] = getattr(obj, attr)

        self.__dict__ = self


APP_CONFIG = Configuration()

if APP_CONFIG.get('ENV') == 'production':
    APP_CONFIG.from_object(ProductionConfig)
elif APP_CONFIG.get('ENV') == 'local':
    APP_CONFIG.from_object(LocalConfig)
else:
    APP_CONFIG.from_object(ConfigurationException)
