# Librairies
import json

# Modules / Dépendances
from tools.exceptions import MissingFilePathException, IncorrectFileExtension


def get_dict_from_json_file(file_path: str = None):
    """
    This function can be used to get json file content into a dictionary.
    It would raise a MissingFilePathException Exception if the 'file_path'
    param is not filled.
    It would raise an IncorrectFileExtension Exception if the 'file_path'
    param does not refer to a json file.
    It would raise a FileNotFoundError Exception if the 'file_path' param
    does not refer to an existing file.
    :param file_path: The path to the json file to open including the file name
    :return: json file content in dict
    """
    if not file_path:
        raise MissingFilePathException('filepath is not filled')

    if not file_path.endswith(".json"):
        raise IncorrectFileExtension('Incorrect file format', file_path)

    with open(file_path, encoding="UTF-8") as file:
        data = json.load(file)

    return data
