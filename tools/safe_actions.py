def safe_dict_get(data, list_keys, fail_result=None):
    """
    Permet d'accéder à la valeur d'un dictionnaire composé: valuer accéssible par
    la succession de clés.
    Fonctionne également pour les listes.
    En cas de clé incorrecte retourne None
    :param data:
    :param list_keys:
    :param fail_result:
    :return: valeur ciblée
    """
    target_value = data
    for key in list_keys:

        if isinstance(key, (str, int)):
            try:
                target_value = target_value[key]
            except:
                return fail_result

    return target_value


def safe_date_convert(date_str):
    """
    Permet de convertir un string() en date.
    En cas d'erreur de conversion retourne None
    :param date_str:
    :return:
    """
    formated_date = None

    from datetime import datetime
    # Format boond = 2022-07-07T09:52:57+0200
    if date_str is not None:
        split_date = date_str.split("T")
        try:
            formated_date = datetime.strptime(split_date[0], "%Y-%m-%d")
        except:
            pass

    return formated_date


def dprint(str_to_print, priority_level=1, preprint="", hashtag_display=True):
    """
    Fonction de print, dépend du paramètre global de l'app DEBUG
    Inclus une fonctionnalité de priorité de print dépendant du paramètre global de l'app PRIORITY_DEBUG_LEVEL
    ainsi que d'une indentation représentée par des "-" liée à la priorité de print
    :param str_to_print:
    :param priority_level:
    :param preprint:
    :param hashtag_display:
    :return:
    """
    from configuration import APP_CONFIG
    if APP_CONFIG.DEBUG and APP_CONFIG.PRIORITY_DEBUG_LEVEL >= priority_level:
        str_ident = "".join("-" for _ in range(priority_level))
        if hashtag_display:
            print(f"{preprint}#{str_ident} {str_to_print}")
        else:
            print(f"{preprint}{str_to_print}")


def taux(numerateur, denominateur):
    """
    Permet de calculer un taux sans erreur lors
    des divisions par 0, retourne 0 pour un div par 0
    :param numerateur:
    :param denominateur:
    :return: taux calculé
    """
    if denominateur != 0:
        return (numerateur / denominateur)
    else:
        return 0


def division(numerateur, denominateur):
    """
    Permet de calculer une division sans erreur lors
    des divisions par 0, retourne None pour un div par 0
    :param numerateur:
    :param denominateur:
    :return: quotient
    """
    if denominateur != 0:
        return numerateur / denominateur
    else:
        None


def copy_template_file(to_copy, to_paste):
    """
    Permet de copier coller un fichier d'un endroit
    à l'autre tout en le renommant
    :param to_copy:
    :param to_paste:
    :return:
    """
    import os.path
    import shutil

    if not os.path.isfile(to_paste):
        shutil.copy(
            to_copy,
            to_paste
        )


def get_tableau_size(sheet, dim):
    """
    Permet de récupérer les dimensions d'un tableau.
    Fonctionne uniquement pour les feuilles excel qui ne comporte qu'un tableau
    Retourne colonne de début / de fin | ligne de début / de fin, en fonction du paramètre
    :param sheet:
    :param dim:
    :return: dimension demandée (hauteur ou largeur)
    """

    def _get_ligne(sheet, ligne, cols):
        return [sheet[f"{col}{ligne}"].value for col in cols]

    letters = [chr(x + 65) for x in range(26)]
    numbers = [str(x) for x in range(10)]
    dims = {"col": None, "ligne": None}
    dims["col"] = (''.join(x for x in sheet.dimensions if x not in numbers)).split(":")

    # Début du tableau
    start_ligne = 1
    while True:
        if any(_get_ligne(sheet, start_ligne,
                          letters[letters.index(dims["col"][0]):letters.index(dims["col"][1]) + 1])):
            break
        start_ligne += 1

    # Fin du tableau
    end_ligne = start_ligne
    while True:
        if not any(
                _get_ligne(sheet, end_ligne, letters[letters.index(dims["col"][0]):letters.index(dims["col"][1]) + 1])):
            break
        end_ligne += 1

    dims["ligne"] = [start_ligne, end_ligne]

    return dims[dim]


def get_ligne_values(sheet, inspect_ligne, titles_ligne, start_col, end_col):
    """
    Permet de récupérer sous forme de dict() une ligne
    d'un tableau.
    Résutlat: {"titre de la colonne", "valeur à la ligne d'inspection", ...}
    :param sheet:
    :param inspect_ligne:
    :param titles_ligne:
    :param start_col:
    :param end_col:
    :return: ligne sous forme de dict()
    """
    ligne_data = dict()
    if isinstance(start_col, str):
        start_col = ord(start_col.lower()) - 96
    if isinstance(end_col, str):
        end_col = ord(end_col.lower()) - 96

    for col in range(start_col, end_col + 1):
        ligne_data[sheet.cell(row=titles_ligne, column=col).value] = sheet.cell(row=inspect_ligne, column=col).value

    return ligne_data


def save_file(path, data):
    """
    Permet de enregistrer un fichier.
    A utilser pour enregistrer des instances python du type dict() / list()
    :param path:
    :param data:
    :return:
    """
    with open(path, mode="w", encoding="utf-8") as file:
        file.write(
            str(data).replace("'", '"').replace("None", "null").replace("False", "false").replace("True", "true"))
