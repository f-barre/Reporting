# Librairies
import xlwings as xw

# Modules / Dépendances
from tools.find_path import get_agency_path


# Sous-Fonctions
def fill_kpis_mensuels(bp_sheet, manager_sheet, start_ligne, month):
    """
    Calcul des avancements dans le tableau KPIS MENSUELS
    :param bp_sheet:
    :param manager_sheet:
    :param start_ligne:
    :param month:
    :return:
    """
    # BP
    # CA total
    manager_sheet[f"E{start_ligne + 3}"].value = bp_sheet.cell(row=8, column=month + 2).value

    # Marge production total
    manager_sheet[f"E{start_ligne + 4}"].value = bp_sheet.cell(row=7, column=month + 2).value
    # Marge BU
    manager_sheet[f"E{start_ligne + 6}"].value = bp_sheet.cell(row=11, column=month + 2).value

    # Avancement
    # CA total
    manager_sheet[f"F{start_ligne + 3}"].value = f"=C{start_ligne + 3} - E{start_ligne + 3}"

    # Marge production total
    manager_sheet[f"F{start_ligne + 4}"].value = f"=C{start_ligne + 4} -  E{start_ligne + 4}"

    # Marge BU
    manager_sheet[f"F{start_ligne + 6}"].value = f"=C{start_ligne + 6} - E{start_ligne + 6}"


def fill_kpis_consolides(bp_sheet, manager_sheet, start_ligne, month, agency_name, manager_name):
    """
    Calcul des avancements dans le tableau KPIS CONSOLIDES
    :param bp_sheet:
    :param manager_sheet:
    :param start_ligne:
    :param month:
    :param agency_name:
    :param manager_name:
    :return:
    """

    def get_kpis_consolide_ligne(sheet):
        """
        Permet de récupérer la ligne de début du tableau des KPIS CONSOLIDES
        :param sheet:
        :return: ligne de début
        """
        ligne = 8
        while sheet[f"H{ligne}"].value != "CA":
            ligne += 1
        return ligne

    try:
        with xw.App(visible=False) as app:
            # On récupère le dernier reporting et la ligne de son tableau KPIS CONSOLIDES
            previous_reporting_workbook = xw.Book(get_agency_path("previous_reporting", agency_name))
            previous_reporting_sheet = previous_reporting_workbook.sheets[manager_name]
            previous_kpis_consolide_ligne = get_kpis_consolide_ligne(previous_reporting_sheet)

            # CA
            manager_sheet[f"J{start_ligne}"].value = bp_sheet.cell(row=8, column=month + 2).value + \
                                                     previous_reporting_sheet[f'I{previous_kpis_consolide_ligne}'].value
            # Marge de production
            manager_sheet[f"J{start_ligne + 1}"].value = bp_sheet.cell(row=7, column=month + 2).value + \
                                                         previous_reporting_sheet[
                                                             f'I{previous_kpis_consolide_ligne + 1}'].value
            # Cout interne
            manager_sheet[f"J{start_ligne + 3}"].value = bp_sheet.cell(row=6, column=month + 2).value + \
                                                         previous_reporting_sheet[
                                                             f'I{previous_kpis_consolide_ligne + 3}'].value
            # Marge BU
            manager_sheet[f"J{start_ligne + 4}"].value = bp_sheet.cell(row=11, column=month + 2).value + \
                                                         previous_reporting_sheet[
                                                             f'I{previous_kpis_consolide_ligne + 4}'].value

    except:
        # CA
        manager_sheet[f"J{start_ligne}"].value = bp_sheet.cell(row=8, column=month + 2).value
        # Marge de production
        manager_sheet[f"J{start_ligne + 1}"].value = bp_sheet.cell(row=7, column=month + 2).value
        # Cout interne
        manager_sheet[f"J{start_ligne + 3}"].value = bp_sheet.cell(row=6, column=month + 2).value
        # Marge BU
        manager_sheet[f"J{start_ligne + 4}"].value = bp_sheet.cell(row=11, column=month + 2).value

    # Taux marge production
    manager_sheet[f"I{start_ligne + 2}"].value = f"=J{start_ligne + 1} / J{start_ligne}"
    # Taux de marge BU
    manager_sheet[f"I{start_ligne + 5}"].value = f"=J{start_ligne + 4} / J{start_ligne}"


# Fonction principale
def calcul_avancement_manager(bp_workbook, reporting_workbook, month, agency_name):
    """
    Calcul des avancements dans les feuilles de managers
    :param bp_workbook:
    :param reporting_workbook:
    :param month:
    :param agency_name:
    :return:
    """

    for sheet_index in range(1, len(bp_workbook.worksheets[
                                    1:]) + 1):  # +1 car la première sheet c'est le suivi / l'overview
        # On récupère les sheet à modifier
        bp_sheet = bp_workbook.worksheets[sheet_index]
        manager_sheet = reporting_workbook.worksheets[sheet_index]

        # Get ligne début des tableaux
        start_ligne = 1
        while manager_sheet[f"B{start_ligne}"].value != "KPIS MENSUELS":
            start_ligne += 1

        # Fill KPIs Mensuels tableau
        fill_kpis_mensuels(bp_sheet, manager_sheet, start_ligne, month)

        # Fill KPIs Consolidé tableau
        fill_kpis_consolides(bp_sheet, manager_sheet, start_ligne, month, agency_name, manager_sheet.title)
