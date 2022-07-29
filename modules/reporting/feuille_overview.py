# Modules / Dépendances
# Tools
from tools.date_info import get_year, get_month
from tools.format_cell_tools import display_border, format_euro, format_pourcentage


# Sous-Fonctions
def fill_overview_table(sheet, ligne):
    """
    Remplissage du 1er tableau de la feuille d'overview
    :param sheet:
    :param ligne:
    :return:
    """
    sheet[f"C{ligne}"] = f"{get_month()}/{get_year()}"


def fill_kpis_mensuels(workbook, sheet, ligne):
    """
    Remplissage du tableau KPIS MENSUELS
    :param workbook:
    :param sheet:
    :param ligne:
    :return:
    """

    def get_ligne_nb_consultants(sheet_manager):
        """
        Permet de récupérer la ligne sur laquelle on retrouve le nombre de consultants
        ligne des totaux du tableau des consultants de la feuille manager
        :param sheet_manager:
        :return: ligne des totaux
        """
        ligne_nb_consultants = 1
        while str(sheet_manager[f"B{ligne_nb_consultants}"].value).lower() != "internes":
            ligne_nb_consultants += 1

        return ligne_nb_consultants - 2

    def get_ligne_tables_kpis(sheet_manager):
        """
        Permet de récupérer la ligne de départ du tableau KPIS MENSUELS
        :param sheet_manager:
        :return: ligne tableau KPIS MENSUELS
        """
        ligne_tables_kpis = 1
        while str(sheet_manager[f"B{ligne_tables_kpis}"].value).lower() != "kpis mensuels":
            ligne_tables_kpis += 1

        return ligne_tables_kpis

    # On itère les sheets manager (on saute la première sheet car c'est l'overview)
    for index, sheet_manager in enumerate(workbook.worksheets[1:]):
        # Nom manager
        sheet[f"B{ligne}"].value = sheet_manager.title

        # Nombre de consultants
        sheet[f"C{ligne}"].value = f"={sheet_manager.title}!B${get_ligne_nb_consultants(sheet_manager)}"

        # Infos suivantes sont à récupérer dans KPI Mensuels et Consolidé => cherhcons les lignes de ces tableaux
        ligne_tables_kpis = get_ligne_tables_kpis(sheet_manager)

        # CA Production
        sheet[f"D{ligne}"].value = f"={sheet_manager.title}!C${ligne_tables_kpis + 3}"
        format_euro(sheet[f"D{ligne}"])

        # Marge production
        sheet[f"E{ligne}"].value = f"={sheet_manager.title}!C${ligne_tables_kpis + 4}"
        format_euro(sheet[f"E{ligne}"])

        # %
        sheet[
            f"F{ligne}"].value = f"={sheet_manager.title}!C${ligne_tables_kpis + 4} / {sheet_manager.title}!C${ligne_tables_kpis + 3}"
        format_pourcentage(sheet[f"F{ligne}"])

        # Internes
        sheet[f"G{ligne}"].value = f"={sheet_manager.title}!C${ligne_tables_kpis + 5}"
        format_euro(sheet[f"G{ligne}"])

        # Marge BU
        sheet[f"H{ligne}"].value = f"={sheet_manager.title}!C${ligne_tables_kpis + 6}"
        format_euro(sheet[f"H{ligne}"])

        # %
        sheet[
            f"I{ligne}"].value = f"={sheet_manager.title}!C${ligne_tables_kpis + 6} / {sheet_manager.title}!C${ligne_tables_kpis + 3}"
        format_pourcentage(sheet[f"I{ligne}"])

        # Evite d'avoir une ligne vide
        if index is not len(workbook.worksheets[1:]) - 1:
            sheet.insert_rows(ligne)

    # Totaux
    # nb consultants
    sheet[
        f"C{ligne + len(workbook.worksheets[1:])}"].value = f"=SUM(C{ligne}:C{ligne + len(workbook.worksheets[1:]) - 1})"

    # CA prod
    sheet[
        f"D{ligne + len(workbook.worksheets[1:])}"].value = f"=SUM(D{ligne}:D{ligne + len(workbook.worksheets[1:]) - 1})"

    # Marge prod
    sheet[
        f"E{ligne + len(workbook.worksheets[1:])}"].value = f"=SUM(E{ligne}:E{ligne + len(workbook.worksheets[1:]) - 1})"

    # %
    sheet[
        f"F{ligne + len(workbook.worksheets[1:])}"].value = f"=E{ligne + len(workbook.worksheets[1:])} / D{ligne + len(workbook.worksheets[1:])}"
    format_pourcentage(sheet[f"F{ligne}"])

    # Internes
    sheet[
        f"G{ligne + len(workbook.worksheets[1:])}"].value = f"=SUM(G{ligne}:G{ligne + len(workbook.worksheets[1:]) - 1})"

    # Marge BU
    sheet[
        f"H{ligne + len(workbook.worksheets[1:])}"].value = f"=SUM(H{ligne}:H{ligne + len(workbook.worksheets[1:]) - 1})"

    # %
    sheet[
        f"I{ligne + len(workbook.worksheets[1:])}"].value = f"=H{ligne + len(workbook.worksheets[1:])} / D{ligne + len(workbook.worksheets[1:])}"
    format_pourcentage(sheet[f"F{ligne}"])


def fill_kpis_consolides(sheet, ligne):
    """
    Remplissage du tableau KPIS CONSOLIDES
    :param sheet:
    :param ligne:
    :return:
    """
    # CA
    sheet[f"C{ligne}"].value = f"=D{ligne - 5}"
    format_euro(sheet[f"C{ligne}"])

    # Marge de prod
    sheet[f"C{ligne + 1}"].value = f"=E{ligne - 5}"
    format_euro(sheet[f"C{ligne + 1}"])

    # Taux marge prod
    sheet[f"C{ligne + 2}"].value = f"=C{ligne + 1} / C{ligne}"
    format_pourcentage(sheet[f"C{ligne + 2}"])

    # Cout interne
    sheet[f"C{ligne + 3}"].value = f"=G{ligne - 5}"
    format_euro(sheet[f"C{ligne + 3}"])

    # Marge BU
    sheet[f"C{ligne + 4}"].value = f"=H{ligne - 5}"
    format_euro(sheet[f"C{ligne + 4}"])

    # Taux marge BU
    sheet[f"C{ligne + 5}"].value = f"=C{ligne + 4} / C{ligne}"
    format_pourcentage(sheet[f"C{ligne + 5}"])


# Fonction principale
def creation_overview_sheet(workbook, overview_sheet):
    """
    Remplissage de ma feuille overview du reporting, récap de touts les feuilles managers
    :param workbook:
    :param overview_sheet:
    :return:
    """
    # Overview
    ligne = 2
    fill_overview_table(overview_sheet, ligne)
    display_border("B", 2, overview_sheet)

    # KPIs Mensuels
    ligne += 3
    fill_kpis_mensuels(workbook, overview_sheet, ligne)
    display_border("B", ligne - 1, overview_sheet)

    # KPIs Consolidés
    ligne += len(workbook.worksheets[1:]) + 5
    fill_kpis_consolides(overview_sheet, ligne)
    display_border("B", ligne - 1, overview_sheet)
