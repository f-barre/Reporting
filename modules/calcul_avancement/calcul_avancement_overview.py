# Sous-Fonctions
def fill_kpis_mensuels(bp_sheet, manager_sheet, month, start_ligne):
    """
    Calcul des avancements dans le tableau KPIS MENSUELS
    :param bp_sheet:
    :param manager_sheet:
    :param month:
    :param start_ligne:
    :return:
    """
    # BP
    # Nb de consultants
    manager_sheet[f"C{start_ligne}"].value = bp_sheet.cell(row=3, column=month + 2).value
    # CA
    manager_sheet[f"D{start_ligne}"].value = bp_sheet.cell(row=6, column=month + 2).value
    # Marge de prod
    manager_sheet[f"E{start_ligne}"].value = bp_sheet.cell(row=9, column=month + 2).value
    # Cout interne
    manager_sheet[f"G{start_ligne}"].value = bp_sheet.cell(row=12, column=month + 2).value
    # Marge BU
    manager_sheet[f"H{start_ligne}"].value = bp_sheet.cell(row=14, column=month + 2).value

    # Avancement
    # Nb de consultants
    manager_sheet[f"C{start_ligne + 1}"].value = f"=C{start_ligne - 1} - C{start_ligne}"
    # CA
    manager_sheet[f"D{start_ligne + 1}"].value = f"=D{start_ligne - 1} - D{start_ligne}"
    # Marge de prod
    manager_sheet[f"E{start_ligne + 1}"].value = f"=E{start_ligne - 1} - E{start_ligne}"
    # Internes
    manager_sheet[f"G{start_ligne + 1}"].value = f"=F{start_ligne - 1} - F{start_ligne}"
    # Marge BU
    manager_sheet[f"H{start_ligne + 1}"].value = f"=H{start_ligne - 1} - H{start_ligne}"


def fill_kpis_consolides(bp_sheet, manager_sheet, month, start_ligne):
    """
    Calcul des avancements dans le tableau KPIS CONSOLIDES
    :param bp_sheet:
    :param manager_sheet:
    :param month:
    :param start_ligne:
    :return:
    """
    # BP
    # CA total
    manager_sheet[f"D{start_ligne + 1}"].value = bp_sheet.cell(row=6, column=month + 2).value
    # Marge de production
    manager_sheet[f"D{start_ligne + 2}"].value = bp_sheet.cell(row=9, column=month + 2).value

    # Coût interne
    manager_sheet[f"D{start_ligne + 4}"].value = bp_sheet.cell(row=12, column=month + 2).value
    # Marge BU
    manager_sheet[f"D{start_ligne + 5}"].value = bp_sheet.cell(row=14, column=month + 2).value

    # Avancement
    # CA total
    manager_sheet[f"E{start_ligne + 1}"].value = f"=C{start_ligne + 1} - D{start_ligne + 1}"
    # Marge production total
    manager_sheet[f"E{start_ligne + 2}"].value = f"=C{start_ligne + 2} - D{start_ligne + 2}"

    # Coût interne
    manager_sheet[f"E{start_ligne + 4}"].value = f"=C{start_ligne + 4} - D{start_ligne + 4}"
    # Marge BU
    manager_sheet[f"E{start_ligne + 5}"].value = f"=C{start_ligne + 5} - D{start_ligne + 5}"


# Fonction principale
def calcul_avancement_overview(bp_workbook, reporting_workbook, month):
    """
    Calcul des avancements dans la feuille d'overview
    :param bp_workbook:
    :param reporting_workbook:
    :param month:
    :return:
    """
    # On récupère les sheet à modifier
    suivi_bp_sheet = bp_workbook.worksheets[0]
    overview_sheet = reporting_workbook.worksheets[0]

    # Fill KPIs Mensuels tableau
    # Get ligne début du tableau
    start_ligne = 1
    while overview_sheet[f"B{start_ligne}"].value != "BP":
        start_ligne += 1

    fill_kpis_mensuels(suivi_bp_sheet, overview_sheet, month, start_ligne)

    # Get ligne début dU TABLEAU
    while overview_sheet[f"B{start_ligne}"].value != "KPIS CONSOLIDÉS":
        start_ligne += 1

    # Fill KPIs Consolidé tableau
    fill_kpis_consolides(suivi_bp_sheet, overview_sheet, month, start_ligne)
