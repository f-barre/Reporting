# Modules / Dépendances
from configuration import APP_CONFIG

# Tools
from tools.gui_class import GUI

# Modules
from modules.gui_callback.a_creation_db import a_creation_db
from modules.gui_callback.b_creation_bp_vide_et_creation_primes_et_creation_payement_forfait import \
    b_creation_bp_vide_et_creation_primes_et_creation_payement_forfait
from modules.gui_callback.c_creation_reporting import c_completer_reporting_mensuel
from modules.gui_callback.d_calcul_avancement import d_calcul_avancement

list_callback = [
    "1°) Créer la base de données", a_creation_db,
    "2°) Créer les feuilles d'exceptions, pour les consultants payés au forfait et les feuilles de primes",
    b_creation_bp_vide_et_creation_primes_et_creation_payement_forfait,
    "3°) Créer les reporting", c_completer_reporting_mensuel,
    "4°) Calculer l'avancement sur les reporting à partir du BP", d_calcul_avancement
]

gui_properties = APP_CONFIG.GUI_PROPERTIES
# Création gui
app = GUI()
app.init(size="x".join(gui_properties["size"]), bg=gui_properties["bg"], title=gui_properties["title"])

# Calcul des dimensions des boutons
btn_height = (int(gui_properties["size"][1]) - (gui_properties["nb btns"] - 1) * gui_properties["space between btn"]) / \
             gui_properties["nb btns"]

# Créations des Btns
for index_btn in range(0, gui_properties["nb btns"] * 2, 2):
    app.add_ctn(
        name=f"ctn_btn{index_btn}",
        width=int(gui_properties["size"][0]),
        height=btn_height,
        bg="grey",
        display_border=False,
        pos=[index_btn, 0]
    )

    app.add_btn(
        name=f"btn{index_btn}",
        x=1, y=1,
        text=list_callback[index_btn],
        bg="grey",
        activebackground="red",
        callback=list_callback[index_btn + 1],
        ctn=app.ctns[f"ctn_btn{index_btn}"]["object"]
    )

    app.add_ctn(
        name=f"ctn_space{index_btn}-ctn_btn{index_btn + 1}",
        width=int(gui_properties["size"][0]),
        height=gui_properties["space between btn"],
        bg="black",
        display_border=False,
        pos=[index_btn + 1, 0]
    )

app.mainloop()
