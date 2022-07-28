def find_doublons_index(database):
    list_index_non_traite = [ nb for nb in range(len(database))]
    doublons_index = []
    avancement_trie = 0

    while list_index_non_traite != []:
        doublons_index.append(
            { "manager": database[list_index_non_traite[0]]["manager"]["nom"],
              "index": [list_index_non_traite[0]]
              }
        )

        # On ajoute tous les index de doublons à la liste du manager
        for index in list_index_non_traite[1:]:
            if doublons_index[avancement_trie]["manager"] == database[index]["manager"]["nom"]:
                doublons_index[avancement_trie]["index"].append(index)

        # On supprime tous les index qui ont été traité (index déjà assocé à une personne)
        for index_traite in doublons_index[-1]["index"]:
            list_index_non_traite.remove(index_traite)
        avancement_trie += 1

    return doublons_index

def merge_doublons(database, list_doublons):

    for doublon in list_doublons:

        first_manager_index = doublon["index"][0]
        for index_doubon in doublon["index"][1:]:

            # Fusion de chaque élément
            for key, value in database[index_doubon].items():
                # Si int or float
                if isinstance(value, (int, float)):
                    database[first_manager_index][key] += value
                # Si list
                elif isinstance(value, list):
                    for element in value:
                        database[first_manager_index][key].append(element)

                # On vide le doublon fusionné
                database[index_doubon] = -1

    # Suppression des doublons vidés
    return [value for value in database if value != -1]

def delete_doublons(database):
    list_index_doublons = find_doublons_index(database)
    return merge_doublons(database, list_index_doublons)