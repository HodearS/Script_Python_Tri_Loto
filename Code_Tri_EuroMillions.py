import pandas as pd
import numpy as np
import os

import pandas as pd

# Chargement du fichier Excel dans un DataFrame
source_file = 'E:/Informatique/Projet/ProjetPython/pythonProject/Fichier_excel/test.xlsx'
df = pd.read_excel(source_file)

# Convertir la colonne "date_de_tirage" en datetime et la reformater
df['date_de_tirage'] = pd.to_datetime(df['date_de_tirage']).dt.strftime('%Y/%m/%d')

# Sélectionner plusieurs colonnes
colonnes_B_C_F_G_H_I_J_K_L_M_N = df[
    ['jour_de_tirage', 'date_de_tirage', 'boule_1', 'boule_2', 'boule_3', 'boule_4', 'boule_5', 'etoile_1', 'etoile_2',
     'boules_gagnantes_en_ordre_croissant', 'etoiles_gagnantes_en_ordre_croissant']]

# Compter le nombre d'occurrences de chaque nombre compris entre 1 et 50 sur les colonnes boule_1 à boule_5
boules = ['boule_1', 'boule_2', 'boule_3', 'boule_4', 'boule_5']
occurences_boules = pd.concat([df[boule].value_counts() for boule in boules], axis=1).fillna(0).sum(axis=1)
occurences_boules = occurences_boules.loc[occurences_boules.index.isin(range(1, 51))].sort_index()

# Compter le nombre d'occurrences de chaque nombre compris entre 1 et 12 sur les colonnes etoile_1 et etoile_2
etoiles = ['etoile_1', 'etoile_2']
occurences_etoiles = pd.concat([df[etoile].value_counts() for etoile in etoiles], axis=1).fillna(0).sum(axis=1)
occurences_etoiles = occurences_etoiles.loc[occurences_etoiles.index.isin(range(1, 13))].sort_index()

# Identifier les boules qui sont apparues le moins récemment
df['date_de_tirage'] = pd.to_datetime(df['date_de_tirage'], format='%Y/%m/%d')
least_recent_and_least_seen_boules = df.melt(id_vars=['date_de_tirage'], value_vars=boules, var_name='boule', value_name='number') \
    .sort_values(by='date_de_tirage') \
    .drop_duplicates(subset='number', keep='last') \
    .sort_values(by='date_de_tirage')

# Identifier les étoiles qui sont apparues le moins récemment
least_recent_and_least_seen_etoiles = df.melt(id_vars=['date_de_tirage'], value_vars=etoiles, var_name='etoile', value_name='number') \
    .sort_values(by='date_de_tirage') \
    .drop_duplicates(subset='number', keep='last') \
    .sort_values(by='date_de_tirage')

# Reformater la colonne "date_de_tirage" en "année/mois/jour"
least_recent_and_least_seen_boules['date_de_tirage'] = least_recent_and_least_seen_boules['date_de_tirage'].dt.strftime('%Y/%m/%d')
least_recent_and_least_seen_etoiles['date_de_tirage'] = least_recent_and_least_seen_etoiles['date_de_tirage'].dt.strftime('%Y/%m/%d')

# Chemin du nouveau fichier Excel
output_file = 'E:/Informatique/Projet/ProjetPython/pythonProject/Fichier_excel/donnees_extraites.xlsx'

# Créer un fichier Excel avec plusieurs feuilles
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Écrire les données extraites dans la feuille 1
    colonnes_B_C_F_G_H_I_J_K_L_M_N.to_excel(writer, sheet_name='Feuille1', index=False)

    # Écrire les occurrences des boules dans une nouvelle feuille
    occurences_boules_df = occurences_boules.reset_index()
    occurences_boules_df.columns = ['Nombre', 'Occurrences']
    occurences_boules_df.to_excel(writer, sheet_name='Occurrences_Boules', index=False)

    # Écrire les occurrences des étoiles dans une nouvelle feuille
    occurences_etoiles_df = occurences_etoiles.reset_index()
    occurences_etoiles_df.columns = ['Nombre', 'Occurrences']
    occurences_etoiles_df.to_excel(writer, sheet_name='Occurrences_Etoiles', index=False)

    # Écrire les boules les moins récemment apparues dans une nouvelle feuille
    least_recent_and_least_seen_boules.to_excel(writer, sheet_name='Least_Recent_Boules', index=False)

    # Écrire les étoiles les moins récemment apparues dans une nouvelle feuille
    least_recent_and_least_seen_etoiles.to_excel(writer, sheet_name='Least_Recent_Etoiles', index=False)

# Afficher un message de confirmation
print(f"Les données et occurrences ont été écrites dans le fichier {output_file}")
