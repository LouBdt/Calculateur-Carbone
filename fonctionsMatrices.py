# -*- coding: utf-8 -*-
"""
Created on Fri Oct 16 09:03:06 2020

@author: lou
"""

import random
import parametres as p
import logging
from colorama import Fore, Style, init
import inspect
init(convert=True)

#Retourne une liste correspondant à la colonne n (de 0 à len(matrice)-1) de matrice
def extraire_colonne_n(n:int,matrice):
    return [matrice[i][n] for i in range(len(matrice))]

#Retourne les éléments de la liste en entrée dans une liste sans doublons
def liste_unique(entree:list):
    result =[]
    for x in entree:
        if x not in result:
            result.append(x)
    return result

#Renvoie une liste de n éléments pseudo-aléatoires pris entre 0 et population
def random_sample(n:int,population:int):
    if n>population:
        return False
    else:
        return random.sample(range(0,population),n)

#Similaire à la fonction RECHERCHE dans excel:
#Retourne l'élément de la colonne col_resultat correspondant à celui de la colonne col_recherche ayant pour nom nom_element
def recherche_elem(nom_element:str,matrice:list,col_recherche:int, col_resultat:int):
    index = (extraire_colonne_n(col_recherche, matrice)).index(nom_element)
    return matrice[index][col_resultat]


#Fait la moyenne des grandeurs concernées par un nom en commun
def moyenne_recherche_elem_commun(nom_commun, matrice:list, col_recherche:int, col_moyenne:int):
    compte = 0
    somme = 0
    for ligne in matrice:
        if ligne[col_recherche] in nom_commun:
            somme += ligne[col_moyenne]
            compte +=1
    return somme/compte


#Range la matrice dans l'ordre indiqué (par défaut décroissant) selon la valeur de la colonne n
def bubbleSortColonne(matrice:list, n:int, decroissant = True ):
    swapping = True
    while swapping:
        swapping = False
        for i in range(len(matrice)-1):
            if matrice[i][n]>matrice[i+1][n] and decroissant:
                swapping = True
                tmp = matrice[i]
                matrice[i] =matrice[i+1]
                matrice[i+1] = tmp
            elif matrice[i][n]<matrice[i+1][n] and not decroissant:
                swapping = True
                tmp = matrice[i]
                matrice[i] =matrice[i+1]
                matrice[i+1] = tmp
    return matrice

import numpy as np
""" implémentation d'une méthode de Levenshtein pour comparer
deux chaines de caractères. Progammation dynamique, code ici https://stackabuse.com/levenshtein-distance-and-text-similarity-in-python/"""
def levenshtein(seq1, seq2):
    size_x = len(seq1) + 1
    size_y = len(seq2) + 1
    matrix = np.zeros ((size_x, size_y))
    for x in range(size_x):
        matrix [x, 0] = x
    for y in range(size_y):
        matrix [0, y] = y

    for x in range(1, size_x):
        for y in range(1, size_y):
            if seq1[x-1] == seq2[y-1]:
                matrix [x,y] = min(
                    matrix[x-1, y] + 1,
                    matrix[x-1, y-1],
                    matrix[x, y-1] + 1
                )
            else:
                matrix [x,y] = min(
                    matrix[x-1,y] + 1,
                    matrix[x-1,y-1] + 1,
                    matrix[x,y-1] + 1
                )
    #print (matrix)
    return (matrix[size_x - 1, size_y - 1])


def print_log_erreur(message:str, fonction:str ):
    if p.DISPLAY_ERROR:
        logging.warning("► " + message+" (survenu dans "+fonction+")")
    p.MESSAGES_DERREUR.append(["ERREUR", message, fonction])
    return None 
def print_log_info(message:str, fonction:str ):
    if p.DISPLAY_INFO:
        logging.info("► "+ message+" (depuis "+fonction+")")
    p.MESSAGES_DERREUR.append(["INFO", message, fonction])
    return None 
def print_log_log(message:str):
    if p.DISPLAY_LOG:
        logging.debug("► "+ message)
    p.MESSAGES_DERREUR.append(["LOG", message, "log"])

import xlrd
def charger_parametres():
    try:
        document = xlrd.open_workbook('Parametres de calcul.xlsx')
    except FileNotFoundError:
        print_log_erreur("Le fichier des intrants n'est pas trouvé à l'emplacement 'Parametres de calcul.xlsx'", inspect.stack()[0][3])
        print_log_erreur("On poursuit avec les paramètres par défaut", inspect.stack()[0][3])
        ok = False
        while ok == False:
            entree = input("Veuillez spécifier manuellement l'année de calcul: ")
            try:
                entree = int(entree)
                if 1900<entree <2050:
                    p.ANNEE = entree
                    ok = True
                else:
                    print_log_erreur("Ceci n'est pas un TARDIS, l'année "+str(entree)+" n'est pas calculable", inspect.stack()[0][3])
            except ValueError:
                print_log_erreur("Veuillez rentrer une date valable (exemple: '2020')", inspect.stack()[0][3])
        print_log_erreur("Paramètres par défaut appliqués (calcul de tous les posts, arborescence du projet classique) pour l'année "+str(p.ANNEE), inspect.stack()[0][3])
        return None
    document = xlrd.open_workbook('Parametres de calcul.xlsx')
    ws = document.sheet_by_index(0)
    n = ws.nrows
    
    colonne_descrip = 1
    colonne_valeur = colonne_descrip+1
    colonne_variable = colonne_valeur+1
    colonne_type = colonne_variable +1
    
    ligne_init = 9
    ligne = ligne_init
    while ligne<n:
            #J'ai pas trouvé de manière plus élégante mais au moins c'est clair
        nom = ws.cell_value(ligne, colonne_variable)
        valeur =  ws.cell_value(ligne, colonne_valeur)
        type_set = ws.cell_value(ligne, colonne_type)
        mpl_logger = logging.getLogger('matplotlib')
        mpl_logger.setLevel(logging.WARNING) 
        if nom == "DISPLAY_INFO":
            p.DISPLAY_INFO= valeur=="True"
            if type(p.DISPLAY_INFO)==type(True) and valeur:
                logging.getLogger().setLevel(logging.INFO)
                logging.info(" ► Affichage des messages informatifs activé")
        elif nom == "DISPLAY_LOG":
            logging.getLogger().setLevel(logging.DEBUG)
            p.DISPLAY_LOG= valeur=="True"
        elif nom == "SAVE_LOG":
            logging.basicConfig(level = logging.DEBUG, filename = "messages.log", filemode = "w")
            p.DISPLAY_LOG= valeur=="True"
        ligne +=1
        
    ligne = ligne_init
    while ligne<n:
        nom = ws.cell_value(ligne, colonne_variable)        
        valeur =  ws.cell_value(ligne, colonne_valeur)
        type_set = ws.cell_value(ligne, colonne_type)
        if nom !="" and valeur != "":
            if nom =="ANNEE":
                p.ANNEE = int(valeur)
            elif nom == "FICHIER_ENTREES_MANUELLES":
                p.FICHIER_ENTREES_MANUELLES= valeur
            elif nom == "FICHIER_BDD_GEOLOC":
                p.FICHIER_BDD_GEOLOC= valeur
            elif nom == "CHEMIN_ECRITURE_RESULTATS":
                p.CHEMIN_ECRITURE_RESULTATS= valeur
            elif nom == "ELECTRICITE_ET_AUTRES":
                p.ELECTRICITE_ET_AUTRES= valeur=="True"
            elif nom == "INTRANTS_ET_FRET":
                p.INTRANTS_ET_FRET= valeur=="True"
            elif nom == "EMBALLAGES_ET_SACHERIE":
                p.EMBALLAGES_ET_SACHERIE= valeur=="True"
            elif nom == "FRET_AVAL":
                p.FRET_AVAL= valeur=="True"
            elif nom == "DISPLAY_CONSOLE":
                p.DISPLAY_CONSOLE=valeur=="True"
            elif nom == "DISPLAY_GRAPH":
                p.DISPLAY_GRAPH= valeur=="True"
            elif nom == "DISPLAY_ERROR":
                p.DISPLAY_ERROR=valeur=="True"
            
            if p.DISPLAY_INFO:
                logging.info(" ► Variable '"+nom+"' (type "+ type_set+") mise à "+str(valeur))
        ligne +=1
        
    ligne = ligne_init
    while ligne<n:
        nom = ws.cell_value(ligne, colonne_variable)
        valeur =  ws.cell_value(ligne, colonne_valeur)
        type_set = ws.cell_value(ligne, colonne_type)
        if nom !="" and valeur != "":
            if type(valeur)==type("Chaine de caracteres"):
                if "ANNEE" in valeur:
                    valeur = valeur.replace("ANNEE", str(p.ANNEE))
                elif "CHEMIN_ECRITURE_RESULTATS" in valeur:
                    valeur = valeur.replace("CHEMIN_ECRITURE_RESULTATS/", str(p.CHEMIN_ECRITURE_RESULTATS))
            if nom == "DIRECTION_FICHIER_INTRANTS":
                p.DIRECTION_FICHIER_INTRANTS= valeur
            elif nom == "DIRECTION_FICHIER_VENTES1":
                p.DIRECTION_FICHIER_VENTES1= valeur
            elif nom == "DIRECTION_FICHIER_VENTES2":
                p.DIRECTION_FICHIER_VENTES2= valeur
            elif nom == "DIRECTION_FICHIER_SACHERIE":
                p.DIRECTION_FICHIER_SACHERIE= valeur
            if p.DISPLAY_INFO:
                logging.info(Fore.CYAN +" ► "+Style.RESET_ALL+"Variable '"+nom+"' (type "+ type_set+") mise à "+str(valeur))
        ligne += 1
        
