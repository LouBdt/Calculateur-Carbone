# -*- coding: utf-8 -*-
"""
Created on Fri Oct 16 09:24:06 2020

@author: lou
"""

import parametres as p
import xlrd
import fonctionsMatrices
import math
import inspect
import sys

def lireBDDgeoloc():
    try:
        document = xlrd.open_workbook(p.FICHIER_BDD_GEOLOC)
    except FileNotFoundError:
        fonctionsMatrices.print_log_erreur("Le fichier de géolocalisation n'est pas trouvé à l'emplacement "+p.FICHIER_BDD_GEOLOC, inspect.stack()[0][3])
        sys.exit("Le fichier de géolocalisation n'est pas trouvé à l'emplacement "+p.FICHIER_BDD_GEOLOC)
    
    feuille_bdd = document.sheet_by_index(0)
    nbrows = feuille_bdd.nrows
    result = []
    for i in range(1,nbrows):
        result.append([feuille_bdd.cell_value(i,1),feuille_bdd.cell_value(i,6),feuille_bdd.cell_value(i,7)])
    return result


#Utilise la BDDgeoloc pour donner la distance de route entre deux codes postaux donnés
def distance_CP(CPA:int,CPB:int):
    listeCP = fonctionsMatrices.extraire_colonne_n(0,p.BDDgeoloc) #Liste des codes postaux
    erreur = False
    try:
        coordA = p.BDDgeoloc[listeCP.index(CPA,1)]
    except ValueError:
        try:
            coordA = p.BDDgeoloc[listeCP.index(10*int(CPA/10),1)] #On elève le CEDEX
        except ValueError:
            coordA = p.BDDgeoloc[listeCP.index(44000,1)] #Si le CP n'est pas dans la liste, on met Nantes, arbitrairement
            erreur = CPA
        except TypeError:
            coordA = p.BDDgeoloc[listeCP.index(44000,1)] #Si le CP est vide, on met Nantes, arbitrairement
            erreur = CPA
    if coordA[1]*coordA[2]==0: #Si une des deux coordonnées est nulle, elle est invalidée
        erreur = CPA
        coordA = p.BDDgeoloc[listeCP.index(44000,1)] #Si le CP est dans la liste, mais incorrect, idem
        
    try:
        coordB = p.BDDgeoloc[listeCP.index(CPB,1)]
    except ValueError:
        try:
            coordB = p.BDDgeoloc[listeCP.index(10*int(CPB/10),1)] #On elève le CEDEX
        except ValueError:
            coordB = p.BDDgeoloc[listeCP.index(44000,1)]
            erreur = CPB
        except TypeError:
            coordB = p.BDDgeoloc[listeCP.index(44000,1)]
            erreur = CPB
    if coordB[1]*coordB[2]==0: #Si une des deux coordonnées est nulle, elle est invalidée
        erreur = CPB
        coordB = p.BDDgeoloc[listeCP.index(44000,1)] #Si le CP est dans la liste, mais incorrect, idem
    
    distanceVO = math.acos(math.sin(coordA[2])*math.sin(coordB[2])+math.cos(coordA[2])*math.cos(coordB[2])*math.cos(coordB[1]-coordA[1]))*6371
    distance_route = p.facteur_route*distanceVO
    return [distance_route,erreur]


def formule_trigo(coordA, coordB):
    latA, longA = coordA[1], coordA[0]    
    latB, longB = coordB[1], coordB[0]
    a = math.sin(latA)*math.sin(latB)+(math.cos(latA)*math.cos(latB)*math.cos(longA-longB))
    dist = math.acos(min(a,1))*6371
    dist*=p.facteur_route
    
    return dist
    
#Idem que la precedente; mais prend en entree une coordonnee GPS et unb code postal
def distance_GPS_CP(longA:float,latA:float,CPB:int):
    listeCP = fonctionsMatrices.extraire_colonne_n(0,p.BDDgeoloc) #Liste des codes postaux
    erreur = False
    coordA = [longA,latA]
    try:
        index_CPB =listeCP.index(CPB,1)
        coordB = [p.BDDgeoloc[index_CPB][1], p.BDDgeoloc[index_CPB][2]]
    except ValueError:
        coordB = [-0.02372788,0.826194905]
        try:
            coordB = p.BDDgeoloc[listeCP.index(10*int(CPB/10),1)] #On enlève le CEDEX
        except ValueError:
            coordB = p.BDDgeoloc[listeCP.index(44000,1)]
            erreur = CPB
        except TypeError:
            coordB = p.BDDgeoloc[listeCP.index(44000,1)]
            erreur = CPB
    if coordB[0]*coordB[1]==0: #Si une des deux coordonnées est nulle, elle est invalidée
        erreur = CPB
        coordB = p.BDDgeoloc[listeCP.index(44000,1)] #Si le CP est dans la liste, mais incorrect, idem
    
    distance_route = formule_trigo(coordA, coordB)
    return [distance_route,erreur]

#Permet d'aller récupérer les codes postaux correspondant au numéro d'usine dans le fichier d'entrées manuelles
def get_cp_sites():
    ligne_debut = 12
    resultat = []
    
    try:
        document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
        fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
        sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
    
    feuille = document.sheet_by_index(0)
    i = ligne_debut
    while len(feuille.cell_value(i,24))>1 and i<30:
        if feuille.cell_value(i,25) !="":
            a = [int(feuille.cell_value(i,23)), feuille.cell_value(i,24), int(feuille.cell_value(i,25)), int(feuille.cell_value(i,27))]
            resultat.append(a)
            p.CP_FLORENTAISE.append(int(feuille.cell_value(i,25)))
            i+=1
            
        else:
            break
    return resultat

def getGPSfromCP(CP):
    listeCP = fonctionsMatrices.extraire_colonne_n(0,p.BDDgeoloc) #Liste des codes postaux
    #Cette fonction est appelée dans Export ou il risque d'y avoir des CP peu formattés
    #Donc d'abord on "nettoie" et on trie le code:
    if CP ==0 or CP==" " or CP =="":     
        CP = 44850
        coord = [0,0]
    pays = "FR"
    for c in str(CP):
        if c.isalpha() or c =='-':
            #Alors il s'agit d'un CP étranger
            #Donc on pleure un bon coup; on se retrousse les manches, et on y va
            pays = "AUTRE"
            break
    if pays =="FR":
        try:
            CP_simp = 10*int(int(CP)/10) #On enlève le cedex
        except ValueError:  #En vrai je vois pas dans quel cas de figure mais eh on sait jamais
            fonctionsMatrices.print_log_erreur("Code postal français invalide en première instance: "+str(CP), inspect.stack()[0][3])
            CP = 44000 #Au pif
        try:
            coord = p.BDDgeoloc[listeCP.index(CP)][1:]
        except ValueError:
            CP_simp = 10*int(int(CP)/10) #On enlève la localité trop précise (cedex)
            try:    #Et on réessaye !
                coord = p.BDDgeoloc[listeCP.index(CP_simp)][1:]
            except ValueError:#Bon là on ne peut plus qu'abandonner
                #print("Code postal français invalide en seconde instance: "+str(CP))
                # corriger_les_erreurs_CP([CP], 'CorrectionsCP.xlsx')
                # fonctionsMatrices.print_log_erreur("Code postal français invalide en seconde instance: "+str(CP), inspect.stack()[0][3])
                
                CP_simp2 = 100*int(int(CP)/100) #On enlève la localité trop précise (cedex)
                try:    #Et on réessaye !
                    coord = p.BDDgeoloc[listeCP.index(CP_simp2)][1:]
                except ValueError:
                    trouve = False
                    for x in CODES_POSTAUX_ETRANGERS:
                        if x[0]==CP:
                            coord = [x[1],x[2]]
                            trouve = True
                    if not trouve:
                        fonctionsMatrices.print_log_erreur("Code postal introuvable en troisième instance: "+str(CP), inspect.stack()[0][3])
                        p.LISTE_CODES_POSTAUX_ERRONES.append([CP, "Ventes"])
                        CP = 44000
                        coord = p.BDDgeoloc[listeCP.index(CP)][1:]
    else:
        trouve = False
        for x in CODES_POSTAUX_ETRANGERS:
            if x[0]==CP:
                coord = [x[1],x[2]]
                trouve = True
                break
        if not trouve:
            p.LISTE_CODES_POSTAUX_ERRONES.append([CP, "Ventes (CP étranger)"])
            fonctionsMatrices.print_log_erreur("Il faut ajouter le CP étranger "+str(CP)+" à la BDD géoloc secondaire", inspect.stack()[0][3])
            coord = [0,0]
    return coord

def get_cp_inter():
    try:
        document = xlrd.open_workbook(p.FICHIER_BDD_GEOLOC)
    except FileNotFoundError:
        fonctionsMatrices.print_log_erreur("Le fichier de géolocalisation n'est pas trouvé à l'emplacement "+p.FICHIER_BDD_GEOLOC, inspect.stack()[0][3])
        sys.exit("Le fichier de géolocalisation n'est pas trouvé à l'emplacement "+p.FICHIER_BDD_GEOLOC)
        
    feuille_bdd = document.sheet_by_index(0)
    ligne = 5
    col = 11
    res = []
    while ligne<100:
        l = []
        if feuille_bdd.cell_value(ligne,col)!="":
            for c in [col, col+1, col+2]:
                l.append(feuille_bdd.cell_value(ligne,c))
            res.append(l)
        ligne +=1
    return res
        
CODES_POSTAUX_ETRANGERS = get_cp_inter()


