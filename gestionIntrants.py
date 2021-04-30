# -*- coding: utf-8 -*-
"""
Created on Fri Oct 16 10:32:07 2020

@author: lou
Ensemble des fonctions relatives au calcul du bilan carbone du fret des matières en amont des usines Florentaise. Contient:
    lireBDDMP()
    gestionpb_volume_sac()
    lireFE_matprem()
    associerMPetFE_fab()
"""
import xlrd
import inspect
import fonctionsMatrices
import parametres as p
import logging
import numpy as np
import sys

#Renvoie la liste des achats de l'année étudiée
def lireBDDMP():
    try:
        document = xlrd.open_workbook(p.DIRECTION_FICHIER_INTRANTS)
    except FileNotFoundError:
        fonctionsMatrices.print_log_erreur("Le fichier des intrants n'est pas trouvé à l'emplacement "+p.DIRECTION_FICHIER_INTRANTS, inspect.stack()[0][3])
        sys.exit("Le fichier des intrants n'est pas trouvé à l'emplacement "+p.DIRECTION_FICHIER_INTRANTS)
    feuille_bdd = document.sheet_by_index(0)
    nbrows = feuille_bdd.nrows
    MPInclues = [182757, 180409]
    CP_exclus = [50500,49520,40210,7170,49700,29530,91410,44850,44850,1370]
    result = [["idMP","nomMP", "refMP", "client","famille","fournisseur", "cp_depart", "cp_depot_arrivee", "pays_depart","quantite","unité"]]
    #On lit chaque ligne du fichier d'achat et on récupère les colonnes intéressantes
    for i in range(1,nbrows):
        idMP = int(feuille_bdd.cell_value(i,0))
        nomMP = feuille_bdd.cell_value(i,2)
        refMP = feuille_bdd.cell_value(i,3)
        num_depot_arrive= int(feuille_bdd.cell_value(i,1))
        client = int(feuille_bdd.cell_value(i,4))
        famille = feuille_bdd.cell_value(i,8)
        fournisseur = feuille_bdd.cell_value(i,10)
        cp_depart = feuille_bdd.cell_value(i,11)
        pays_depart = feuille_bdd.cell_value(i,12)
        
        #On transpose le numero de depot en code postal du site
        try:
            cp_depot_arrivee = int(fonctionsMatrices.recherche_elem(num_depot_arrive,p.CP_SITES_FLORENTAISE,0,2))
        except ValueError:
            if num_depot_arrive!= 369:
                fonctionsMatrices.print_log_erreur("Code postal du site "+str(num_depot_arrive)+" non trouvé", inspect.stack()[0][3])
                cp_depot_arrivee = 44850
                
        if (cp_depart not in CP_exclus or cp_depart==cp_depot_arrivee or idMP in MPInclues) and num_depot_arrive!=369:
            for n in [5,6,7,14]:
                quantite = 0
                unit = "kg"
                qte = feuille_bdd.cell_value(i,n)
                if qte =="":
                    qte = 0
                if qte>0:
                    if n == 5:#KG
                        unit = "kg"
                        quantite = qte
                    elif n== 6:#M3
                        unit = "m3"
                        quantite = qte
                    elif n==7:#PALETTES
                        unit = "kg"
                        quantite = p.MASSE_PAR_PALETTE*qte*1000
                    elif n ==14:#TONNES
                        unit = "kg"
                        quantite = qte*1000
                    break
                    
            
            result.append([idMP,nomMP, refMP, client,famille,fournisseur, cp_depart, cp_depot_arrivee, pays_depart,quantite,unit.lower()])
    return result


#Cette fonction sert à gérer le problème des sacs:
    #on achète des sacs dont la contenance est en KG ou en L dans le nom du produit
    #Cette fonction parcourt donc le nom du produit pour tenter de repérer une séquence
    #Cette séquence doit être une série de chiffres suivie de L ou kg
def gestionpb_volume_sac(nom:str):
    if nom=="TRUFFAUT POUZZOLANE 15L 72 SACS":
        return 15*72/1000,"m3"
    elif nom=="COPODECOR ECORCES DE PIN 10/25 60":
        return 60/1000, "m3"
    elif nom =="CERAMIQUE PPC GREEN GRADE":
        return 22.7, "kg"     #Au pif
    elif nom =="COPOCAO COQUES DE CACAO":
        return 120/1000,"m3"      #Au pif aussi
    elif nom=="Rajouter ici les noms et quantités manuellement":
        return 42,"unite"
    
    #On ne fait pas le cas où la chaine de charactères est nulle
    if nom!="":
        chiffres = ""
        unit = ""
        res = 0
        identifie = False        
        i=0
        while i<len(nom):
            while i<len(nom) and nom[i].isdigit(): #important de tester dans cet ordre
                chiffres+= str(nom[i])
                i+=1
            if i>=len(nom):
                break
            if len(chiffres) != 0 and (nom[i] =="L" or nom[i]=="l"):
                identifie = True
                unit = "m3"
                res = int("".join(chiffres))
                res /= 1000 #Penser à convertir les L en m3 !
                break
            elif len(chiffres) != 0 and (nom[i].lower()=="k" and (i+1)<len(nom)) and nom[i+1].lower()=="g":
                identifie = True
                unit = "kg"
                res = int("".join(chiffres))
                break
            #S'il y a un espace entre la qté et l'unité
            elif (nom[i].lower()==" "and (i+1)<len(nom)) and nom[i+1].isalpha(): 
            #On passe au caractère suivant, ça passe pas par le "else" ni par le second "while" et tout rentre dans l'ordre !c
                i+=1    
            else:
                chiffres = ""
                i+=1
        if identifie:
            return res,unit
        else:
            fonctionsMatrices.print_log_erreur("Detection de la quantité par sac indéchiffrable pour: "+nom+". A rajouter dans les exception de gestionIntrants.gestionpb_volume_sac() ", inspect.stack()[0][3])
            return 1,"kg"  #Par défaut, si on a vraiment rien pu trouver, on donne 1KG
    return 0,'kg'    
        
        
#Va lire dans le ficher Entrees Manuelles les FE des matières premières. Retourne les tableaux
#FE_familles: [id, nom de la famille, FE de la famille]
#MP_familles_N: [id, n°MP (?), nom MP, réf interne (?), famille florentaise (?), famille bilan carbone, %N des engrais]
#masses_vol_MP: [id, nom MP, masse volumique (t/m3)]
#FE_engrais: [id, nom engrais, FE engrais]
def lireFE_matprem():
    debut = 10   #Ligne de début des tableaux T1 T2 et T3
    fin_T1 = 47  #ligne de fin du premier tableau (FE par familles)
    debut_T4 = 52 #debut et fin du tableau 4 (engrais)
    fin_T4 = 56
    
    
    try:
        document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
        fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
        sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
    
    feuille_MP = document.sheet_by_index(1)

    FE_familles = []
    MP_familles_N = []
    masse_vol_MP = []
    FE_engrais = []
    
    for i in range(debut-1,fin_T1):
        FE_familles.append([feuille_MP.cell_value(i,2),feuille_MP.cell_value(i,3),feuille_MP.cell_value(i,4)])
        
    for i in range(debut-1,feuille_MP.nrows):
        ligne = []
        for j in range(7,14):
            ligne.append(feuille_MP.cell_value(i,j))
        MP_familles_N.append(ligne)
        
    for i in range(debut-1,feuille_MP.nrows):
        if feuille_MP.cell_value(i,16)!="":
            masse_vol_MP.append([feuille_MP.cell_value(i,15),feuille_MP.cell_value(i,16),feuille_MP.cell_value(i,17)])
        
    for i in range(debut_T4-1,fin_T4):
        if feuille_MP.cell_value(i,3) !="":
            FE_engrais.append([feuille_MP.cell_value(i,2),feuille_MP.cell_value(i,3),feuille_MP.cell_value(i,4)])
        
    return FE_familles, MP_familles_N, masse_vol_MP, FE_engrais

def associerMPetFE_fab(MP_familles_N:list,FE_engrais:list,FE_familles:list):
    liste_engrais_eol = [["Nom engrais", "Taux d'azote (%)", "FE épandage (tCO2e/t_azote)", "Masse (t)", "FE épandag (kgCO2e/t_engrais)"]]
    MP_et_FE = [["Nom MP", "FE (kgCO2e/t)"]]
    for i in range(len(MP_familles_N)):
        nom_famille_BC = MP_familles_N[i][5]
        if "engrais" in nom_famille_BC.lower() or "ammonitrate" in nom_famille_BC.lower() or "urée" in nom_famille_BC.lower() or "uree" in nom_famille_BC.lower():
            tx_N= MP_familles_N[i][6]
            
            try:
                tx_N = float(tx_N)
            except ValueError:
                fonctionsMatrices.print_log_erreur("Taux d'azote invalide pour l'engrais "+str(MP_familles_N[i][2])+". On lui attribue un taux de 10%", inspect.stack()[0][3])
                tx_N = 0.1
            try:
                FE = tx_N * fonctionsMatrices.recherche_elem(nom_famille_BC,FE_engrais, 1,2)
            except ValueError:
                fonctionsMatrices.print_log_erreur("La famille de l'engrais "+str(MP_familles_N[i][2])+" ("+str(nom_famille_BC)+
                                                   ") n'a pas été trouvé. On lui attribue un FE de 479.5 kgCO2/t (engrais moyen azoté à 10%)", inspect.stack()[0][3])
                FE = 479.5
            liste_engrais_eol.append([MP_familles_N[i][2],tx_N, p.FE_EMPANDAGE_ENGRAIS*p.PRG_N2O, 0,1000*tx_N*p.FE_EMPANDAGE_ENGRAIS*p.PRG_N2O])
        else:
            try:
                FE= fonctionsMatrices.recherche_elem(nom_famille_BC,FE_familles, 1,2)
            except ValueError:
                fonctionsMatrices.print_log_erreur("L'élément "+str(nom_famille_BC)+" n'a pas été trouvé. On lui attribue un FE de 0", inspect.stack()[0][3])                
                FE = 0
        MP_et_FE.append([MP_familles_N[i][2],FE])
        
    return MP_et_FE, liste_engrais_eol


def associer_nom_FEfab(listeAchats:list, MP_et_FE:list, liste_engrais_eol:list):
    nomsMP = fonctionsMatrices.liste_unique(fonctionsMatrices.extraire_colonne_n(1, listeAchats))
    
    MP_et_FEok = [x for x in MP_et_FE if x[0]!=""]
    n = len(MP_et_FEok)
    compte = 0
    c_m1 = 0
    c_m2 = 0
    c_m3 = 0
    c_m4 = 0
    liste_introuves = []
    for i in range(1,len(nomsMP)):
        nom = nomsMP[i]
        nomFamille =  "A"
        #On cherche le facteur d'émission à la FABRICATION de la matière première (en kgCO2e par tonne de MP)
        trouve = False
        #On parcourt la liste des facteurs d'émission
        for k in MP_et_FEok[1:]:
            try:
                float(k[1])
            except ValueError:#L'exception arrive souvent quand on a un caractère vide -> on suppose que c'est 0
                k[1] = 0
                logging.info("Le FE de "+nom+" est vide. On lui attribue un FE de 0.", inspect.stack()[0][3])

            if nom.lower()==k[0].lower():
                trouve = True
                compte +=1
                c_m1 +=1
                FE_MPfab = k[1]
                break
            
        if not trouve:        #Mais on va essayer de le trouver quand même
            #D'abord on vire les caractères qui ne sont pas des lettres
            nom_recherche = ""
            for c in nom:
                if c.isalpha():
                    nom_recherche+=c.lower()
            for k in MP_et_FEok[1:]:
                nom_k = ""
                for c in k[0]:#On met aussi en minuscule et simplifié cette liste
                    if c.isalpha():
                        nom_k+=c.lower()
                
                #Là on trouve si la seule différence était un chiffre ou autre symbole
                if nom_k!="" and (nom_k ==nom_recherche or nom_k in nom_recherche or nom_recherche in nom_k):
                    p.LISTE_ASSIMILATIONS.append([nom, k[0], "Comparaison des lettres"])
                    trouve = True
                    compte +=1
                    FE_MPfab = k[1]
                    c_m2 +=1
                    MP_et_FEok.append([nom, k[1], k[0]])
                    break
                
        if not trouve: #Si on a toujours pas trouvé, on récupère la distance de Levenshtein entre les deuxchaines de caractères
                       #et on prend celle qui ressemble le plus
            distance_min = np.inf
            temp = []
            tolerance = 8 # nombre de caractères d'erreurs acceptés
            taux = 0.4
            for k in MP_et_FEok[1:]:
                distanceL = fonctionsMatrices.levenshtein(k[0].lower(), nom.lower())
                if distanceL<tolerance:
                    p.LISTE_ASSIMILATIONS.append([nom, k[0], "Distance de Levenshtein ("+distanceL+")"])
                    trouve = True
                    c_m3 +=1
                    FE_MPfab = k[1]
                    MP_et_FEok.append([nom, k[1],k[0]])
                    compte +=1
                    break
                elif distanceL<distance_min:
                    distance_min = distanceL
                    temp = [nom, k[1], k[0]]
                    break
            if not trouve and len(temp)>0 and distance_min< taux*min(len(k[0]), len(nom)):
                p.LISTE_ASSIMILATIONS.append([nom, k[0], "Distance de Levenshtein ("+distanceL+")"])
                trouve = True
                c_m3 +=1
                compte +=1
                FE_MPfab = temp[1]
                MP_et_FEok.append(temp)
        if not trouve: #Si là on a toujours pas trouvé, on tente l'identification des mots
            rech = []
            mot = ""
            for c in nom:
                if c.isdigit():
                    break
                elif c != " " and c.isalpha():
                    mot +=c
                elif c ==" " and mot!= "":
                    if mot not in ["DE", "TaN", "TN", "TRUFFAUT", "BOTANIC", "SYSTEME", "U", "DOR", "D", "N", "K"]:
                        if mot == "TER":
                            rech.append("TERREAU")
                        elif mot =="INT":
                            rech.append("INTERIEUR")
                        elif mot == "PTES" or mot == "PLTES":
                            rech.append("PLANTES")
                        elif mot =="PAI":
                            rech.append("PAILLIS")
                        elif mot in "AROM":
                            rech.append("AROMATIQUES")
                        elif mot in "AQUAT":
                            rech.append("AQUATIQUES")
                        rech.append(mot)
                    mot = ""
            if len(mot)!= 0 and mot.lower() not in ["l", "kg", "mm", ]:
                rech.append(mot)
            stemp = 0
            temp = []
            for k in MP_et_FEok[1:n]:
                mots_k = []
                mot = ""
                for c in k[0]:
                    if c.isdigit():
                        break
                    elif c != " " and c.isalpha():
                        mot +=c
                    elif c ==" " and mot!= "":
                        if mot not in ["DE", "TaN", "TN", "TRUFFAUT", "BOTANIC", "SYSTEME", "U", "DOR", "D", "N", "K"]:
                            if mot == "TER":
                                mots_k.append("TERREAU")
                            elif mot =="INT":
                                mots_k.append("INTERIEUR")
                            elif mot == "PTES" or mot == "PLTES":
                                mots_k.append("PLANTES")
                            elif mot =="PAI":
                                mots_k.append("PAILLIS")
                            elif mot =="AROM":
                                mots_k.append("AROMATIQUES")
                            elif mot in "AQUAT":
                                rech.append("AQUATIQUES")
                            mots_k.append(mot)
                        mot = ""
                if len(mot)!= 0 and mot.lower() not in ["l", "kg", "mm", ]:
                    mots_k.append(mot)    
                score = 0; stemp = 0; comm = [];
                comm = []
                for v in rech:
                    for w in mots_k:
                        if v.lower()==w.lower():
                            score += 1
                            comm.append(v)
                if score >= 3:
                    trouve = True
                    p.LISTE_ASSIMILATIONS.append([nom, k[0], str(score)+" mot(s) identique(s): "+''.join([m+";" for m in comm])])
                    FE_MPfab = k[1]
                    MP_et_FEok.append([nom, k[1],k[0]])
                    c_m4 +=1
                    compte +=1
                    break
                elif score >=2 and score>stemp:
                    temp, stemp, elco = [nom, k[1], k[0]], score, comm
                    
                elif score==len(mots_k):
                    p.LISTE_ASSIMILATIONS.append([nom, k[0], str(score)+" mot(s) identique(s): "+''.join([m+";" for m in comm])])
                    trouve = True
                    FE_MPfab = k[1]
                    c_m4 +=1
                    MP_et_FEok.append([nom, k[1],k[0]])
                    compte +=1
                    break
            if temp != [] and stemp>=1:
                p.LISTE_ASSIMILATIONS.append([nom, k[0], str(stemp)+" mot(s) en commun: "+''.join([m+";" for m in elco])])
                trouve = True
                compte +=1
                c_m4 +=1
                FE_MPfab = temp[1]
                MP_et_FEok.append(temp)
        if not trouve:
            liste_introuves.append(nom)
    for assim in p.LISTE_ASSIMILATIONS:
        for engrais in liste_engrais_eol:
            if assim[1] == engrais[0]:
                liste_engrais_eol.append([assim[0], engrais[1], engrais[2], engrais[3], engrais[4]])
                break
    if len(liste_introuves)>1:
        fonctionsMatrices.print_log_erreur(str(len(liste_introuves))+ 
                                           " facteurs d'émission d'intrants sont à ajouter manuellement:", inspect.stack()[0][3])                
    elif len(liste_introuves)==1:
        fonctionsMatrices.print_log_erreur("Un facteur d'émission est à ajouter manuellement: "+liste_introuves[0], inspect.stack()[0][3])                
    for introuve in liste_introuves:
        fonctionsMatrices.print_log_erreur("Ajouter " +introuve+" à la liste des intrants (FE et masse volumique) dans le fichier des Entrées Manuelles, svp", inspect.stack()[0][3])
        
    # print(liste_introuves)
    if len(MP_et_FEok)-n != 0:
        fonctionsMatrices.print_log_info("Ajout de "+str(len(MP_et_FEok)-n)+" entrées de FE dans les intrants par assimilation"+
                                         " (détail des assimilations dans le fichier résultat)", inspect.stack()[0][3])

    return MP_et_FEok, liste_engrais_eol


def BC_intrants_par_site(resultat_usine_MP:list, MP_et_FE:list):
    BC_MP_site = [[[x[0],x[2],x[1]],0] for x in p.CP_SITES_FLORENTAISE] 
    
    resultat_complet = []
    for matprem in resultat_usine_MP[1:]:
        nom = matprem[0]
        
        fe = 1
        trouve = False
        for fact in MP_et_FE:
            el = ""
            for c in fact[0]:
                if c.isalpha():
                    el+=c
            if nom.lower() in fact[0].lower() or fact[0].lower() in nom.lower():
                fe = fact[1]
                trouve = True
                break
        if not trouve:
            fonctionsMatrices.print_log_erreur("FE non trouvé pour: "+nom, inspect.stack()[0][3])


        for usine in matprem[1]:
            nom_usine = usine[0][1]
            tonnage = usine[1]
            emissions = tonnage*fe/1000 #on convertit en tCO2e
            if emissions!=0:
                resultat_complet.append([nom, nom_usine, tonnage,fe,"kgCO2e/t", emissions, "tCO2e"])
            # print(nom + "  " +nom_usine+"  BC mp: "+ str(bc_mp_usine))
            for res in BC_MP_site:
                if tonnage !=0 and res[0][2]==nom_usine:
                    res[1]+= emissions 
                    # print("Bilan carbone: "+str(emissions) +"tCO2e ("+nom+")")
                    if "tourbe" in nom.lower():
                        p.COMPARAISON_TOURBE[1]+=emissions
                        # print("Tourbe : "+str(tonnage*fe/1000)+" teCO2")
                    break
                        
    return BC_MP_site, resultat_complet

def corriger_noms_massesvol(masses_vol:list,MP_et_FE:list, listeAchats:list):
    n = len(masses_vol)
    introuves = []
    compteur = 0
    
    ajouts = [x for x in MP_et_FE if len(x)==3]
    to_add = []
    for x in ajouts:
        ajoute = False
        for mv in masses_vol:
            if x[2].lower() in mv[1].lower() or mv[1] in x[2].lower():
                compteur +=1
                ajoute = True
                to_add.append([n+compteur, x[0], mv[2]])
                break
        if not ajoute:
            introuves.append(x[0])
            
    for aj in to_add:
        masses_vol.append(aj)
    
    if len(introuves)==1:
        pass
        fonctionsMatrices.print_log_erreur("Une masse volumique est à ajouter manuellement:"+introuves[0], inspect.stack()[0][3])
    elif len(introuves)>1:
        pass
        fonctionsMatrices.print_log_erreur(str(len(introuves))+" masses volumiques sont à ajouter manuellement", inspect.stack()[0][3])
        print(introuves)
    fonctionsMatrices.print_log_info("Ajout de "+str(len(masses_vol)-n)+" entrées dans les masses volumiques par assimilation", inspect.stack()[0][3])
    return masses_vol


def calc_fret_final(fret:list,MP_et_FE:list):
    for mp in fret[1:]:
        nomMP = mp[0]
        
        #Recherche du FE de fabrication dans MP_et_FE
        FE_MPfab = 0
        for fe in MP_et_FE[1:]:
            if fe[0].lower() in mp[0].lower() or mp[0].lower() in fe[0].lower():
                FE_MPfab = fe[1]
                break
        tonnage = mp[7]
        FE_MPtrans_cumule = mp[8]+mp[9]
        FE_MPtrans_moy = mp[10]+mp[11]
        #Les FE sont toujours en kgCO2e/t
        mp.append(FE_MPfab)
        mp.append(FE_MPfab+FE_MPtrans_cumule)#Pas beaucoup de sens tel quel mais besoin pour la suite du calcul
        mp.append(FE_MPfab+FE_MPtrans_moy)#Pas nécessaire à la suite mais utile pour le BC produits
        mp.append(tonnage*mp[15]/1000) #BC cumulé matière avec conversion en t de CO2e   #(en colonne 15)
        
    
    fret[0].append("FE fabrication/extraction (kgCO2e/t)")
    fret[0].append("Somme FE fabri. +FE tranport cumulé (kgCO2e/t)")
    fret[0].append("Somme FE fabri. +FE tranport moyen (kgCO2e/t)")
    fret[0].append("BC cumulé amont de l'intrant (tCO2e)")
    
    return fret

def calcul_protoxyde(resultat_fret_usine_MP, liste_engrais_eol):

    res_protoxyde = [[x[0],0] for x in resultat_fret_usine_MP[-1][1]]
    for engrais in liste_engrais_eol:
        nom = engrais[0]
        tx_N = engrais[1]
        FE_epandage = engrais[2]
        
        for res_part in resultat_fret_usine_MP:
            if res_part[0].lower()==nom.lower():
                for res_usine in res_part[1]:
                    nom_usine = res_usine[0][1]
                    for x in res_protoxyde:
                        if nom_usine.lower()==x[0][1].lower():
                            tonnage_usine= res_usine[1]
                            emissions = tonnage_usine*FE_epandage*tx_N #Emissions totales
                            x[1] += emissions
                            break               
                break
    return  res_protoxyde, liste_engrais_eol 

def calcul_co2_tourbe(resultat_fret_usine_MP, massesvol):
    FE_eol = lire_FE_tourbe_eol()
    tourbes_eol = [["Nom tourbe", "FE EoL (kgCO2e/t)"]]
    res_co2_tourbe = [[x[0],0] for x in resultat_fret_usine_MP[-1][1]] #contient le résultat
    for mp in resultat_fret_usine_MP:
        nom = mp[0]
        if "tourbe" in nom.lower():
            #Recherche du FE adéquat
            if "blond" in nom.lower():
                FE = FE_eol[0]
            elif "brun" in nom.lower():
                FE = FE_eol[1]
            elif "noir" in nom.lower():
                FE = FE_eol[2]
            else:
                FE = FE_eol[1]
            tourbes_eol.append([nom, FE])
            
            
            for us in mp[1]: #us = usine
                nom_us = us[0][1]
                tonnage = us[1]
                emissions = tonnage*FE/1000
                for us_res in res_co2_tourbe:
                    if us_res[0][1].lower() ==nom_us.lower():
                        us_res[1] += emissions
                        p.COMPARAISON_TOURBE[2]+=emissions
                        break
            
    return res_co2_tourbe, tourbes_eol

def lire_FE_tourbe_eol():
    try:
        document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
        fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
        sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
    feuille_MP = document.sheet_by_index(1)
    tourbe_blonde= feuille_MP.cell_value(9,21)
    tourbe_brune = feuille_MP.cell_value(10,21)
    tourbe_noire = feuille_MP.cell_value(11,21)
    
    return [tourbe_blonde, tourbe_brune, tourbe_noire]
def calc_tourbe_par_usine(eol_tourbe:list,MP_par_usine:list, massesvol:list):
    res = [[usine, 0, 0] for usine in p.CP_SITES_FLORENTAISE]
    for usine in res:
        for tourbe in eol_tourbe:
            for MP in MP_par_usine:
                if tourbe[0]==MP[0]:
                    for site in MP[1]:
                        if site[0][1]==usine[0][1]!="Support":
                            #print('site: '+str(site[0][1]) +"  - tourbe: "+tourbe[0])
                            masse_vol_tourbe = 0.27 #Masse volumique moyenne d'une tourbe
                            for mv in massesvol:
                                if mv[1]==tourbe[0]:
                                    masse_vol_tourbe = mv[2]
                            
                            partMP_vol = site[1][0]
                            partMP_mas = site[1][1]
                            vtot = MP[2][0]
                            mtot = MP[2][1]
                            #print("partVOL: "+str(partMP_vol)+"  -  partMAS: "+str(partMP_mas)+" de masse et vol total de cette tourbe")
                            if mtot+vtot!=0 and partMP_vol+partMP_mas!=0:
                                part_usine= ((partMP_vol*vtot*masse_vol_tourbe)+(partMP_mas*mtot))/(mtot+(masse_vol_tourbe*vtot))
                                usine[1]+=part_usine*part_usine*(mtot+(masse_vol_tourbe*vtot))
                                usine[2]+=part_usine*tourbe[2] #Bilan carbone (tCO2/t tourbe)
                                # print("quantité totale: "+str((mtot+(masse_vol_tourbe*vtot))/1000)+" tonnes")
                                # print("part de matière dans l'usine': "+str(part_usine)+" %")
                                # print("Qté de cette MP sortant de l'usine': "+str(part_usine*(mtot+(masse_vol_tourbe*vtot)))+" tonnes")
                                # print("Qté de GES de cette tourbe sortant de cette usine :"+str(part_usine*tourbe[2]))
                                # print("")
    return res

def ajoutEoLtableauFret(fret, liste_engrais_eol,tourbes_eol):
    liste_engrais = [x[0] for x in liste_engrais_eol[1:]]
    
    for ligne in fret[1:]:
        FE_epandage = 0
        FE_EoL_tourbe = 0
        nomMP = ligne[0]
        if nomMP in liste_engrais:
            for engrais in liste_engrais_eol:
                distanceL = fonctionsMatrices.levenshtein(nomMP.lower(), engrais[0].lower())
                if distanceL<4:
                    FE_epandage =  engrais[4]  
                    break
        if "tourbe" in nomMP.lower():
            for tourbe in tourbes_eol[1:]:
                distanceL = fonctionsMatrices.levenshtein(nomMP.lower(), tourbe[0].lower())
                if distanceL<4:
                    FE_EoL_tourbe = tourbe[1]
        ligne.append(FE_epandage)
        ligne.append(FE_EoL_tourbe)
        ligne.append(ligne[13]+ligne[10]+ligne[11]+ligne[17]+ligne[18])
        ligne.append((ligne[19]*ligne[7]/1000))
    fret[0].append("FE épandage engrais (kgCO2e/t)")
    fret[0].append("FE EoL tourbe (kgCO2e/t)")
    fret[0].append("Total MP (fab+fret+EoL en kgCO2e/t)")
    fret[0].append("BC cumulé avec EoL (fab+fret+EoL en tCO2e)")
    return fret
    