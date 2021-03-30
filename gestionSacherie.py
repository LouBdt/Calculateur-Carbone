# -*- coding: utf-8 -*-
"""
Created on Wed Oct 28 09:03:36 2020

@author: lou
"""

import parametres as p
import xlrd
import fonctionsMatrices
import geolocalisation
import inspect
import sys

def lireConsoSacherie():
    try:
         document = xlrd.open_workbook(p.DIRECTION_FICHIER_SACHERIE)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des consos sacherie n'est pas trouvé à l'emplacement "+p.DIRECTION_FICHIER_SACHERIE, inspect.stack()[0][3])
         sys.exit("Le fichier des consos sacherie n'est pas trouvé à l'emplacement "+p.DIRECTION_FICHIER_SACHERIE)
    feuille_bdd = document.sheet_by_index(0)
    nbrows = feuille_bdd.nrows
    listeRefs = []
    for lin in range(1,nbrows):
        ref = feuille_bdd.cell_value(lin,4)
        if ref not in listeRefs:
            listeRefs.append(ref)
    liste_depots = fonctionsMatrices.extraire_colonne_n(0,p.CP_SITES_FLORENTAISE)
    resultat = [[x, [[y, 0] for y in liste_depots]] for x in listeRefs]

    for i in range(len(resultat)):
        for lin in range(1,nbrows):
            if feuille_bdd.cell_value(lin,4)==listeRefs[i]:
                depot = feuille_bdd.cell_value(lin,0)
                for dep in resultat[i][1]:
                    if dep[0]==depot:
                        dep[1]+= int(feuille_bdd.cell_value(lin,2))   #Apparemment il y des commandes de nombre pas entiers de sacs...
    
    return resultat

def lireFE_sacherie():
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
    feuille = document.sheet_by_index(0)
    liste_mat = ["pebd", "pet"]
    carac = ["vierge", "neuf", "recyclé", "recycle"]
    for i in range(12,40):
        if feuille.cell_value(i,13)!="":
            for l in liste_mat:
                if l in feuille.cell_value(i,13).lower():
                    mat = l
                    break
            for c in carac:
                if c in feuille.cell_value(i,13).lower():
                    car = c
        if mat=="pebd" and car in ["neuf", "vierge"]:
            p.FE_PEBD_vierge = feuille.cell_value(i,14)
        elif mat=="pebd" and car in ["recyclé", "recycle"]:
            p.FE_PEBD_recy = feuille.cell_value(i,14)
        elif mat=="pet" and car in ["neuf", "vierge"]:
            p.FE_PET_vierge = feuille.cell_value(i,14)
        elif mat=="pet" and car in ["recyclé", "recycle"]:
            p.FE_PET_recy = feuille.cell_value(i,14)
    
    
    return None

 
    
def qte_materiaux_sacherie(refs_sacherie:list):
    #Lecture des dimensions de sacherie
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
    
    feuille_bdd = document.sheet_by_index(5)
    nbrows = feuille_bdd.nrows
    reference_materiaux = []
    #On va lire les caractéristiques de matériau, recyclage et masse de chaque réf
    for lin in range(10,nbrows): #On lit touuuutes les refs, peu importe l'année
        ref = feuille_bdd.cell_value(lin,4)
        mat = feuille_bdd.cell_value(lin,10)
        annee = int(feuille_bdd.cell_value(lin,1))
        imprimeur = feuille_bdd.cell_value(lin,14)
        if mat =="PEBD":
            masse = feuille_bdd.cell_value(lin,12)
            recy = feuille_bdd.cell_value(lin,11)
        elif mat =="Papier":
            masse = feuille_bdd.cell_value(lin,13)
            recy = 0    #Eventuellement à faire évoluer
        else:
            print("Matériaux du sac pas ok")
            mat = "Papier"
            masse = 0
            recy = 0
        reference_materiaux.append([ref, mat, masse, recy, imprimeur, annee])

    liste_depots = fonctionsMatrices.extraire_colonne_n(0,p.CP_SITES_FLORENTAISE)
    conso_materiaux_par_site =  [[x, [["PEBD",0], ["PEBDr",0],["Papier",0]], 0] for x in liste_depots] #Matrice résultat
    #Pour chaque usine on a :
        # [Nom , [["PEBD",massePEBD], ..], fret total pour cette usine]
       
    fret_sacherie_par_usine = [[x, 0, 0] for x in liste_depots] #Matrice résultat, avec [usine;distance; distance.poids] 
        
         #Okay attention là ça devient un peu funky ~~
    #Ca c'est la liste des imprimeurs qu'on trouve dans la bdd de sacherie
    imprimeurs = fonctionsMatrices.liste_unique(fonctionsMatrices.extraire_colonne_n(4,reference_materiaux))
    #D'abord on va faire les correspondances entre imprimeur et code postal 
    imprimeries = []
    for ligne in range(9,35):
        if feuille_bdd.cell_value(ligne,18)!="":
            nom =   feuille_bdd.cell_value(ligne,16).lower()
            long =  feuille_bdd.cell_value(ligne,18)
            lat =  feuille_bdd.cell_value(ligne,19)
            imprimeries.append([nom, lat, long]) 
    
    trajet_total = [0, 0]   #Contiendra la distance et la masse totale transportée
    #Déjà, on inverse la liste des références de sacherie, ça sera utile pour la parcourir dans la deuxième boucle for dans l'ordre de la plus récente à la plus ancienne ref
    reference_materiaux = [reference_materiaux[-(i+1)] for i in range(len(reference_materiaux)) ]
    for i in range(len(refs_sacherie)):   #On parcourt la liste des sacheries avec les quantités achetées par chaque site
        trouveref = False
        trouveimp = False
        longA = -0.0240728298345268 
        latA =0.826995076124778
        ref = "".join([c for c in refs_sacherie[i][0] if c.isalpha() or c.isdigit()])
        if "bigba" in ref.lower():
            trouveref, mat, masse, recy, imprimeur = gestionBigBag(ref)
        else:
            for elem in reference_materiaux:    #On parcourt la liste des sacheries avec les masses de chaque matériau
            #On la parcourt à l'envers, du bas vers le haut par rapport au tableau excel source. Ainsi, grâce au break, on tombera sur la référence la plus récente (sans avoir à comparer les années)
                ref_compare = "".join([c for c in elem[0] if c.isalpha() or c.isdigit()])    
                if ref_compare.lower()==ref.lower() and elem[5]<=p.ANNEE:    #Si on a une correspondance entre les références de sacherie
                    mat = elem[1]           #On enregistre le nom du matériau du sac (PEBD ou Papier)
                    masse = float(elem[2])    #La masse est en grammes
                    recy = float(elem[3])   #Son taux de recyclage (pour l'instant utile que si mat = PEBD)
                    imprimeur = elem[4]
                    trouveref = True
                    trouveimp= False
                    nom_imp_simp= []
                    for c in imprimeur:
                        if c.isalpha():
                            nom_imp_simp+=c.lower()
                    if nom_imp_simp=="" or nom_imp_simp==" " :
                        nom_imp_simp="placel"
                    for imp in imprimeries:
                        autre_nom_imp_simp = []
                        for c in imp[0]:
                            if c.isalpha():
                                autre_nom_imp_simp+=c.lower()
                        if autre_nom_imp_simp == nom_imp_simp:
                            trouveimp = True
                            longA = imp[1]
                            latA = imp[2]
                            break
                    if not trouveimp:
                        fonctionsMatrices.print_log_erreur(
                            "Il faut renseigner les coordonnes de l'imprimeur " + imprimeur
                                                           +" (réf. "+ref+")", inspect.stack()[0][3])
                    break       #Une fois trouvée, on arrete de parcourir la liste
        
        #Selon la matière on donne la masse de chaque composant
        if mat =="PEBD":
            massePEBD_r = recy*masse
            massePEBD =  masse-massePEBD_r
            massePap = 0
        elif mat=="Papier":
            massePEBD = 0
            massePEBD_r = 0
            massePap = masse
        if not trouveref:
            p.Sacherie_manquante.append(refs_sacherie[i][0])
            fonctionsMatrices.print_log_erreur(
                "La référence de sacherie " + refs_sacherie[i][0] +" n'a pas été trouvée", inspect.stack()[0][3])
        
        
        #On parcout la liste des usines dans lesquelles cette référence est livrée
        for site in refs_sacherie[i][1]:
            num_site = site[0]
            qte = int(site[1])                      #Quantité de cette référence livrée dans ce dépot
            
            try:
                cp_site_liv = fonctionsMatrices.recherche_elem(num_site, p.CP_SITES_FLORENTAISE,0,2)
            except ValueError:
                cp_site_liv = 44000
                fonctionsMatrices.print_log_erreur("Code postal du site " 
                                                   + str(num_site)+" introuvable", inspect.stack()[0][3])
            
            trajet = geolocalisation.distance_GPS_CP(longA,latA, cp_site_liv)

            if not trajet[1]:                       #Pour enregistrer le fret de sacherie par usine
                trajet_total[0]+= trajet[0]
                
                trajet_total[1]+= masse
            for usine in fret_sacherie_par_usine:
                if usine[0]==num_site:
                    usine[1] += trajet[0]
                    usine[2] += trajet[0]*qte*masse/(1000*1000*1000) #de grammes en tonnes et de kgCO2e en tCO2e
                    break
            
            # print("Masse : "+str(masse)+"g, qté: "+str(qte))
            for res in conso_materiaux_par_site:    #On parcourt les cases de la matrice résultat
                if res[0]==num_site:                #Si on a trouvé le site correspondant
                    res[1][0][1] += massePEBD*qte/1000000   #On incrémente la quantité de matière en ajoutant la masse de matière de la réf multipliée par la quantité livrée
                    res[1][1][1] += massePEBD_r*qte/1000000 #On convertit en tonnes au passage
                    res[1][2][1] += massePap*qte/1000000
                    if not trajet[1]:
                        res[2]+=trajet[0]
                    break #Une fois le site correspondant trouvé on arrête de chercher       

    return conso_materiaux_par_site, fret_sacherie_par_usine

def gestionBigBag(ref:str):
    mat = "PEBD"
    recy = 0  #Ici on règle le taux de recyclage éventuel des bigbags. On peut faire une fonction de l'année (p.ANNEE)
    imprimeur = "Bigbag"
    
    FE_pet = recy*p.FE_PET_recy+(1-recy)*p.FE_PET_vierge
    FE_pebd= recy*p.FE_PEBD_recy+(1-recy)*p.FE_PEBD_vierge
    #C'est ce rapport avec lequel il faut multiplier la masse de PET pour avoir la masse de PEBD à impact climatique équivalent :
    rapport_PET_PEBD = FE_pet/FE_pebd
    try:
        litrage = int("".join([c for c in ref if c.isdigit()]))
    except ValueError:
        return False, mat, 0, recy, imprimeur
    
    # On considère un cube avec seulement 5 faces en PET, dont le volume vaut litrage/1000
    if litrage!=0:
        surface = 5*((litrage/1000)**(2/3)) #en m3
        
    else:
        litrage = 1000
        
    masse_PET = surface*p.MASSE_SURFACIQUE_PET/1000000
    masse_PEBD_eq = rapport_PET_PEBD*masse_PET
    return True, mat, masse_PEBD_eq, recy, imprimeur

def lireFE_emballage():
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
    
    feuille_bdd = document.sheet_by_index(2)
    res = []
    for ligne in range(12,40):
        if feuille_bdd.cell_value(ligne,13)!="":
            res.append([feuille_bdd.cell_value(ligne,13), feuille_bdd.cell_value(ligne,14)])
    return res



def BC_sacherie(fret_sacherie_par_usine:list, conso_materiaux_par_site:list, FE_route:float, FE_emballage:list):
    BC_fret_sacherie = [[[x[0],x[2],x[1]],0] for x in p.CP_SITES_FLORENTAISE]
    BC_fret_sacherie.insert(0, ["Site", "BC du fret sacherie"])
    for fret in fret_sacherie_par_usine:
        for usine in BC_fret_sacherie[1:]:
            if fret[0]==usine[0][0]:
                usine[1]+=fret[2]*FE_route  #fret[2] c'est en km.t
    
                
    BC_MP_sacherie = [[[x[0],x[2],x[1]],[0, 0, 0, 0, 0, 0]] for x in p.CP_SITES_FLORENTAISE]
    BC_MP_sacherie.insert(0, ["Site", ["BC PEBD", "BC PEBDr", "BC papier", 
                                       "BC PVC", "BC carton", "BC autre"]])
    for site in conso_materiaux_par_site:
        site_num = site[0]
        for usine in BC_MP_sacherie[1:]:
            if site_num==usine[0][0]:
                for mat in site[1]:
                    materiau = mat[0]
                    qte = mat[1]
                    FE = 0
                    for fact in FE_emballage:
                        if materiau in fact[0]:
                            FE = fact[1]
                            break
                    if materiau=="PEBD": #Les quantités sont en tonnes et le FE en kgCO2e/t donc on convertit le produit en tCO2e
                        usine[1][0] += FE*qte/1000
                    elif materiau=="PEBDr":
                        usine[1][1] += FE*qte/1000
                    elif materiau=="Papier":
                        usine[1][2] += FE*qte/1000
                    elif materiau=="PVC":
                        usine[1][3] += FE*qte/1000
                    elif materiau=="Carton":
                        usine[1][4] += FE*qte/1000
                    else:
                        usine[1][5] += FE*qte/1000
                break
                    
    return BC_fret_sacherie, BC_MP_sacherie
       
   
def calculBC_fret_sacherie(fret_sacherie:list, FE_route:float):
    return FE_route*fret_sacherie[0]*fret_sacherie[1]/1000