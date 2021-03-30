    # -*- coding: utf-8 -*-
"""
Created on Fri Oct 23 14:37:00 2020

@author: lou
"""
import parametres as p
import xlrd
import fonctionsMatrices
import geolocalisation
import math
import matplotlib.pyplot as plt
import inspect
import sys

def lireExport():
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
    feuille_bdd = document.sheet_by_index(4)
    nbrows = feuille_bdd.nrows
    #TERRESTRE    
    bdd_exp_terre = []
    for i in range(20,min(42,nbrows)):
        ligne = []
        if feuille_bdd.cell_value(i, 2)!="":
            for j in [1,2,3,6,7]:
                ligne.append(feuille_bdd.cell_value(i, j))
            bdd_exp_terre.append(ligne)
    #MARITIME
    bdd_exp_mari = []   
    liaisons = []
    for i in range(3,11):
        if feuille_bdd.cell_value(i, 14)!="":
            chaine = []
            for j in [14,15,16]:
                chaine.append(feuille_bdd.cell_value(i, j))
            liaisons.append(chaine)

    ports = []
    for i in range(2,17):
        if feuille_bdd.cell_value(i, 2)!="":
            port = []
            for j in [1,2,3,6,7]:
                port.append(feuille_bdd.cell_value(i, j))
            ports.append(port)
            
    for chaine in liaisons:
        port_depart = ports[int(chaine[0]-1)]
        port_arrivee = ports[int(chaine[1]-1)]
        bdd_exp_mari.append([port_depart, port_arrivee, chaine[2]])
    return bdd_exp_terre, bdd_exp_mari

def lire_couts_camion_aval():
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
         
    feuille_bdd = document.sheet_by_index(2)
    #En €/km selon le cout    
    bdd_coutkm_E = []
    for i in range(47,58):
        l = []
        if feuille_bdd.cell_value(i, 3)!="":
            for j in [3,4,5,6]:
                l.append(feuille_bdd.cell_value(i, j))
        bdd_coutkm_E.append(l)
    #En €/km selon la distance
    bdd_coutkm_dist = []
    for i in range(60,70):
        l = []
        if feuille_bdd.cell_value(i, 3)!="":
            for j in [3,4,5]:
                l.append(feuille_bdd.cell_value(i, j))
        bdd_coutkm_dist.append(l)
    return bdd_coutkm_E, bdd_coutkm_dist

def liregroupements():
    try:
         document = xlrd.open_workbook(p.DIRECTION_FICHIER_VENTES1)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des ventes 1 n'est pas trouvé à l'emplacement "+p.DIRECTION_FICHIER_VENTES1, inspect.stack()[0][3])
         sys.exit("Le fichier des ventes 1 n'est pas trouvé à l'emplacement "+p.DIRECTION_FICHIER_VENTES1)
    document = xlrd.open_workbook(p.DIRECTION_FICHIER_VENTES1)
    feuille = document.sheet_by_index(0)
    nbrows = feuille.nrows
    resultat = [ ["N° commande", "N° groupement"]]
    ligne = 1
    while ligne<nbrows:
        num_commande = feuille.cell_value(ligne, 10)
        num_gpt = feuille.cell_value(ligne, 0)
        resultat.append([num_commande,num_gpt])
        ligne +=1
    return resultat
    
def lireDensitesPF():
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
    
    feuille = document.sheet_by_index(7)
    nbrows = feuille.nrows
    resultat = []
    ligne = 2
    while ligne<nbrows:
        
        if feuille.cell_value(ligne, 7)!="" and feuille.cell_value(ligne, 8)!=-1:
            l = []
            l.append(feuille.cell_value(ligne, 7))
            try:
                l.append(abs(float(feuille.cell_value(ligne, 8))))
            except ValueError:
                l.append(0.387)
            resultat.append(l)
        ligne +=1
    return resultat



def lire_ventes(densitesPF:list, groupements:list):
    stat = [0, 0, 0]
    plus_grand_n_regroupement = max([x[1] for x in groupements[1:]]+[1])+1
    #Il y a deux fichiers, on lit d'abord le premier, celui avec les regroupements
    try:
         document = xlrd.open_workbook(p.DIRECTION_FICHIER_VENTES2)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des ventes 1 n'est pas trouvé à l'emplacement "+p.DIRECTION_FICHIER_VENTES2, inspect.stack()[0][3])
         sys.exit("Le fichier des ventes 1 n'est pas trouvé à l'emplacement "+p.DIRECTION_FICHIER_VENTES2)
    document = xlrd.open_workbook(p.DIRECTION_FICHIER_VENTES2)
    feuille = document.sheet_by_index(0)
    nbrows = feuille.nrows
    col_commande = 5
    col_type_client = 8
    col_cp_depart = 4
    col_reference = 13
    col_volume = 16
    col_cp_liv=11
    
    entetes = [col_commande,col_type_client,col_cp_depart, col_cp_liv, col_reference]
    result = [["Numéro de commande","Type de livraison", "CP depot depart","CP livraison","Référence",
               "Num groupage", "Tonnage"]]
    
    col_qte = 15
    col_volume = 16
    col_unit = 17
    
    ligne= 1
    while ligne <nbrows:   
        cp_depot = feuille.cell_value(ligne, col_cp_depart)
        if cp_depot not in [64260, 69150, 68108]:
            lres = []
            for col in entetes:
                lres.append(feuille.cell_value(ligne, col))
            num_groupage, plus_grand_n_regroupement = recherche_num_groupage(lres, groupements, plus_grand_n_regroupement)
            lres.append(num_groupage)
            qte_vol_unit = [feuille.cell_value(ligne, j) for j in [col_qte, col_volume,col_unit ] ]
            tonnage = tonnage_par_ref(lres, densitesPF, qte_vol_unit)
            lres.append(tonnage)
            result.append(lres)
        ligne +=1
        
        
    for li in result:
        if li[1].lower() in ["cla", "clb", "anj"]:
            stat[0] +=1
        elif li[1] =="INT":
            stat[1] +=1
        else:
            stat[2]+=1
    print('GP/INT/PRO:')
    print(stat)
    result = regrouper_par_vente(result)
    stat= [0,0,0]
    for li in result:
        if li[1].lower() in ["cla", "clb", "anj"]:
            stat[0] +=1
        elif li[1] =="INT":
            stat[1] +=1
        else:
            stat[2]+=1
    print(stat)
    return result

def recherche_num_groupage(lres, groupements, n):
    commande,type_client,cp_depart,cp_arrivee,reference = lres
    try:
        num_gpt = fonctionsMatrices.recherche_elem(commande, groupements[1:], 0,1)
        return num_gpt, n
    except ValueError:
        n+=1
        return n,n
        

def tonnage_par_ref(lres, densitesPF,qte_vol_unit ):
    commande,type_client,cp_depart,cp_arrivee,reference, num_gpt = lres
    quantite, volume, unite = qte_vol_unit
    if unite.lower() in ["kg", "t", "ton"]:
        if unite.lower() =="kg":
            tonnage = quantite/1000
        else:
            tonnage = quantite
    else:
        densite = 0.38687
        for d in densitesPF:
            if d[0].lower() == str(reference).lower():
                densite = d[1]
                break
        tonnage = volume*densite
    return tonnage

def regrouper_par_vente(listeDepart):
    result = []
    dejaVu = []#N° de BL déjà rencontrés, liste remplie en même temps que result aux mêmes endroits
    for ligne in listeDepart[1:]:
        num_commande = ligne[0]
        try:
            i = dejaVu.index(num_commande)
            result[i][-1] += ligne[-1]#On incrémente le tonnage
        except ValueError:
            dejaVu.insert(0, num_commande)
            result.insert(0, ligne)
    result.insert(0,listeDepart[0])
    return result

def regrouper_livraison(listeVentes:list):
    
    liste_groupes_deja_vu = []
    livraisons = []
    for vente in listeVentes[1:]:
        numGPT = vente[5]
        numDEPOT = vente[2]
        typeClient = vente[1]
        try:
            i = liste_groupes_deja_vu.index(numGPT)
            livraisons[i][3].append(vente)
        except ValueError:
            livraisons.append([numGPT, numDEPOT, typeClient, [vente]])
            liste_groupes_deja_vu.append(numGPT)
        
    livraisons.insert(0, ["num groupement", "num depot", "type client", "sous-livraisons"])
    
    stat = [0,0,0]
    for li in livraisons:
        if li[2].lower() in ["cla", "clb", "anj"]:
            stat[0] +=1
        elif li[2] =="INT":
            stat[1] +=1
        else:
            stat[2]+=1
    print(stat)
    return livraisons


def calc_fret_aval(livraisons:list,FE_route:float, FE_bateau:float, bdd_exp_terre:list, bdd_exp_mari:list):
    stat = [0,0,0]
    nlivtot = len(livraisons)-1
    statis = [[l for l in range(5, 200, 5)]+[l for l in range(200, 1251, 50)]]
    statis.append([0 for x in statis[0]])
    trajets = [["pays","depart", "arrivee", "distance (km)","tonnage", "cout carbone (kgCO2e)", 'type']]
    fret_aval = [[x[0],x[1], x[2], x[3], 0, 0] for x in livraisons[1:]]
    lire_conversion_sacherie_GCO()
    #pour chaque marché: distance+distance.t+distance.t.fe
    totGP = [0,0,0]     #Ventes en jardinerie
    totINT = [0,0,0]    #Déplacements internes
    totPRO = [0,0,0]    #Ventes pro
    res = [[[x[0],x[2],x[1]],[0, 0, 0]] for x in p.CP_SITES_FLORENTAISE] #L'élément 2 est le BC pour [GP INT et PRO]

    for liv in fret_aval:
        bilan_carb =0
        trouve = False
        cp_depot = liv[3][0][2]
        cp_livraison = liv[3][0][3]
        tonnage = liv[3][0][6]
        if len(liv[3])>1:
            dist, dist_t, bilan_carb  = optimiser_livraisons(liv, FE_route) #Algo d'estimation des livraisons groupées
           
        #S'il n'y a pas de regroupement et que la livraison se fait en une seule fois
        else:
            coord_arrivee = geolocalisation.getGPSfromCP(cp_livraison)
            distance = geolocalisation.distance_GPS_CP(coord_arrivee[0],coord_arrivee[1],cp_depot)
            if distance[1]==False:
                dist = distance[0]
            else:
                p.LISTE_CODES_POSTAUX_ERRONES.append([distance[1], "Ventes"])
                fonctionsMatrices.print_log_erreur("Erreur pour le code postal "+str(distance[1])+" (livraison n° "+str(liv[0])+" )", inspect.stack()[0][3])
                dist = 100
            dist_t = tonnage*dist #le tonnage multiplié par la distance (t.km)
            
            if dist>2000:#km
                bilan_carb = (FE_bateau*0.95+FE_route*0.05)*dist_t/1000
            else:
                bilan_carb = FE_route*dist_t/1000
            
        #Pour tracer le graphe de la proximité des ventes
        for s in range(len(statis[0])):
            if dist<statis[0][s]:
                statis[1][s] +=1
        #Dans le cas où le bilan carbone est négatif (masse transportée négative) on enregistre la valeur absolue
            #il s'agit d'un retour (donc dans l'autre sens, mais BC toujours positif)
        if bilan_carb<0:
            bilan_carb = abs(bilan_carb)
        
                
        if liv[2].lower() in ["cla", "clb", "anj"]:
            totGP[0]+=dist  #On range le BC par poste
            totGP[1]+=dist_t
            totGP[2]+=bilan_carb
            stat[1]+=1
        
        elif liv[2].lower() in ["int"]:
            if cp_depot not in p.CP_FLORENTAISE:
                bilan_carb = 0
                dist = 0
                dist_t = 0
            else:
                totINT[0]+=dist
                totINT[1]+=dist_t
                totINT[2]+=bilan_carb
                stat[0]+=1        
        else:
            totPRO[0]+=dist
            totPRO[1]+=dist_t
            totPRO[2]+=bilan_carb
            stat[2]+=1
        
        #Correspondance: CLA/CLB/ANJ ==GP ; INT==INT; autre==pro*
        for k in res:
            if k[0][1]==liv[1]:  #On identifie l'usine à laquelle attribuer la vente (id CP)
                if liv[2] in ["CLA", "CLB", "ANJ"]:
                    k[1][0] += bilan_carb   #On range dans le bilan carbone par usine
                elif liv[2] in ["INT"]:
                    k[1][1] += bilan_carb
                else:
                    k[1][2] += bilan_carb
                break
        trajets.append(["FR", geolocalisation.getGPSfromCP(cp_depot),
                        geolocalisation.getGPSfromCP(cp_livraison), dist,liv[3][0][3], bilan_carb, liv[2]])
    
    
    if p.DISPLAY_GRAPH:
        #Tracé de la carte de proximité
        tracer_ventes(trajets, statis, nlivtot)
    return res, [totGP, totINT, totPRO]

def tracer_ventes(trajets, statis, nlivtot):
    import matplotlib.pyplot as plt
    if True:
        #Tracé du graphe de proximité
        plt.figure()
        plt.xlim([0, 1250])
        plt.ylim([0, 100])
        plt.plot([0]+statis[0], [0]+[100*x/nlivtot for x in statis[1]])
        plt.xlabel("Distance (km)")
        plt.grid(True)
        plt.title('Graphe de proximité des ventes en '+str(p.ANNEE))
        plt.ylabel("%age des livraisons en fret aval à moins de X km")
        plt.show()
    if True:
        grossissement = 1
        plt.figure(figsize=(18,18))
        plt.axis("off")
        plt.title('Ventes [rouge: GP, bleu: pro, vert: interdepot]')
        i = 0
        for tra in trajets[1:]:
            i+=1
            if i%1==0:
                if tra[-1] in  ["CLA", "CLB", "ANJ"] or False:
                    plt.plot([tra[1][0], tra[2][0]], [tra[1][1], tra[2][1]], '-r', linewidth =  grossissement*0.5, zorder = 1, alpha = 0.6)
                elif tra[-1] in  ["INT"] or False:
                    plt.plot([tra[1][0], tra[2][0]], [tra[1][1], tra[2][1]], '-g', linewidth =  grossissement*0.5, zorder = 1, alpha = 0.6)
                else:
                    plt.plot([tra[1][0], tra[2][0]], [tra[1][1], tra[2][1]], '-b', linewidth =  grossissement*0.5, zorder = 1, alpha = 0.6)
        i = 0
        for ville in p.BDDgeoloc:
            i +=1
            if i %40==0:
                plt.scatter(ville[1], ville[2], s = grossissement*40, zorder = 0, c = 000000, alpha = 0.5)
        plt.ylim([0.725,0.92])
        plt.xlim([-0.11,0.22])
        # plt.set_aspect('equal', adjustable = 'box')
        plt.savefig('Tracé des ventes.png')
        plt.show()

    
    return None

def lire_conversion_sacherie_GCO():
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
    
    feuille_bdd = document.sheet_by_index(0)
    nb = feuille_bdd.nrows
    p.CONVERSION_SACHERIE_GCO = []
    ligne = 12
    while ligne<nb:
        if feuille_bdd.cell_value(ligne,29)!="":
            p.CONVERSION_SACHERIE_GCO.append([feuille_bdd.cell_value(ligne,29), 
                                              feuille_bdd.cell_value(ligne,30),
                                              feuille_bdd.cell_value(ligne,31)])            
        ligne +=1
    return None

def optimiser_livraisons(liv:list, FE_route):
    
    cp_depot = liv[3][0][2]
    coord = geolocalisation.getGPSfromCP(cp_depot)
    coordonnees_et_masse =[[coord[0], coord[1], 0]]
    n = len(liv[3])
    if n>4 and p.DISPLAY_GRAPH:
        trace = False
    else:
        trace = False
    listeCP_visites = [cp_depot]
    for sous_liv in liv[3]:
        tonnage = sous_liv[6]
        CP_liv = sous_liv[3]
        if tonnage!=0: #On ne compte pas passer par les endroit où rien n'est livré
        
            if CP_liv not in listeCP_visites: #Si on n'a pas déjà à livrer ce CP
                listeCP_visites.append(CP_liv)
                coord = geolocalisation.getGPSfromCP(CP_liv)
                #On stocke pour chaque point Longitude Latitude et Masse à livrer
                coordonnees_et_masse.append([coord[0], coord[1], tonnage])
                #On dit aussi que le premier point part avec la somme de chaque livraison en masse
                coordonnees_et_masse[0][2] += tonnage
            #Si on compte déjà passer par là, on fusionne les deux points: on ne le rajoute pas dans la liste des endroits où passer, mais on assemble les masses
            else:
                coord = geolocalisation.getGPSfromCP(CP_liv)
                #on cherche quel point fusionner
                for dejavu in coordonnees_et_masse:
                    if dejavu[0]==coord[0]and dejavu[1]==coord[1]:
                        dejavu[2]+=tonnage#On augmente simplement la masse
                        coordonnees_et_masse[0][2] += tonnage
                        break
                    
    if len(coordonnees_et_masse)>1:
        dist, dist_t, bilan_carb = parcours(coordonnees_et_masse, 1+n,0.002, trace, FE_route, liv)
        return dist, dist_t, bilan_carb 
    else:
        #Rien n'est livré?
        return 0,0,0
    
def distance(Xa, Ya, Xb, Yb):
    latA, longA = Ya, Xa
    latB, longB = Yb, Xb
    a = math.sin(latA)*math.sin(latB)+(math.cos(latA)*math.cos(latB)*math.cos(longA-longB))

    dist = math.acos(min(a,1))*6371
    dist*=p.facteur_route
    return dist

def parcours(pts:list, n:int, poids:float, trace:bool, FE_route:float, liv):
    tailles_points = 50
    taille_trait = 5    
    if trace and len(pts)>2:
        plt.figure()
        plt.title("Optimisation du groupement de livraison BL n°"+str(int(liv[0]))+" n="+str(len(pts)))
        i = 0
        for pt in pts[1:]:
            i+=1
            plt.scatter(pt[0],pt[1],s= tailles_points*pt[2]/pts[0][2], zorder=2, c="r")
            plt.text(pt[0],pt[1],"  "+str(i))
        plt.scatter(pts[0][0],pts[0][1],s= tailles_points, zorder=3, c="g")
    
    court_chemin = solve_tsp_dynamic(pts,liv)
    if court_chemin != False:
        court_chemin.append(court_chemin[0])#Retour au point de départ
    else: 
        court_chemin= [x for x in range(len(pts))]
    dist = 0
    dist_t = 0
    charge = pts[0][2]
    charge_tot = pts[0][2]
    prec = court_chemin[0]
    for i in court_chemin[1:]:
        if trace and len(pts)>2:
            if taille_trait*charge/charge_tot>1:
                plt.plot([pts[prec][0], pts[i][0]], [pts[prec][1], pts[i][1]], "-y",  linewidth =taille_trait*charge/charge_tot, zorder=0)
            else:
                plt.plot([pts[prec][0], pts[i][0]], [pts[prec][1], pts[i][1]], ":y",  linewidth =1, zorder=0)
        d = distance(pts[prec][0], pts[prec][1],pts[i][0], pts[i][1])
        
        dist+=d
        dist_t += d*charge
        charge -= pts[i][2]
        prec =i
    if trace and len(pts)>2:
        # plt.plot([pts[0][0], pts[court_chemin[-1]][0]], [pts[0][1], pts[court_chemin[-1]][1]], ":y",  linewidth =2, zorder=0)
        plt.show()
    return dist, dist_t, FE_route*dist_t/1000


import itertools
def solve_tsp_dynamic(pts, liv): #https://gist.github.com/mlalevic/6222750
    #calc all lengths
    all_distances = [[distance(pts[x][0], pts[x][1],pts[y][0], pts[y][1]) for y in range(len(pts))] for x in range(len(pts))]
    #initial value - just distance from 0 to every other point + keep the track of edges
    A = {(frozenset([0, idx+1]), idx+1): (dist, [0,idx+1]) for idx,dist in enumerate(all_distances[0][1:])}
    cnt = len(pts)
    for m in range(2, cnt):
        B = {}
        for S in [frozenset(C) | {0} for C in itertools.combinations(range(1, cnt), m)]:
            for j in S - {0}:
                B[(S, j)] = min( [(A[(S-{j},k)][0] + all_distances[k][j], A[(S-{j},k)][1] + [j]) for k in S if k != 0 and k!=j])  #this will use 0th index of tuple for ordering, the same as if key=itemgetter(0) used
        A = B
    try:
        res = min([(A[d][0] + all_distances[0][d[1]], A[d][1]) for d in iter(A)])
        return res[1]   
    except ValueError:
        fonctionsMatrices.print_log_erreur("Optimisation de la sous livraison n°"+str(int(liv[0]))+" impossible", inspect.stack()[0][3])
        i = 0
        plt.title("Problème pour trouver le chemin optimal de BL n°"+str(int(liv[0])))
        for pt in pts[1:]:
            i+=1
            plt.scatter(pt[0],pt[1],s= 50*pt[2]/pts[0][2], zorder=2, c="r")
            plt.text(pt[0],pt[1],"  "+str(i))
        plt.scatter(pts[0][0],pts[0][1],s= 50, zorder=3, c="g")
        plt.show()
        return False