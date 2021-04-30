# -*- coding: utf-8 -*-
"""
Created on Wed Oct 21 10:45:35 2020

@author: lou


Ensemble des fonctions relatives au calcul du bilan carbone du fret des matières en amont des usines Florentaise. Contient:
    lire_FEtransport()
    calc_fret_par_MP()
    lireImportMaritime()
    lireImportTerrestre()
    calculer_FE_par_t()
    renverserMP_usine_fret()
"""
import parametres as p
import fonctionsMatrices
import affichageResultats
import geolocalisation
import math
import xlrd
import inspect
import sys

#Permet de retourner les valeurs des facteurs d'émission du transport
def lire_FEtransport():
     try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
     except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
         
     feuille = document.sheet_by_index(2)
     type_camion = feuille.cell_value(7,4)
     type_bateau = feuille.cell_value(8,4)
     
     FE_transports = []
     
     for ligne in range(12,31):
         if (feuille.cell_value(ligne,3)!=""):  #On ne prend pas les lignes vides
             FE = []
             for col in [2,3,4,5]:
                 FE.append(feuille.cell_value(ligne,col))
             FE_transports.append(FE)
     FE_route = fonctionsMatrices.recherche_elem(type_camion,FE_transports,1,2) #idem
     FE_bateau = fonctionsMatrices.recherche_elem(type_bateau,FE_transports,1,2) #idem
     return FE_route, FE_bateau
 
#Retourne les valeurs, pour chaque matière première de
#   ->La MP en question  /  la somme de la distance de fret   /  la moyenne de la distance de fret
#   ->la somme de la distance de fret routier de cette MP
#   ->la somme de la distance de fret maritime de cette MP
#   ->la moyenne de la distance de fret routier de cette MP
#   ->la moyenne de la distance de fret maritime de cette MP
#   ->Un tableau contenant ["M3", quantite de M3]["KG", quantite de kg]...etc
#      Attention, ce tableau n'est pas une conversion, mais bien les achats bruts. On a pas encore les densités pour convertir
#Permet également de corriger les erreurs de codes postaux au passage
def calc_fret_par_MP(listeAchatsFret:list, masse_vol_MP:list, FE_route, FE_bateau, nom_fichier):
    
    trajets = [["nomMP", "nomFournisseur","Pays", "depart", "arrivee", "distance route (km)","distance bateau (km)","qte",'unit',"depot_arrive"]]
    Stat = [0,0]
    listeAchats = listeAchatsFret[1:] 
    fretMP = []
    
    
    import_maritime = lireImportMaritime()
    import_terrestre = lireImportTerrestre(p.FICHIER_ENTREES_MANUELLES)
    liste_u_achats = fonctionsMatrices.liste_unique(fonctionsMatrices.extraire_colonne_n(1,listeAchats))
    #Parallèlement au calcul, on va enregistrer la distance parcourue pour livrer chaque dépot
    usines = [[x[2],x[1]] for x in p.CP_SITES_FLORENTAISE if x[3]<=p.ANNEE]
    compte_km_total = 0
                            #tonnage route bateau BCroute BCbateau
    resultat_usine_MP  = [[x, [[y, 0,0,0,0,0] for y in usines if y[1]!="Support"]] for x in liste_u_achats]
    resultat_usine_MP .insert(0,["MP", "id usine (CP/Nom)", "tonnage", "km route", "km bateau", "BC route (tco2e)", "BC bateau (tco2e)"])
    #La structure est un peu particuliere là. Pour chaque matière premiere, resultat_usine_MP a pour ligne
    #  [nom MP, [Usine depot,[%de la masse cette MP livree a ce depot, % du volume de...],fret route, fret bateau]
    
    for mp in resultat_usine_MP [1:]:
        nombre_livraisons_route_tot = 0   #Pour compter le nombre d'entrées (=livraisons) de cette MP
        nombre_livraisons_bateau_tot = 0
        sommeRoute_tot = 0.0          #Distance cumulée de camion pour les livraisons de cette MP
        sommeBateau_tot = 0.0         #Idem en bateau
        nomMP = mp[0]
        #Pour compter les quantités
        tonnage_tot = 0
        idMP = fonctionsMatrices.recherche_elem(nomMP, listeAchats, 1,0)
        #On va chercher la masse volumique pour déterminer le tonnage
        masse_volumique = 1
        for mv in masse_vol_MP:
            if mv[1].lower() in nomMP.lower() or nomMP.lower() in mv[1].lower():
                try:
                    masse_volumique = float(mv[2])
                except ValueError:
                    masse_volumique = 1
                break
        
        for usine in mp[1]:
            nombre_livraisons_route_usine = 0   #Pour compter le nombre d'entrées (=livraisons) de cette MP
            nombre_livraisons_bateau_usine = 0
            cp_site_depot = usine[0][0] 
            for j in range(len(listeAchats)):   #On parcourt la liste des achats
                depot_arrive = listeAchats[j][7] #Code postal du depot d'arrivee de la matiere premiere pour cette livraison
                
                if cp_site_depot== depot_arrive and listeAchats[j][1]==nomMP:        #Si c'est la MP qui nous intéresse 
                    famille = listeAchats[j][4]
                    
                    nombre_livraisons_route_tot   +=1
                    nombre_livraisons_route_usine +=1
                    quantite = listeAchats[j][9]
                    
                    #On compte la quantité transportée, selon l'unité
                    if listeAchats[j][10]=="m3":
                        m3 = quantite
                        kg = 0
                    elif listeAchats[j][10]=="kg":
                        m3 = 0
                        kg = quantite
                    else:               
                        fonctionsMatrices.print_log_erreur("Unité de l'intrant non reconnue : "+listeAchats[j][10]+" pour "+nomMP, inspect.stack()[0][3])
                    tonnage = (kg/1000)+(m3*masse_volumique)
                    tonnage_tot+=tonnage
                    usine[1]+=tonnage
                    
                    #On récupère le calcul de la distance de fret 
                    distanceRoute, distanceBateau, trajets, nombre_livraisons_bateau_supp = calcul_distance_fret_amont(
                        listeAchats[j], trajets, import_terrestre,import_maritime)
                    compte_km_total += (distanceRoute+distanceBateau)
                    sommeRoute_tot  += distanceRoute
                    sommeBateau_tot += distanceBateau
                    nombre_livraisons_bateau_usine += nombre_livraisons_bateau_supp
                    nombre_livraisons_bateau_tot   += nombre_livraisons_bateau_supp
                    
                    emissionsRoute  = (distanceRoute *FE_route )*tonnage/1000
                    emissionsBateau = (distanceBateau*FE_bateau)*tonnage/1000
                    # print(nomMP+"  "+str(int(emissionsRoute*1000))+"  "+str(int(emissionsBateau*1000)))
                    usine[2]+=distanceRoute
                    usine[3]+=distanceBateau
                    usine[4]+=emissionsRoute
                    usine[5]+=emissionsBateau
                    if "tourbe" in nomMP.lower():
                        p.COMPARAISON_TOURBE[0]+= ((distanceRoute*FE_route+distanceBateau*FE_bateau)*tonnage)/1000
                    
            # print(nomMP+"  "+str(cp_site_depot)+"  "+str(usine[2])+"km  "+str(usine[4])+"tCO2e" )
        if nombre_livraisons_route_tot  !=0:
           moyenneRoute_tot = sommeRoute_tot/nombre_livraisons_route_tot
        else:
            moyenneRoute_tot = 0
        if nombre_livraisons_bateau_tot !=0:
            moyenneBateau_tot = sommeBateau_tot/nombre_livraisons_bateau_tot
        else:
            moyenneBateau_tot = 0
        
        # print("TOT: "+ nomMP+"  "+str(masse_volumique)+"t/m3  " )
        #Associe un FE de transport en kgCO2e/t
        #Le but ici n'est pas d'avoir le BC du fret amont mais seulement un FE par tonne: on ne s'intéresse pas encoe aux quantités.
        #On prend juste (distance cumulée ou moyenne)*(FE du moyen de transport)
        #Comme ce FE est en kgCO2.t-1.km-1 on a bien à la fin des kgCO2.t-1
        fretMP.append([nomMP,idMP,famille, sommeRoute_tot,sommeBateau_tot, 
                       moyenneRoute_tot, moyenneBateau_tot, tonnage_tot,
                       FE_route*sommeRoute_tot, FE_bateau*sommeBateau_tot,
                       FE_route*moyenneRoute_tot, FE_bateau*moyenneBateau_tot,
                       masse_volumique
                       ])
    fretMP.insert(0,["nomMP","ID mp", "Famille BC", "Somme du fret route (km)", "Somme du fret naval (km)",
                     "Moyenne de fret routier (km)", "Moyenne de fret naval (km)", "Tonnage",
                     "FE route(kgCO2e/t) cumulé", "FE bateau(kgCO2e/t) cumulé", 
                     "FE route(kgCO2e/t) moyen","FE bateau(kgCO2e/t) moyen",
                     "Masse volumique (t/m3)"])
    print("Km amont total : "+str(int(compte_km_total))+"km")
    if p.DISPLAY_GRAPH:
        tracer_achats(trajets)
    affichageResultats.sauveTrajetsAmont(trajets, nom_fichier)
    return fretMP, resultat_usine_MP



def calcul_distance_fret_amont(listeAchatsj, trajets, import_terrestre, import_maritime):
    #Codes postaux des sites livrés en tourbe par la route
    nomMP = listeAchatsj[1]
    numFourn  = listeAchatsj[3]
    pays =listeAchatsj[8] 
    CP_depart = listeAchatsj[6]
    depot_arrive = listeAchatsj[7]
    quantite = listeAchatsj[9]
    unit = listeAchatsj[10]
    nom_fourn = listeAchatsj[5]
    nombre_livraisons_bateau = 0
    
    #Les achats de tourbes qui proviennent d'ESTONIE mais qui transitent 
    #par Eurotourbes et le port de Montoir sont à corriger
    if nom_fourn.lower() == "eurotourbes":
        listeAchatsj[8] = "EST"
        pays = "EST"
        
        
    if pays.startswith("FRA"):  #Si c'est de France
        distanceBateau = 0      #C'est pas en bateau
        #On calcule par GPS la distance de route entre point de départ et dépot de livraison
        distanceRoute = geolocalisation.distance_CP(CP_depart,depot_arrive)
        
        if distanceRoute[1]==False:  #Si pas d'erreur dans le calcul de la route
            distanceRoute = distanceRoute[0]
        else:
            distanceRoute = geolocalisation.distance_CP(int(CP_depart/10)*10,depot_arrive)
            if distanceRoute[1]==False:  #Si pas d'erreur dans le calcul de la route
                distanceRoute = distanceRoute[0]
            else:
                distanceRoute = geolocalisation.distance_CP(int(CP_depart/100)*100,depot_arrive)
                if distanceRoute[1]==False:  #Si pas d'erreur dans le calcul de la route
                    distanceRoute = distanceRoute[0]
                else:
                    fonctionsMatrices.print_log_erreur("Code postal FR introuvable dans BDDgeoloc "+str(CP_depart), inspect.stack()[0][3])
                    p.LISTE_CODES_POSTAUX_ERRONES.append([distanceRoute[1], "Achats"])
                    distanceRoute = 0
        try:
            trajets.append([nomMP, nom_fourn,pays,geolocalisation.getGPSfromCP(CP_depart), 
                            geolocalisation.getGPSfromCP(depot_arrive), distanceRoute, 0,quantite, unit,depot_arrive])
        except ValueError:
            print("ValEror ")
    else: #Sinon ça vient de l'étranger: exception: si un tourbe est livrée dans certains dépots, c'est par camion
        if pays in ["ALL","BEL","ITA","PB", "NL","ESP", "LUX"] or numFourn == 500094:
            #Si ça provient de pays proches, pas de bateau
            distanceBateau = 0
            latA = -180
            for terrestre in import_terrestre:
                if nom_fourn.lower() in terrestre[1].lower():
                    longA = terrestre[2]
                    latA = terrestre[3]
                    break
            if latA == -180:#C'est signe qu'on a pas trouvé le fournisseur dans la boucle for précédente, on va chercher par pary
                for terrestre in import_terrestre:
                    if pays.lower() in terrestre[0].lower():
                        longA = terrestre[2]
                        latA = terrestre[3]
                        break
            if latA == -180:#Si on a toujours pas trouvé, on le met au milieu de la Lituanie! Centre de l'Europe à peu près
                longA = 25.317*3.1415926/180
                latA = 54.9*3.1415926/180
                fonctionsMatrices.print_log_erreur("Fournisseur européen routier non trouvé: "+nom_fourn, inspect.stack()[0][3])
            distanceRoute = geolocalisation.distance_GPS_CP(longA,latA,depot_arrive)
            if distanceRoute[1]==False:  #Si pas d'erreur dans le calcul de la route
                distanceRoute = distanceRoute[0]
            else:
                distanceRoute = geolocalisation.distance_GPS_CP(longA,latA,int(depot_arrive/100)*100)
                if distanceRoute[1]==False:  #Si pas d'erreur dans le calcul de la route
                    distanceRoute = distanceRoute[0]
                else:
                    p.LISTE_CODES_POSTAUX_ERRONES.append([distanceRoute[1], "Achats"])
                    distanceRoute = 0 
            trajets.append([nomMP, nom_fourn,pays,[longA, latA], 
                            geolocalisation.getGPSfromCP(depot_arrive), distanceRoute,0, quantite, unit,depot_arrive])
        #=======BATEAU===
        else: #Sinon ça prend le bateau
            nombre_livraisons_bateau +=1
            trouve = False
            lieux_de_provenance = []
            for lieux in import_maritime:
                if nom_fourn.lower() in lieux[1].lower(): #On essaye de retrouver par fournisseur d'abord
                    #On cherche avec tout en minuscule pour éviter les pb de casse
                    trouve= True
                    lieux_de_provenance = lieux
                    distanceRoute = float(lieux[2])
                    distanceBateau = float(lieux[3])
            if not trouve:      #Si on a pas réussi à trouver par fournisseur, on regarde comment on recoit en général de ce pays
                for lieux in import_maritime:
                    if pays.lower() in lieux[0].lower(): #On essaye de retrouver par fournisseur d'abord
                        trouve=True
                        lieux_de_provenance = lieux
                        distanceRoute = float(lieux[2])
                        distanceBateau = float(lieux[3])
            if not trouve:
                fonctionsMatrices.print_log_erreur("La provenance "+pays+"/"+nom_fourn+" pour "+nomMP +" n'a pas pu être trouvée", inspect.stack()[0][3])
                distanceRoute = 0
                distanceBateau = 0
            else: #Si on a bien trouvé d'où il venait 
                #Reste à calculer l'itinéraire du port d'arrivée à l'usine d'arrivée
                longPort =lieux_de_provenance[4][0]
                latPort = lieux_de_provenance[4][1] 
                calcul_distance = geolocalisation.distance_GPS_CP(longPort,latPort,depot_arrive)
                if not calcul_distance[1]: #Si pas d'erreur dans le calcul de distance:
                    distanceRoute_supplementaire = p.facteur_route*calcul_distance[0]
                    distanceRoute += distanceRoute_supplementaire
                    trajets.append([nomMP, nom_fourn,"bateau",[ longPort,latPort], 
                            geolocalisation.getGPSfromCP(depot_arrive), distanceRoute,distanceBateau, quantite, unit,depot_arrive])
                    
                else:
                    fonctionsMatrices.print_log_erreur("Le dépot d'arrivée "+depot_arrive+" n'a pas été trouvé pour "+nomMP, inspect.stack()[0][3])
                    p.LISTE_CODES_POSTAUX_ERRONES.append([distanceRoute[1], "Achats"])
            # print(pays+"/"+nom_fourn+" Bateau : "+str(distanceBateau)+"km  |  Route : "+str(distanceRoute)+"km")       
    return distanceRoute, distanceBateau, trajets, nombre_livraisons_bateau
    
#Permet d'aller lire dans la feuille 4 des entrees manuelles les infos maritime
#Renvoie pour chaque usine fournisseur une ligne
# [nom du lieu, nom fournisseur, distance route/port de départ, distance en bateau jusqu'au port d'arrivée]
def lireImportMaritime():
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
         
    feuille_bdd_mar = document.sheet_by_index(3)
   
    #Récupération des données du tableau
    lieux =[]
    for lin in range(2,20):
        if feuille_bdd_mar.cell_value(lin,2)!="":
            ligne = []
            for col in [1,2,3,6,7]:
                ligne.append(feuille_bdd_mar.cell_value(lin,col))
            lieux.append(ligne)
    liens_livr = []
    for lin in range(3,19):
        if feuille_bdd_mar.cell_value(lin,14)!="":
            ligne =[]
            for col in [14,15,16,17]:
                ligne.append(feuille_bdd_mar.cell_value(lin,col))
            liens_livr.append(ligne)    
        
    coordPort_arrive = []
    #route entre usine et port
    
    depart_usine = []
    for chaine in liens_livr:
        if chaine[0]>0:
            longA  = float(fonctionsMatrices.recherche_elem(chaine[0],lieux,0,3))
            latA = float(fonctionsMatrices.recherche_elem(chaine[0],lieux,0,4))
            longB  = float(fonctionsMatrices.recherche_elem(chaine[1],lieux,0,3))
            latB = float(fonctionsMatrices.recherche_elem(chaine[1],lieux,0,4))
            distance_VO_usine_port = math.acos(math.sin(latA)*math.sin(latB)+math.cos(latA)*math.cos(latB)*math.cos(longB-longA))*6371
            distance_route_usine_port =int( p.facteur_route*distance_VO_usine_port)
            #Donc on a la route entre l'usine et le port de départ
            #La distance en bateau est déjà inscrite en dur
            distance_bateau = int(chaine[3])
            usine_loc = fonctionsMatrices.recherche_elem(chaine[0],lieux,0,1)   #Localité de l'usine (ville, pays, ...)
            usine_nom = fonctionsMatrices.recherche_elem(chaine[0],lieux,0,2)   #Nom de l'usine
            longPA = fonctionsMatrices.recherche_elem(chaine[2],lieux,0,3)   #GPS du port d'arrivee
            latPA = fonctionsMatrices.recherche_elem(chaine[2],lieux,0,4)
            coordPort_arrive = [longPA,latPA]
            depart_usine.append([usine_loc, usine_nom,distance_route_usine_port,distance_bateau,coordPort_arrive])
    return depart_usine

def lireImportTerrestre(fichier:str):
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
         
    feuille_bdd_ter = document.sheet_by_index(3)
    fournisseurs =[]
    for lin in range(20,64):
        if feuille_bdd_ter.cell_value(lin,2)!="":
            ligne = []
            for col in [2,3,6,7]:
                ligne.append(feuille_bdd_ter.cell_value(lin,col))
            fournisseurs.append(ligne)
    return fournisseurs

   
def renverserMP_usine_fret(resultat_usine_MP:list):    
    #[[id usine],masse, kmroute, kmbateau, BCroute, BCBateau]
    resultat = [[x[0][1],0,0,0,0,0] for x in resultat_usine_MP[-1][1]]

    for mp in resultat_usine_MP[1:]:
        usines = mp[1]
        nomMP = mp[0]
        
        for us in usines:
            depot_nom = us[0][1]
            tonnage = us[1]
            bc_route = us[4]
            bc_bateau = us[5]
            
            for i in resultat:
                if i[0].lower()==depot_nom.lower():
                    i[1] += tonnage
                    i[2] += 0
                    i[3] += 0
                    i[4] += bc_route
                    i[5] += bc_bateau
                    break
    return resultat

def tracer_achats(trajets):
    import matplotlib.pyplot as plt
    if True:
        plt.figure()
        grossissement = 1
        plt.figure(figsize=(18*grossissement,18*grossissement))
        plt.axis("off")
        plt.title('Achats [rouge:français; vert:maritime; bleu:européen]')
        for tra in trajets[1:]:
            
            if "fra" in tra[2].lower():
                plt.plot([tra[3][0], tra[4][0]], [tra[3][1], tra[4][1]], '-r',linewidth = grossissement*0.3, zorder = 1, alpha = 0.6)#lw = 100/(tra[3]+1)
            elif tra[0]=="bateau":
                plt.plot([tra[3][0], tra[4][0]], [tra[3][1], tra[4][1]], '-g',linewidth = grossissement*0.3, zorder = 1, alpha = 0.6)#lw = 100/(tra[3]+1)
            else:
                plt.plot([tra[3][0], tra[4][0]], [tra[3][1], tra[4][1]], '-b',linewidth = grossissement*0.3, zorder = 1, alpha = 0.6)#lw = 100/(tra[3]+1)
            
        i = 0
        for ville in p.BDDgeoloc:
            i +=1
            if i %20==0:
                plt.scatter(ville[1], ville[2], s = grossissement*20, zorder = 0, c = 000000, alpha = 0.5)
        plt.ylim([0.725,0.92])
        plt.xlim([-0.11,0.22])
        # plt.set_aspect('equal', adjustable = 'box')
        plt.savefig('Tracé des achats.png')
        plt.show()
    if True:
        plt.figure()
        absi= ["<dist"]
        ordo = ["taux"]
        for dis in range(0,2000,50):
            tx_achats_a_moins_de_Xkm = 0
            somme = 0
            for tra in trajets[1:]:
                latA =  tra[3][0]
                longA = tra[3][1]
                latB =  tra[4][0]
                longB = tra[4][1]
                distance = math.acos(math.sin(latA)*math.sin(latB)+math.cos(latA)*
                      math.cos(latB)*math.cos(longB-longA))*6371*p.facteur_route
                somme +=1
                if distance<dis:
                    tx_achats_a_moins_de_Xkm+=1
            absi.append(dis)
            ordo.append(float(100*tx_achats_a_moins_de_Xkm)/somme)
        plt.ylabel("% de livraisons")
        plt.xlabel("km")
        
        plt.plot(absi[1:], ordo[1:], '-g')
        plt.grid(True)
        plt.xlim(xmin=0, xmax=2000)
        plt.ylim(ymin=0, ymax=100)
        plt.title("Taux d'achats à moins de x km en "+str(p.ANNEE))
        plt.show()
    return None


