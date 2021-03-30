# -*- coding: utf-8 -*-
"""
Created on Wed Dec  9 13:33:21 2020

@author: lou
"""

import time
import fonctionsMatrices
def init():
    
    #Les données à récupérer dans les paramètres de calcul (Excel)
    global retrieve
    retrieve = ["ANNEE", "CHEMIN_ECRITURE_RESULTATS", "FICHIER_ENTREES_MANUELLES"]
    retrieve += ["FICHIER_BDD_GEOLOC","DIRECTION_FICHIER_INTRANTS", "DIRECTION_FICHIER_VENTES"]
    retrieve += ["DIRECTION_FICHIER_SACHERIE", "DIRECTION_ERREURS_CP"]
    retrieve += ["ELECTRICITE_ET_AUTRES", "INTRANTS_ET_FRET", "EMBALLAGES_ET_SACHERIE", "FRET_AVAL"]
    retrieve = set(retrieve)
    
    global ANNEE; ANNEE = 2020
    
    #Directions de lecture et d'écriture
    global CHEMIN_ECRITURE_RESULTATS; CHEMIN_ECRITURE_RESULTATS  = "D:\\Resultat/"
    global FICHIER_ENTREES_MANUELLES; FICHIER_ENTREES_MANUELLES = 'Sources/Entrees Manuelles.xlsx'
    global FICHIER_BDD_GEOLOC; FICHIER_BDD_GEOLOC = 'Sources/BDDgeoloc.xlsx'
    global DIRECTION_FICHIER_INTRANTS; DIRECTION_FICHIER_INTRANTS= 'Donnees/'+str(ANNEE)+'/Intrants '+str(ANNEE)+'.xlsx'
    global DIRECTION_FICHIER_VENTES1; DIRECTION_FICHIER_VENTES1 = 'Donnees/'+str(ANNEE)+'/Ventes1 '+str(ANNEE)+'.xlsx'
    global DIRECTION_FICHIER_VENTES2; DIRECTION_FICHIER_VENTES2 = 'Donnees/'+str(ANNEE)+'/Ventes2 '+str(ANNEE)+'.xlsx'
    global DIRECTION_FICHIER_SACHERIE; DIRECTION_FICHIER_SACHERIE = 'Donnees/'+str(ANNEE)+'/Conso sacherie '+str(ANNEE)+'.xlsx'
    
    
    #Calculer ou non certaines parties
    global ELECTRICITE_ET_AUTRES; ELECTRICITE_ET_AUTRES = True
    global INTRANTS_ET_FRET; INTRANTS_ET_FRET = True
    global EMBALLAGES_ET_SACHERIE; EMBALLAGES_ET_SACHERIE = True
    global FRET_AVAL; FRET_AVAL = True
    
    #Affichage de l'exécution
    global DISPLAY_CONSOLE;DISPLAY_CONSOLE = True
    global DISPLAY_GRAPH;DISPLAY_GRAPH= False
    global DISPLAY_ERROR;DISPLAY_ERROR= True    
    global DISPLAY_INFO; DISPLAY_INFO= True
    global DISPLAY_LOG;DISPLAY_LOG = True
    global SAVE_LOG; SAVE_LOG = True
    
    #Initialisation de listes
    global MESSAGES_DERREUR; MESSAGES_DERREUR = [["Niveau", "Message", "Fonction"]]
    global LISTE_ASSIMILATIONS; LISTE_ASSIMILATIONS = [["Element Initial", "Element auquel il a été assimilé", "Méthode/raison"]]
    global LISTE_CODES_POSTAUX_ERRONES; LISTE_CODES_POSTAUX_ERRONES = [["Code postal", "Achats/Ventes"]]
    global CONVERSION_SACHERIE_GCO; CONVERSION_SACHERIE_GCO = []
    global Sacherie_manquante; Sacherie_manquante = []
    global PROD_PAR_SITE; PROD_PAR_SITE = []    
    global BDDgeoloc; BDDgeoloc = []
    global CP_SITES_FLORENTAISE; global CP_FLORENTAISE; CP_FLORENTAISE = []
    global MATRICE_RESULTAT
    global COMPARAISON_TOURBE; COMPARAISON_TOURBE = [0,0,0] #fret, MP, EoL
    
        
    #Paramètres techniques par défaut:Paramètres constants de calcul
    global VOLUME_PAR_BIGBAG; VOLUME_PAR_BIGBAG = 2       # en m3
    global MASSE_PAR_PALETTE; MASSE_PAR_PALETTE = 1    # en tonnes
    global tonnage_camion; tonnage_camion = 15 #t
    global DENSITE_TERREAU_PRO_MOYENNE; DENSITE_TERREAU_PRO_MOYENNE = 0.441 #t/m3
    global facteur_route; facteur_route = 1.4
    global FE_photovoltaique; FE_photovoltaique = 0.055 #kgCO2e/kWh
    
    global FE_EMPANDAGE_ENGRAIS; FE_EMPANDAGE_ENGRAIS = 0.021#kgCO2e/kg_azote
    global PRG_CH4; PRG_CH4 = 28
    global PRG_N2O; PRG_N2O  = 265
    
    global FE_PET_vierge; FE_PET_vierge = 3270#kgCO2e/t
    global FE_PET_recy; FE_PET_recy = 202#kgCO2e/t
    global FE_PEBD_vierge; FE_PEBD_vierge = 2090#kgCO2e/t
    global FE_PEBD_recy; FE_PEBD_recy = 202  #kgCO2e/t
    global MASSE_SURFACIQUE_PET; MASSE_SURFACIQUE_PET = 150 #g/m²
    
    
    fonctionsMatrices.charger_parametres()
    
    afficherFleurs()
    
    
    #Chronomètre
    global starting_time; starting_time = time.time()
    #Esthétique Excel
    global STYLES_EXCEL
    

        
def afficherFleurs():
    print("                _")
    print("              _(_)_                 wWWWw   _")
    print("  OOOOO      (_)@(_)    _     @OO@  (___) _(_)_")
    print(" OO(·)OO ___   (_)\   _(_)_  @O()O@   Y  (_)@(_)")
    print("  OOOOO (_|_)    `|/ (_)@(_)  @OO@   \|/   (_)\ ")
    print("   /      Y      \|  / (_)    \|      |/       |")
    print("\\\\ |    \ |/      | \|/        |/    \|       \|/")
    print("\\\\\|//Λ\\\\\|///Λ\\\\\|/\|////Λ\\\\\\\|///Λ\\\|////Λ\\\\\\|////")
    kr = "►◄►◄►◄►◄►♦ Bilan carbone "+str(ANNEE)+" Florentaise ♦◄►◄►◄►◄►◄"
    print("►◄"*int(len(kr)/2));print(kr);print("►◄"*int(len(kr)/2));print()
    
    
    