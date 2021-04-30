#       -*- coding: utf-8 -*-
"""
Created on Thu Oct 15 15:38:37 2020

@author: lou

Ce programme est destiné à calculer le bilan carbone annuel dû aux activités de Florentaise. Il se sert pour cela de la lecture de fichiers issus
d'extractions du logiciel de GCO. Le fonctionnement du programme est détaillé dans le guide "Guide Technique du programme python de traitement du
bilan carbone" dans J:\QUALITE-SECURITE-ENVIRONNEMENT\bilan CO2\Documents Lou\Partage des résultats\Calculateur Carbone.

Un autre document, moins axé sur le code mais davantage sur la méthodologie de bilan carbone est disponible dans le même emplacement. Il justifie
les choix, hypothèse et facteurs d'émissions utilisés ici.


pour compiler:
pyinstaller --onefile -i Logo.ico main.py
"""


import time
import parametres as p
p.init()

print("[{t:6.2f}] ►Import des modules...".format(t =time.time()-p.starting_time))

import fonctionsMatrices
import geolocalisation
import gestionIntrants
import affichageResultats
import fonctionsFret
import gestionExport
import gestionSacherie
import resultat
import autres





def main():
    print("[{t:6.2f}] ►Préparation des bases de données...".format(
        t =time.time()-p.starting_time))
    
    #Préparation du tableau de résultats généraux
    nom_fichier_resultat= affichageResultats.creationFichierResultat()
    p.CP_SITES_FLORENTAISE = geolocalisation.get_cp_sites()
    CA_flor, employ_par_site, p.PROD_PAR_SITE= autres.lire_donnees_entreprise()
    
    affichageResultats.sauve_resultats_generaux(nom_fichier_resultat)
    
    p.BDDgeoloc = geolocalisation.lireBDDgeoloc()
    
    
    #On va lire les valeurs des FE pour le transport
    FE_route, FE_bateau = fonctionsFret.lire_FEtransport()
    
    production_par_site = p.PROD_PAR_SITE
    p.MATRICE_RESULTAT = resultat.creerMatriceResultat()
    
    #===================================================#
    #============== ELECTRICITE ET AUTRES ==============#
    #===================================================#
    if p.ELECTRICITE_ET_AUTRES:
        print("[{t:6.2f}] ►Lecture des consommations énergétiques...".format(
            t =time.time()-p.starting_time))
        FE_elect = autres.lire_FE_elec()
        conso_elec_sites,conso_elec_totale = autres.lire_conso_elec()
        FE_fuel = autres.lire_FE_fuel()
        conso_fuel_sites,conso_fuel_totale = autres.lire_conso_fuel()
        
        print("[{t:6.2f}] ►Calcul des bilan carbones liés à l'énergie...".format(
            t =time.time()-p.starting_time))
        BC_elec_par_site = autres.calc_BC_elec(FE_elect,conso_elec_sites)
        BC_fuel_par_site = autres.calc_BC_elec(FE_fuel,conso_fuel_sites)
        
        print("[{t:6.2f}] ►Enregistrement des résultats sur l'énergie...".format(
            t =time.time()-p.starting_time))
        resultat.enregistre_BC_elec(BC_elec_par_site,production_par_site)
        resultat.enregistre_BC_fuel(BC_fuel_par_site,production_par_site)
        
        print("[{t:6.2f}] ►Lecture & enregistrement des immobilisations...".format(
            t =time.time()-p.starting_time))
        immobilisations = autres.ajouter_immobilisations()
        resultat.enregistre_immobilisations(immobilisations,production_par_site)
        del(immobilisations)
        
        print("[{t:6.2f}] ►Calcul des déplacements individuels...".format(
            t =time.time()-p.starting_time))
        deplacements = autres.calc_deplacements(employ_par_site)
        resultat.enregistre_deplacements(deplacements,production_par_site)
        del(deplacements)
    #============================================================#
    #============== INTRANTS ET MATIERES PREMIERES ==============#
    #============================================================#
    if p.INTRANTS_ET_FRET:
        print("[{t:6.2f}] ►Lecture des intrants...".format(
            t =time.time()-p.starting_time))
        #Lecture du fichier d'achats
        listeAchats= gestionIntrants.lireBDDMP()
        
        
        print("[{t:6.2f}] ►Bilan carbone des intrants...".format(
            t =time.time()-p.starting_time))
        #Lecture des données fixes sur les matières premières (les FE sont en kgCO2/t, les masses volumiques en t/m3)
        FE_familles, MP_familles_N, masse_vol_MP,FE_engrais = gestionIntrants.lireFE_matprem()
        #On associe chaque matière première, engrais compris, à son facteur d'émission de fabrication (on ne touche pas aux unités)
        MP_et_FE, liste_engrais_eol = gestionIntrants.associerMPetFE_fab(
            MP_familles_N,FE_engrais,FE_familles)
        #Fait le lien entre la table fret et la table des FE des intrants
        MP_et_FE, liste_engrais_eol = gestionIntrants.associer_nom_FEfab(listeAchats, MP_et_FE,liste_engrais_eol)
        #Puis à une masse volumique
        masse_vol_MP = gestionIntrants.corriger_noms_massesvol(masse_vol_MP, MP_et_FE, listeAchats)
        
        print("[{t:6.2f}] ►Calcul du fret amont...".format(
            t =time.time()-p.starting_time))
        #Calcul du fret par matiere premiere
        #On calcule ensuite le bilan carbone du fret et de fabrication en même temps
        fret, resultat_fret_usine_MP= fonctionsFret.calc_fret_par_MP(
            listeAchats,masse_vol_MP, FE_route, FE_bateau, nom_fichier_resultat)
        print("[{t:6.2f}] ►Finalisation du BC fabrication et fret amont...".format(
            t =time.time()-p.starting_time))
        usines_et_fret = fonctionsFret.renverserMP_usine_fret(
            resultat_fret_usine_MP)
        BC_intrants, resultat_complet_BCMP_usine = gestionIntrants.BC_intrants_par_site(
            resultat_fret_usine_MP, MP_et_FE)
        
        fret = gestionIntrants.calc_fret_final(fret,MP_et_FE)
        #Tri à bulle de la liste de fret par plus grandes émissions
        
        print("[{t:6.2f}] ►Bilan carbone de la tourbe et des engrais (EoL)...".format(
            t =time.time()-p.starting_time))
        res_EoL_engrais, liste_engrais_eol = gestionIntrants.calcul_protoxyde(
            resultat_fret_usine_MP, liste_engrais_eol)
        res_EoL_tourbe, tourbes_eol = gestionIntrants.calcul_co2_tourbe(resultat_fret_usine_MP, masse_vol_MP)

        fret = gestionIntrants.ajoutEoLtableauFret(fret, liste_engrais_eol, tourbes_eol)
        
        fret_FE_sorted =  fonctionsMatrices.bubbleSortColonne(fret[1:],-1, decroissant = False)
        fret_FE_sorted.insert(0, fret[0])
        
                
        print("[{t:6.2f}] ►Enregistrement des résultats sur les matières premières...".format(
            t =time.time()-p.starting_time))
        resultat.enregistre_fret_amont_par_usine(usines_et_fret,production_par_site)
        resultat.enregistre_eol(production_par_site,resultat_fret_usine_MP,
                                masse_vol_MP ,res_EoL_engrais,res_EoL_tourbe)
    
        resultat.enregistre_BC_intrants(BC_intrants, production_par_site)
        #Ecriture du fret dans le fichier résultat
        affichageResultats.sauveFret(
            fret_FE_sorted, p.CHEMIN_ECRITURE_RESULTATS, nom_fichier_resultat)
        #Suppression des variables inutiles pour libérer de la mémoire vive
        del(listeAchats)
        del(FE_familles); del(MP_familles_N);del(masse_vol_MP);del(FE_engrais)
        del(MP_et_FE);del(fret);del(fret_FE_sorted);del(usines_et_fret);
    #====================================================#
    #============== EMBALLAGES ET SACHERIE ==============#
    #====================================================#
    if p.EMBALLAGES_ET_SACHERIE:
        print("[{t:6.2f}] ►Lecture des consommations sacherie...".format(
            t =time.time()-p.starting_time))
        refs_sacherie= gestionSacherie.lireConsoSacherie()
        
        print("[{t:6.2f}] ►Calcul du bilan carbone de la sacherie...".format(
            t =time.time()-p.starting_time))
        conso_materiaux_par_site, fret_sacherie_par_usine = gestionSacherie.qte_materiaux_sacherie(
            refs_sacherie)
        FE_emballage = gestionSacherie.lireFE_emballage()
        BC_fret_sacherie, BC_MP_sacherie = gestionSacherie.BC_sacherie(
            fret_sacherie_par_usine, conso_materiaux_par_site, FE_route, FE_emballage)
        
        print("[{t:6.2f}] ►Enregistrement des résultats sur la sacherie...".format(
            t =time.time()-p.starting_time))
        resultat.enregistre_conso_sacherie(BC_fret_sacherie, BC_MP_sacherie,production_par_site)
        del(conso_materiaux_par_site);del(refs_sacherie);del(BC_fret_sacherie);del(BC_MP_sacherie)
    
    #=======================================#
    #============== FRET AVAL ==============#
    #=======================================#
    if p.FRET_AVAL:
        bdd_exp_terre, bdd_exp_mari= gestionExport.lireExport()
        print("[{t:6.2f}] ►Lecture des densités...".format(t =time.time()-p.starting_time))
        densitesPF = gestionExport.lireDensitesPF()
        print("[{t:6.2f}] ►Lecture des regroupements...".format(t =time.time()-p.starting_time))
        groupements = gestionExport.liregroupements()
        print("[{t:6.2f}] ►Lecture des ventes...".format(t =time.time()-p.starting_time))
        listeVentes = gestionExport.lire_ventes(densitesPF, groupements)
        
        print("[{t:6.2f}] ►Calcul du bilan carbone du fret aval...".format(
            t =time.time()-p.starting_time))
        print("[{t:6.2f}] ►Groupement des livraisons...".format(
            t =time.time()-p.starting_time))
        livraisons = gestionExport.regrouper_livraison(listeVentes)
        print("[{t:6.2f}] ►Calcul des distances...".format(
            t =time.time()-p.starting_time))
        BC_fret_aval_par_usine, BC_fret_aval = gestionExport.calc_fret_aval(
            livraisons, FE_route, FE_bateau, bdd_exp_terre, bdd_exp_mari, nom_fichier_resultat)
        print("[{t:6.2f}] ►Enregistrement du fret aval...".format(
            t =time.time()-p.starting_time))
        resultat.enregistre_fret_aval_par_usine(BC_fret_aval_par_usine,production_par_site)
        del(bdd_exp_terre);del(bdd_exp_mari);
        del(listeVentes); del(livraisons); del(BC_fret_aval); del(BC_fret_aval_par_usine)
    
    
    
    print("[{t:6.2f}] ►Enregistrement des résultats généraux...".format(
        t =time.time()-p.starting_time))
    recap = resultat.regrouper_postes_resultat()
    affichageResultats.sauveTout(nom_fichier_resultat, recap)
    affichageResultats.ecrire_messages_derreur(nom_fichier_resultat)
    
    print("[{t:6.2f}] ►Programme terminé !".format(
        t =time.time()-p.starting_time))
    somme, somme_m3 = resultat.resultat_somme_simple()
    
    print("╦═╦═════════════════════════════════════════════════════════════")
    print("║ ╠═Résultat :")
    print("║ ╠═Bilan carbone "+str(p.ANNEE)+" : {:10.1f} tCO2".format(somme))
    print("║ ╠═Ramené à la production : {:10.2f} kgCO2e/m3".format(somme_m3))
    print("║ ╠═Détail des résultats dans '"+p.CHEMIN_ECRITURE_RESULTATS+"'")
    print("╩═╩═════════════════════════════════════════════════════════════")
    print("Temps d'exécution :{t:6.2f}s".format(t =time.time()-p.starting_time))
    return None


try:
    main()
    matres = p.MATRICE_RESULTAT
    input('Presser ENTREE pour fermer')
except KeyboardInterrupt:
    print("Arrêt du programme, calcul incomplet")