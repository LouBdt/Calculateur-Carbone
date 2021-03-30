# -*- coding: utf-8 -*-
"""
Created on Tue Nov 10 11:12:09 2020

@author: lou
"""
import xlrd
import parametres as p
import fonctionsMatrices
import sys
import inspect

def lire_FE_elec():
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
         
    feuille = document.sheet_by_index(2)    
    nlignes = feuille.nrows
    type_mixe =feuille.cell_value(7 ,9) 
    ligne = 12
    while ligne<nlignes:
        if feuille.cell_value(ligne ,8) == type_mixe:
            FE = feuille.cell_value(ligne ,9)
            break
        ligne+=1
    return FE

def lire_conso_elec():
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
    
    feuille= document.sheet_by_index(0)
    conso_par_site = [[[x[0],x[2],x[1]],0] for x in p.CP_SITES_FLORENTAISE]
    nmax = feuille.nrows
    i = 3
    while i<nmax and feuille.cell_value(i ,0)!="Conso Electrique (kWh)":
        i+=1
    colonne = 2
    ligne = i
    while colonne<40:
        if feuille.cell_value(ligne ,colonne)==p.ANNEE:
            break
        colonne+=1
    ligne +=1
    while ligne<=i + len(p.CP_SITES_FLORENTAISE):
        conso = feuille.cell_value(ligne,colonne)
        for x in conso_par_site:
            if x[0][0]==feuille.cell_value(ligne,0):
                try:
                    x[1]=float(conso)
                except ValueError:
                    fonctionsMatrices.print_log_erreur("Consommation électrique erronée", inspect.stack()[0][3])
                    x[1]=0
                    
                break
        ligne +=1
    total = 0
    for x in conso_par_site:
        total += x[1]
    return conso_par_site,total

def calc_BC_elec(FE_elec:float,conso:list):
    for x in conso:
        if x[0][2].lower()=="lavilledieu" and p.ANNEE >2017:
            x.append(p.FE_photovoltaique*x[1]/1000)
        else:
            x.append(FE_elec*x[1]/1000) #Conversion en tonnes
    return conso

def lire_FE_fuel():
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
    
    feuille = document.sheet_by_index(2)
    nlignes = feuille.nrows
    type_mixe =feuille.cell_value(8 ,9) 
    ligne = 12
    while ligne<nlignes:
        if feuille.cell_value(ligne ,8) == type_mixe:
            FE = feuille.cell_value(ligne ,9)
            break
        ligne+=1
    return FE

def lire_conso_fuel():
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
    
    feuille= document.sheet_by_index(0)
    conso_par_site = [[[x[0],x[2],x[1]],0] for x in p.CP_SITES_FLORENTAISE]
    nmax = feuille.nrows
    i = 3
    while i<nmax and feuille.cell_value(i ,0)!="Conso Fuel (m3)":
        i+=1
    colonne = 2
    ligne = i
    while colonne<40:
        if feuille.cell_value(ligne ,colonne)==p.ANNEE:
            break
        colonne+=1
    ligne +=1
    while ligne<=i+len(p.CP_SITES_FLORENTAISE):
        conso = feuille.cell_value(ligne,colonne)
        for x in conso_par_site:
            if x[0][0]==feuille.cell_value(ligne,0):
                try:
                    x[1]=float(conso)
                except ValueError:
                    fonctionsMatrices.print_log_erreur("Consommation de fuel erronée", inspect.stack()[0][3])
                    x[1]=0
                break
        ligne +=1
    total = 0
    for x in conso_par_site:
        total += x[1]
    return conso_par_site,total

def calc_BC_fuel(FE_fuel:float,conso:list):
    for x in conso:
        x.append(FE_fuel*x[1]/1000)
    return conso

def lire_donnees_entreprise():
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)
    
    feuille= document.sheet_by_index(0)
    employ_par_site = [[[x[0],x[2],x[1]],0] for x in p.CP_SITES_FLORENTAISE]
    production_par_site = [[[x[0],x[2],x[1]],0] for x in p.CP_SITES_FLORENTAISE]
    ligne = 7
    colonne = 2
    while colonne<22:
        if p.ANNEE==feuille.cell_value(ligne,colonne):
            break
        colonne+=1
    ligne +=1
    CA =  feuille.cell_value(ligne,colonne)
    ligne = 12
    while ligne<23:
        conso = feuille.cell_value(ligne,colonne)
        for x in employ_par_site:
            if x[0][0]==feuille.cell_value(ligne,0):
                x[1]=conso
                break
        ligne +=1
        
    ligne = 24
    while ligne<36:
        conso = feuille.cell_value(ligne,colonne)
        for x in production_par_site:
            if x[0][0]==feuille.cell_value(ligne,0):
                x[1]=conso
                break
        ligne +=1
    
    return CA, employ_par_site, production_par_site
    
def ajouter_immobilisations():
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)   
    feuille= document.sheet_by_index(6)
    maximum = feuille.nrows
    
    immobilisations= []
    ligne = 7
    while ligne<maximum:
        if feuille.cell_value(ligne,9) =="Total amorti (tCO2e)":
            usine = feuille.cell_value(ligne-5,0)
            ligne +=1
            amorti = feuille.cell_value(ligne,9)
            immobilisations.append([usine, amorti])
        ligne += 1
    return immobilisations

def calc_deplacements(nb_salariees):
    try:
         document = xlrd.open_workbook(p.FICHIER_ENTREES_MANUELLES)
    except FileNotFoundError:
         fonctionsMatrices.print_log_erreur("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES, inspect.stack()[0][3])
         sys.exit("Le fichier des entrées manuelles n'est pas trouvé à l'emplacement "+p.FICHIER_ENTREES_MANUELLES)   
    feuille= document.sheet_by_index(2)
    ligne = 12
    maximum = feuille.nrows
    #Valeurs par défaut
    dpt_perso = 5.53
    dpt_pro = 43.66
    while ligne<maximum:
        FE = feuille.cell_value(ligne,13)
        if FE=="Deplacements quotidiens":
            dpt_perso = feuille.cell_value(ligne,14)
        elif FE == "Deplacements pro":
            dpt_pro = feuille.cell_value(ligne,14)
        ligne += 1
    
    deplacements= []
    for k in nb_salariees:
        if k[0][2] == "Support":
            emissions = k[1]*dpt_pro
        else:
            emissions = k[1]*dpt_perso
        deplacements.append([k[0][2], emissions])
    return deplacements
    