# -*- coding: utf-8 -*-
"""
Created on Thu Nov  5 09:44:22 2020

@author: lou
"""
import parametres as p

def resultat_somme_simple():
    somme = 0
    for ligne in p.MATRICE_RESULTAT[1:]:
        somme += ligne[2][0]
    somme_m3 =1000*somme/max(sum([x[1]for x in p.PROD_PAR_SITE]),1)
    return somme, somme_m3

def regrouper_postes_resultat():
    total = [["", "Fret amont", "Fret aval", "Sacherie", "Mat. Prem.", "Energie","Utilisation/EoL", "Autre", "Total"]]
    total.append(["Emissions totales (tCO2e)"]+[0 for x in total[0][1:]])    #BC total
    
    for poste in p.MATRICE_RESULTAT[1:]:
        if poste[0] in ["MP bateau", "MP route"]:
            total[1][1]+=poste[2][0]
            
        elif poste[0] in ["Ventes pro", "Ventes Jardineries", "Interdepot"]:
            total[1][2]+=poste[2][0]
            
        elif poste[0] in ["Matières Premières"]:
            total[1][4]+=poste[2][0]
            
        elif poste[0] in ["CO2 issu de la tourbe","Protoxyde d'azote"]:
            total[1][6]+=poste[2][0]
            
        elif poste[0] in ["Electricite","Fuel", "Electricité"]:
            total[1][5]+=poste[2][0]
            
        elif poste[0] in ["PE neuf","PE recyclé", "PVC", "Carton",
                          "Autres Emballages", "Fret amont des emballages"]:
            total[1][3]+=poste[2][0]
            
        else:
            total[1][-2]+=poste[2][0]
            
    total[1][-1] = sum(total[1][1:-1])
    
    return total

def creerMatriceResultat():
    usines = []
    for x in p.CP_SITES_FLORENTAISE:
        if x[3]<=p.ANNEE:
            usines.append([x[0],x[1]])
    
    FretAmont = ["MP bateau", "MP route"]
    FretAval = ["Ventes pro","Ventes Jardineries", "Interdepot"]
    Intrants = ["Matières Premières", "CO2 issu de la tourbe"]
    HorsEnergie = ["Protoxyde d'azote"]
    Energie = ["Electricité", "Fuel"]
    Emballages = ["PE neuf", "PE recyclé", "Terre cuite", "Fret amont des emballages"]
    Autres = ["Deplacements", "Immobilisations"]
    entetes = FretAmont+FretAval+Intrants+HorsEnergie+Energie+Emballages+Autres
    colonnes = ["Poste d'émission", "Site", "Empreinte carbone","Unité", "Emission par m3", "Unité"]
    matriceRes = [colonnes]
    for poste in entetes:
        for j in range(len(usines)):
            matriceRes.append([poste, usines[j], [0, "teCO2",0, "kgCO2e/m3"] ])
    return matriceRes


def enregistre_BC_intrants(BC_intrants:list, production_par_site:list):
    i = 0
    while i<len(p.MATRICE_RESULTAT):
        usine = p.MATRICE_RESULTAT[i][1][1] #Nom de l'usine
        poste = p.MATRICE_RESULTAT[i][0]
        if poste=="Matières Premières":
            for site in BC_intrants:
                if site[0][2]==usine:
                    p.MATRICE_RESULTAT[i][2][0] = site[1]
                    break
            for u in production_par_site:
                if u[0][2]==usine and u[1]!=0:
                    p.MATRICE_RESULTAT[i][2][2] =1000*p.MATRICE_RESULTAT[i][2][0]/u[1]
                    
                elif u[0][2]==usine:
                    p.MATRICE_RESULTAT[i][2][2] =0
        i+=1

def enregistre_conso_sacherie(BC_fret_sacherie:list, BC_MP_sacherie:list,production_par_site:list):
    i = 0
    
    while i<len(p.MATRICE_RESULTAT):
        usine = p.MATRICE_RESULTAT[i][1][1] #Nom de l'usine
        poste = p.MATRICE_RESULTAT[i][0]
        if poste=="Fret amont des emballages":
            pos = 1
            for k in range(len(BC_fret_sacherie[1:])):
                 if BC_fret_sacherie[k+1][0][2]==usine:
                     pos = k+1
                     break
            p.MATRICE_RESULTAT[i][2][0] = BC_fret_sacherie[pos][1]
            for u in production_par_site:
                if u[0][2]==usine and u[1]!=0:
                    p.MATRICE_RESULTAT[i][2][2] =1000*p.MATRICE_RESULTAT[i][2][0]/u[1]
                elif u[0][2]==usine:
                    p.MATRICE_RESULTAT[i][2][2] =0     
        elif poste in ["PE neuf", "PE recyclé", "Papier", "PVC", "Carton", "Autres Emballages"]:
            pos =1
            for k in range(len(BC_MP_sacherie[1:])):
                 if BC_MP_sacherie[k+1][0][2]==usine:
                     pos = k+1
                     break
            if poste=="PE neuf":
                p.MATRICE_RESULTAT[i][2][0] = BC_MP_sacherie[pos][1][0]
            elif poste=="PE recyclé":
                p.MATRICE_RESULTAT[i][2][0] = BC_MP_sacherie[pos][1][1]
            elif poste=="Papier":
                p.MATRICE_RESULTAT[i][2][0] = BC_MP_sacherie[pos][1][2]
            elif poste=="PVC":
                p.MATRICE_RESULTAT[i][2][0] = BC_MP_sacherie[pos][1][3]
            elif poste=="Carton":
                p.MATRICE_RESULTAT[i][2][0] = BC_MP_sacherie[pos][1][4]
            elif poste=="Autres Emballages":
                p.MATRICE_RESULTAT[i][2][0] = BC_MP_sacherie[pos][1][5]
            
        
            for u in production_par_site:
                if u[0][2]==usine and u[1]!=0:
                    p.MATRICE_RESULTAT[i][2][2] =1000*p.MATRICE_RESULTAT[i][2][0]/u[1]
                elif u[0][2]==usine:
                    p.MATRICE_RESULTAT[i][2][2] =0     
        i+=1
    return None

def enregistre_fret_amont_par_usine(usines_fret:list,prod_par_site:list):
    i = 0
    while i<len(p.MATRICE_RESULTAT):
        poste = p.MATRICE_RESULTAT[i][0]
        site = p.MATRICE_RESULTAT[i][1][1] #Nom de l'usine
        
        if poste=="MP bateau":
            for usine in usines_fret:
                if usine[0].lower()==site.lower():
                    
                    p.MATRICE_RESULTAT[i][2][0]=usine[5]
                    p.MATRICE_RESULTAT[i][2][1]='teCO2'
                    for u in prod_par_site:
                        if u[0][2]==site and u[1]!=0:
                            p.MATRICE_RESULTAT[i][2][2] = (1000*usine[5])/(u[1])
                            break     
                    p.MATRICE_RESULTAT[i][2][3] = "kgeCO2/m3"
                    break
        elif poste=="MP route":
            for usine in usines_fret:
                if usine[0].lower()==site.lower():
                    p.MATRICE_RESULTAT[i][2][0]=usine[4]
                    p.MATRICE_RESULTAT[i][2][1]='teCO2'
                    for u in prod_par_site:
                        if u[0][2]==site and u[1]!=0:
                            p.MATRICE_RESULTAT[i][2][2] = (1000*usine[4])/(u[1])
                        elif u[0][2]==usine:
                            p.MATRICE_RESULTAT[i][2][2] =0
                    p.MATRICE_RESULTAT[i][2][3] = "kgeCO2/m3"
                    break
        i+=1

def enregistre_BC_elec(BC_par_site:list, prod_par_site:list):
     i = 0
     while i<len(p.MATRICE_RESULTAT):
        poste = p.MATRICE_RESULTAT[i][0]
        if poste=="Electricité":
            for usine in BC_par_site:
                if usine[0][0]==p.MATRICE_RESULTAT[i][1][0]:
                    p.MATRICE_RESULTAT[i][2][0]=usine[2]
                    p.MATRICE_RESULTAT[i][2][1]='teCO2'
                    for u in prod_par_site:
                        if u[0][0]==usine[0][0]and u[1]!=0:
                            p.MATRICE_RESULTAT[i][2][2] = (1000*usine[2])/(u[1])
                        elif u[0][2]==usine:
                            p.MATRICE_RESULTAT[i][2][2] =0                   
                    p.MATRICE_RESULTAT[i][2][3] = "kgeCO2/m3"
                    break
        i+=1
        
def enregistre_BC_fuel(BC_par_site:list, prod_par_site:list):
     i = 0
     while i<len(p.MATRICE_RESULTAT):
        poste = p.MATRICE_RESULTAT[i][0]
        if poste=="Fuel":
            for usine in BC_par_site:
                if usine[0][0]==p.MATRICE_RESULTAT[i][1][0]:
                    p.MATRICE_RESULTAT[i][2][0]=usine[2]
                    p.MATRICE_RESULTAT[i][2][1]='teCO2'
                    for u in prod_par_site:
                        if u[0][0]==usine[0][0]and u[1]!=0:
                            p.MATRICE_RESULTAT[i][2][2] = (1000*usine[2])/(u[1])
                        elif u[0][2]==usine:
                            p.MATRICE_RESULTAT[i][2][2] =0  
                    p.MATRICE_RESULTAT[i][2][3] = "kgeCO2/m3"
                    break
        i+=1
 
def enregistre_eol(prod_par_site:list,MP_par_usine:list,massesvols:list, res_EoL_engrais:list,res_EoL_tourbe:list):
    i = 0
    while i<len(p.MATRICE_RESULTAT):
        poste = p.MATRICE_RESULTAT[i][0]
        if poste=="CO2 issu de la tourbe" and p.MATRICE_RESULTAT[i][1][0]!=-1:
            for usine in res_EoL_tourbe:
                if usine[0][1]==p.MATRICE_RESULTAT[i][1][1]:
                    p.MATRICE_RESULTAT[i][2][0]=usine[1]
                    p.MATRICE_RESULTAT[i][2][1]='teCO2'
                    for u in prod_par_site:
                        if u[0][2]==usine[0][1]and u[1]!=0:
                            p.MATRICE_RESULTAT[i][2][2] = (1000*usine[1])/u[1]
                        elif u[0][2]==usine:
                            p.MATRICE_RESULTAT[i][2][2] =0  
                    p.MATRICE_RESULTAT[i][2][3] = "kgeCO2/m3"
                    break        
        elif poste=="Protoxyde d'azote" and p.MATRICE_RESULTAT[i][1][0]!=-1:
            for usine in res_EoL_engrais:
                if usine[0][1]==p.MATRICE_RESULTAT[i][1][1]:
                    p.MATRICE_RESULTAT[i][2][0]=usine[1]
                    p.MATRICE_RESULTAT[i][2][1]='teCO2'
                    for u in prod_par_site:
                        if u[0][2]==usine[0][1]and u[1]!=0:
                            p.MATRICE_RESULTAT[i][2][2] = (1000*usine[1])/u[1]
                        elif u[0][2]==usine:
                            p.MATRICE_RESULTAT[i][2][2] =0
                    p.MATRICE_RESULTAT[i][2][3] = "kgeCO2/m3"
                    break
        i+=1
        

def enregistre_fret_aval_par_usine(usines_fret_aval:list,prod_par_site:list):
    i = 0
    while i<len(p.MATRICE_RESULTAT):
        poste = p.MATRICE_RESULTAT[i][0]
        site = p.MATRICE_RESULTAT[i][1][1] #Nom de l'usine
        if poste=="Ventes Jardineries":
            for usine in usines_fret_aval:
                if usine[0][2]==site:
                    p.MATRICE_RESULTAT[i][2][0]=usine[1][0]
                    p.MATRICE_RESULTAT[i][2][1]='teCO2'
                    for u in prod_par_site:
                        if u[0][2]==site and u[1]!=0:
                            p.MATRICE_RESULTAT[i][2][2] = (1000*usine[1][0])/u[1]
                        elif u[0][2]==usine:
                            p.MATRICE_RESULTAT[i][2][2] =0     
                    p.MATRICE_RESULTAT[i][2][3] = "kgeCO2/m3"
                    break
        elif poste=="Ventes pro":
            for usine in usines_fret_aval:
                if usine[0][2]==site:
                    p.MATRICE_RESULTAT[i][2][0]=usine[1][2]
                    p.MATRICE_RESULTAT[i][2][1]='teCO2'
                    for u in prod_par_site:
                        if u[0][2]==site and u[1]!=0:
                            p.MATRICE_RESULTAT[i][2][2] = (1000*usine[1][2])/u[1]
                        elif u[0][2]==usine:
                            p.MATRICE_RESULTAT[i][2][2] =0   
                    p.MATRICE_RESULTAT[i][2][3] = "kgeCO2/m3"
                    break
        elif poste=="Interdepot":
            for usine in usines_fret_aval:
                if usine[0][2]==site:
                    p.MATRICE_RESULTAT[i][2][0]=usine[1][1]
                    p.MATRICE_RESULTAT[i][2][1]='teCO2'
                    for u in prod_par_site:
                        if u[0][2]==site and u[1]!=0:
                            p.MATRICE_RESULTAT[i][2][2] = (1000*usine[1][1])/u[1]
                        elif u[0][2]==usine:
                            p.MATRICE_RESULTAT[i][2][2] =0   
                    p.MATRICE_RESULTAT[i][2][3] = "kgeCO2/m3"
                    break
        i+=1
    return None

def enregistre_immobilisations(immos:list,prod_par_site:list ):
    i = 0
    while i<len(p.MATRICE_RESULTAT):
        poste = p.MATRICE_RESULTAT[i][0]
        if poste=="Immobilisations":
            site = p.MATRICE_RESULTAT[i][1][1] #Nom de l'usine
            for m in immos:
                if m[0].lower()==site.lower():
                    p.MATRICE_RESULTAT[i][2][0]=m[1]
                    p.MATRICE_RESULTAT[i][2][1]='teCO2'
                    for u in prod_par_site:
                        if u[0][2]==site and u[1]!=0:
                            p.MATRICE_RESULTAT[i][2][2] = (1000*m[1])/u[1]
                        elif u[0][2]==site:
                            p.MATRICE_RESULTAT[i][2][2] =0
                    p.MATRICE_RESULTAT[i][2][3] = "kgeCO2/m3"
                    break
        i += 1
    return None

def enregistre_deplacements(depl:list, prod_par_site:list ):
    i = 0
    while i<len(p.MATRICE_RESULTAT):
        poste = p.MATRICE_RESULTAT[i][0]
        if poste=="Deplacements":
            site = p.MATRICE_RESULTAT[i][1][1] #Nom de l'usine
            for m in depl:
                if m[0].lower()==site.lower():
                    p.MATRICE_RESULTAT[i][2][0]=m[1]
                    p.MATRICE_RESULTAT[i][2][1]='teCO2'
                    for u in prod_par_site:
                        if u[0][2]==site and u[1]!=0:
                            p.MATRICE_RESULTAT[i][2][2] = (1000*m[1])/u[1]
                        elif u[0][2]==site:
                            p.MATRICE_RESULTAT[i][2][2] =0
                    p.MATRICE_RESULTAT[i][2][3] = "kgeCO2/m3"
                    break
        i += 1
    return None