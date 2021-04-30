# -*- coding: utf-8 -*-
"""
Created on Tue Nov 10 14:08:31 2020

@author: lou
"""

import os
from os import listdir
from os.path import isfile, join


path = "/Users/mar_altermark/Documents/Calculateur Carbone"
onlyfiles = [f for f in listdir(path) if isfile(join(path, f))]

sommeLignes = 0
sommeCharac = 0
sommeTaille = 0
for file in onlyfiles:
    
    if '.py' in file and ("meta.py" not in file or True):
        print(file)
        with open(file,encoding='utf-8') as my_file:
            sommeLignes += sum(1 for _ in my_file)
        with open(file,encoding='utf-8') as my_file:
            sommeCharac += len(my_file.read())
        sommeTaille += os.stat(file).st_size
print(str(sommeTaille/1024)+" ko")
print(str(sommeLignes)+" lignes")
print(str(sommeCharac)+" caractères")
input("Presser <entrée> pour fermer")
