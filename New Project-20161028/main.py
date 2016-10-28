#!/usr/bin/python
# -*- coding: utf-8 -*-


#Softtek Costa Rica
#Automatizar hojas de excel para trabajo
#Francis Alexander Lopez Montero
#28/10/2016

import struct
import string
#import openpyx1
import os
import shutil

def load_exc(a):
    
    direc=os.getcwd();
    #excpath=os.path.join(direc,"example.xlsx");
    excpath=os.path.join(direc,"Newfile.txt");
    #excpath=os.path.join(direc,"example.xlsx");
    exccopypath=os.path.join(direc,"NewfileW.txt");
    
    if (os.path.isfile(excpath)== True):
        print "Archivo excel existente \n";
        if (os.path.isfile(exccopypath)== True):
            print "Copia de excel para trabajo existente";
        else:
            print "No hay copia de excel para trabajo existente \n";
            print "Creando copia \n";
            shutil.copy2(excpath, exccopypath);
            
        #wb=openpyx1.load_workbook("exampleW.xls");
        wb=open(exccopypath);
    else:
        print "No hay archivo excel para leer";
        a=False;
        
    return wb;    
            
            
        
    



def main():
    oper=0
    typ=0#Formato simple=0;Formato doble=1
    n=1024;#Cantidad de datos
    a=True
    wb=load_exc(a)
    
    menu = {}
    menu['1']="Cambiar Operacion a realizar" 
    menu['2']="Formato simple/Formato doble"
    menu['3']="Crear valores aleatorios en hexadecimal"
    menu['4']="(Postsimulacion) Calcular y graficar error"
    menu['5']="Exit"
    while a:
        
	    options=menu.keys()
	    options.sort()
	    print "Operacion %d" %oper;
	    print "Formato %d" %typ;
	    print "Cantidad %d" %n;
	    for entry in options: 
			print entry, menu[entry]
	    selection=raw_input("Please Select:")
	    if selection =='1':
		    print
	    elif selection == '5': 
			break
	    print
	    wb.close();
	########Simulacion en Verilog(vivado o otra herramienta)
main();