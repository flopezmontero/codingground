#!/usr/bin/python
# -*- coding: utf-8 -*-


#Softtek Costa Rica
#Automatizar hojas de excel para trabajo
#Francis Alexander Lopez Montero
#28/10/2016


#Lista de Trabajo
#1. Carga de los archivos de Excel
#2. Almacenamiento de la configuración (Posicion de las celdas correspondientes a mi trabajo)
#3. localizar rango de celdas de trabajo
#4.  Catalogar errores (N/A, Passed, Failed, Failed Automatic)
#5. Escritura a hoja de excel

import struct
import string
import os
import shutil
import shelve
#import pyperclip http://coffeeghost.net/2010/10/09/pyperclip-a-cross-platform-clipboard-module-for-python/
#import openpyx1 #pip install openpyxl

def load_exc(a):#Carga de las hojas de excel
    
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
            
            
        
def config (oper,typ):
    oper=oper+typ+3;
    typ=oper*2;
    print "suma"
    return oper,typ;
    

def save (lista_pos,lista_fila,lista_col):
    list_file = shelve.open("lista")
    list_file['lista_pos'] = lista_pos
    list_file['lista_fila'] =lista_fila
    list_file['lista_col'] = lista_col
    list_file.close()
    
def load():
    lista_pos=[];
    lista_fila=[];
    lista_col=[];
    if (os.path.isfile("lista")== True):
        list_file = shelve.open("lista")
        lista_pos = list_file['lista_pos']
        lista_fila = list_file['lista_fila']
        lista_col = list_file['lista_col']
        list_file.close()
    return lista_pos,lista_fila,lista_col
#########################################################################################################################
#En prueba
########################################################################################################################

def find_and_list(wb,lista):
    name=raw_input("Digite nombre: ")
    n=0;
    sheet = wb.get_sheet_by_name('Hoja1')
    for i in range(2, sheet.max_row+1):
        if sheet.cell(row=i, column=1).value == name :
            lista.append(sheet.cell(row=i, column=1).coordinate);
            n=n+1;
            
    if lista == []:
        print ("No se encontraron similitudes")
    else:
        print ("Se encontraron y cargaron %d similitudes",n)
        
    return lista
    
    
def show_test (wb,lista_pos,lista_fila): #muestra el primer elemento de la lista con sus detalles
    sheet = wb.get_sheet_by_name('Hoja1')
    if (lista_pos==[]):
        print("No hay elementos cargados o ya ha finalizado su trabajo ;)")
    else:
        for i in range(1,sheet.max_column):
            print("%s: %s",sheet.cell(row=1, column=i).value, sheet.cell(row=lista_fila[0], column=i).value)
            if(sheet.cell(row(row=1, column=i).value=="Autocase")):
                pyperclip.copy(sheet.cell(row=lista_fila[0], column=i).value)
        print("###########################")
    return 0
    
#def modify_testcase(wb,lista_pos,lista_fila): #añadir la informacion extra


def main():
	
    lista_pos=[];
    lista_fila=[];
    lista_col=[];
    
    oper=0
    typ=0
    n=1024
    a=True
    wb=load_exc(a)
    lista_pos,lista_fila,lista_col = load ();
    
    menu = {}
    menu['1']="Almacenar rango de trabajo" 
    menu['2']="Obtener info de un testcase"
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
	        oper,typ=config(oper,typ);
	        print
	    elif selection == '2':
	        #lista.append(oper)
	        #pyperclip.copy("ola k ase")
	        print
	    elif selection == '3':
	        for i in lista:
	            print(i)
	        print
	    elif selection == '5':
	        save(lista_pos,lista_fila,lista_col)
	        break
	    print
	    wb.close();
main();
