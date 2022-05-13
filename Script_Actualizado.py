from sqlalchemy import create_engine
import numpy as np
import pandas as pd
import tkinter 
from tkinter import messagebox
import pymysql
import xlrd
#Desempaquetador 
from zipfile import ZipFile

def desempaquetar():
    documentZip = "C:\\Users\\USER\\Desktop\\Examen\\Examen.rar"
    with ZipFile(file = path, mode = 'r', allowzip64=True) as file:
        document = file.open(name=file.list()[0], mode ='r')
        print(document.read())
        document.close()
        navegation = ("C:\\Users\\USER\\Desktop\\Examen")
        file.extractall(path = navegation)
        print("Archivo desempaquetado")
desempaquetar()
#Creación de la base de datos y tabla
def firstconnection ():
        try:
                #aqui se habre la conexion utilizando una base de datos predeterminada
                connection=pymysql.Connection(host="localhost",user="root",password="",db="mysql")
                thecursor=connection.cursor()
                messagebox.showinfo("mensaje","La Conexion con el servidor ha sido exitosa ")
                #aqui se crea la base de datos utilizando la conexion existente
                thecursor.execute("CREATE DATABASE IF NOT EXISTS Exam")
                #cierre de conexion 
                connection.close()
                thecursor.close()
                #aqui se habre la conexion utilizando una base de datos creada anteriormente
                connection=pymysql.Connection(host="localhost",user="root",password="",db="Exam")
                thecursor=connection.cursor()
                #aqui se crea la tabla tblestudiantes
                thecursor.execute("CREATE TABLE IF NOT EXISTS JobTraveler (id_Serial_Number VARCHAR(255) NOT NULL, Panel_Number VARCHAR(255), Seal TINYINT(255), Job_Number VARCHAR(255), Job_Name VARCHAR(255), Type VARCHAR(255), Modbus_ID INT(255), Meter_No INT(255),  CONSTRAINT SerialNumber_pk PRIMARY KEY (id_Serial_Number) ) ")
                messagebox.showinfo("mensaje","La tabla Job_TRAVELER se ha creado correctamente")
                connection.commit()
                connection.close()
        except:
                messagebox.showinfo("mensaje","Por favor revise si la base de datos Exam\n y la Tabla JOB_TRAVELER ya Existe ")
command = firstconnection()

path = "C:\\Users\\USER\\Desktop\\Examen\\-32DPEA.xls" 

#Lectura de xls y adición a la base de datos 
connection= pymysql.Connection(host="localhost",user="root",password="",db="Exam") 
thecursor=connection.cursor()

l = list()

a=xlrd.open_workbook(path)
sheet = a.sheet_by_index(0)
sheet.cell_value(0,0)

path = "C:\\Users\\USER\\Desktop\\Examen\\-32DPEA.xls"
input_cols1 = [3]
input_cols2 = [3]
input_cols3 = [3]
input_cols4 = [1]
input_cols5 = [2]
input_cols6 = [2]
input_cols7 = [1]
input_cols8 = [9]

df1 = pd.read_excel(path, sheet_name="Sheet1", header = 2, usecols = input_cols1)
df2 = pd.read_excel(path, sheet_name="Sheet1", header = 3, usecols = input_cols2)
df3 = pd.read_excel(path, sheet_name="Sheet1", header = 4, usecols = input_cols3)
df4 = pd.read_excel(path, sheet_name="Sheet1", header = 27, usecols = input_cols4)
df5 = pd.read_excel(path, sheet_name="Sheet1", header = 32, usecols = input_cols5) 
df6 = pd.read_excel(path, sheet_name="Sheet1", header = 49, usecols = input_cols6) 
df7 = pd.read_excel(path, sheet_name="Sheet1", header = 49, usecols = input_cols7)
df8 = pd.read_excel(path, sheet_name="Sheet1", header = 2, usecols = input_cols8)

lista = []
lista.append(df6.columns)#Serial_Number
lista.append(df1.columns)#Panel_Number
lista.append(df8.columns)#seal
lista.append(df2.columns)#Job_Number
lista.append(df3.columns)#Job_Name
lista.append(df4.columns)# Type
lista.append(df5.columns)# Modbus_ID
lista.append(df7.columns)#Meter_No

print(len(lista))
arreglo = np.array(lista)
lista_arreglo=arreglo.tolist()
print(lista_arreglo)
sql = "INSERT INTO JobTraveler(id_Serial_Number, Panel_Number, Seal, Job_Number, Job_Name, Type, Modbus_ID, Meter_No) VALUES %s,%s,%s,%s,%s,%s,%s,%s"
thecursor.execute(sql,lista_arreglo)
connection.commit()
connection.close()



