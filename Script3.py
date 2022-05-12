from sqlalchemy import create_engine
import pandas as pd
import tkinter 
from tkinter import messagebox
import pymysql

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

sql = "INSERT INTO JobTraveler(id_Serial_Number, Panel_Number, Seal, Job_Number, Job_Name, Type, Modbus_ID, Meter_No) VALUES()"


db="mysql"
mydb = pymysql.Connection(host="localhost",user="root",password="",db="Exam") 
table = "JobTraveler"
path = "C:\\Users\\USER\\Desktop\\Examen\\-32DPEA.xls"

url = "myslq+mysqlconnector://root: /"

engine = create_engine(url + mydb, echo = False )






input_cols = [0,3,6,9]
input_cols1 = [0,3]
input_cols2 = [0,3]
input_cols3 = [0,1]
input_cols4 = [0,2]
