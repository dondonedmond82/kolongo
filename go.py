import panel as pn
import pandas as pd
import hvplot.pandas
import warnings
import mysql.connector
import json
import datetime as dt
import os
# import create_user as nomusername
import connexion as connect
import login as log
import ctypes

warnings.filterwarnings("ignore")
pn.extension('tabulator')
pn.extension('echarts')


host_socasp, user_socasp, password_socasp, database_socasp, database_user = log.returnLogDetails()

#df_importation_between['Annee'] = df_importation_between['Annee'].astype("str")

def connexion(host_socasp, user_socasp, password_socasp):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp)
        mycursor = mydb_socasp.cursor()
        
        connect.createDatabaseSocasp(host_socasp, user_socasp, password_socasp)
        connect.createTableImportation(host_socasp, user_socasp, password_socasp, database_socasp)
        connect.createTableDistribution(host_socasp, user_socasp, password_socasp, database_socasp)
        connect.connectionUsername(host_socasp, user_socasp, password_socasp, database_socasp)
        connect.createTableUser(host_socasp, user_socasp, password_socasp, database_socasp)

        connect.createTableOrigine(host_socasp, user_socasp, password_socasp, database_socasp)
        connect.createTableProvenance(host_socasp, user_socasp, password_socasp, database_socasp)
        connect.createTableMarketeur(host_socasp, user_socasp, password_socasp, database_socasp)

        connect.connectionImportation(host_socasp, user_socasp, password_socasp, database_socasp)
        connect.connectionDistribution(host_socasp, user_socasp, password_socasp, database_socasp)

        connect.conUsername(host_socasp, user_socasp, password_socasp, database_socasp)

        # connect.connectionUser(host_socasp, user_socasp, password_socasp, database_socasp)

        connect.connectionOrigine(host_socasp, user_socasp, password_socasp, database_socasp)
        connect.connectionProvenance(host_socasp, user_socasp, password_socasp, database_socasp)
        connect.connectionMarketeur(host_socasp, user_socasp, password_socasp, database_socasp)

    except Exception as Ex:
        print("%s"%(Ex), "Alert!", 0)