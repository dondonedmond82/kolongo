import panel as pn
import pandas as pd
import hvplot.pandas
import warnings
import mysql.connector
import json
import login as log
import ctypes

warnings.filterwarnings("ignore")

host_socasp, user_socasp, password_socasp, database_socasp, database_user = log.returnLogDetails()


def createDatabaseSocasp(host_socasp, user_socasp, password_socasp):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp)

        mycursor = mydb_socasp.cursor()

        mycursor.execute("create database if not exists socasp")

    except Exception as Ex:
        print("%s"%(Ex))


def createTableImportation(host_socasp, user_socasp, password_socasp, database_socasp):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        mycursor = mydb_socasp.cursor()

        mycursor.execute("create table if not exists importation(carbuId int not null auto_increment primary key, origine varchar(255) not null, provenance varchar(255) not null, "
        "marketeur varchar(255) not null, essence int not null, jet int not null, petrole int not null, gazoil int not null, anneemois date not null)")
    
    except Exception as Ex:
        print("%s"%(Ex))


def createTableDistribution(host_socasp, user_socasp, password_socasp, database_socasp):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        mycursor = mydb_socasp.cursor()
        
        mycursor.execute("create table if not exists distribution(carbuId int not null auto_increment primary key, marketeur varchar(255) not null,"
        "essence int not null, jet int not null, petrole int not null, gazoil int not null, anneemois date not null)")

    except Exception as Ex:
        print("%s"%(Ex))



def createTableUser(host_socasp, user_socasp, password_socasp, database_socasp):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        mycursor = mydb_socasp.cursor()

        mycursor.execute("create table if not exists username(user varchar(255) not null primary key, password varchar(255) not null)")
    
    except Exception as Ex:
        print("%s"%(Ex))



def connectionImportation(host_socasp, user_socasp, password_socasp, database_socasp):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        df_importation_excel = './data/importation.xlsx'

        mycursor_socasp = mydb_socasp.cursor()
        mycursor_socasp.execute("""select carbuId, origine, provenance, marketeur, essence, jet, petrole, gazoil, anneemois, year(anneemois), CASE MONTH(anneemois) 
            WHEN 1 THEN 'Janvier'
            WHEN 2 THEN 'Février'
            WHEN 3 THEN 'Mars'
            WHEN 4 THEN 'Avril'
            WHEN 5 THEN 'Mai'
            WHEN 6 THEN 'Juin'
            WHEN 7 THEN 'Juillet'
            WHEN 8 THEN 'Aôut'
            WHEN 9 THEN 'Septembre'
            WHEN 10 THEN 'Octobre'
            WHEN 11 THEN 'Novembre'
            WHEN 12 THEN 'Décembre'
            END as Mes, month(anneemois) from importation ORDER BY 1 ASC""")
        myresult_socasp = mycursor_socasp.fetchall()

        columns_socasp = ['Id', 'Origine', 'Provenance', 'Marketeur', 'Super', 'JetA1', 'Petrole', 'Gazoil', 'Submit', 'Annee', 'Mois', 'No']
        data_socasp = []

        for x_socasp in myresult_socasp:
            data_socasp.append(x_socasp)

        df_socasp = pd.DataFrame(data_socasp, columns=columns_socasp)
        df_socasp.to_excel(df_importation_excel)

    except Exception as Ex:
        print("%s"%(Ex))




def conUsername(host_socasp, user_socasp, password_socasp, database_socasp):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        df_importation_excel = './data/username.xlsx'

        mycursor_socasp = mydb_socasp.cursor()
        mycursor_socasp.execute("""select user, password from username""")
        myresult_socasp = mycursor_socasp.fetchall()

        columns_socasp = ['Utilisateur', 'Mots de Passe']
        data_socasp = []

        for x_socasp in myresult_socasp:
            data_socasp.append(x_socasp)

        df_socasp = pd.DataFrame(data_socasp, columns=columns_socasp)
        df_socasp.to_excel(df_importation_excel)

    except Exception as Ex:
        print("%s"%(Ex))


def connectionUser(host_socasp, user_socasp, password_socasp, database_socasp):

    try:

        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        mycursor_socasp = mydb_socasp.cursor()
        mycursor_socasp.execute("select user, password from username")

        rows = mycursor_socasp.fetchall()

        thisdict = dict(rows)


        with open("credentials.json", "w") as final:
            json.dump(thisdict, final)

    except Exception as Ex:
        print("%s"%(Ex))


def insertUser(host_socasp, user_socasp, password_socasp, database_socasp, user, password):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)
        mycursor = mydb_socasp.cursor()
        
        sql = "INSERT INTO username (user, password) VALUES (%s, %s)"
        val = (user, password)

        mycursor.execute(sql, val)
        mydb_socasp.commit()

    except Exception as Ex:
        print("%s"%(Ex))



def connection_importation_bis(host_socasp, user_socasp, password_socasp, database_socasp, value_marketeur):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        df_importation_excel = './data/importationparmarketeurbis.xlsx'

        mycursor_socasp = mydb_socasp.cursor()
        mycursor_socasp.execute(f"""select carbuId, origine, provenance, marketeur, essence, jet, petrole, gazoil, year(anneemois), CASE MONTH(anneemois) 
            WHEN 1 THEN 'Janvier'
            WHEN 2 THEN 'Février'
            WHEN 3 THEN 'Mars'
            WHEN 4 THEN 'Avril'
            WHEN 5 THEN 'Mai'
            WHEN 6 THEN 'Juin'
            WHEN 7 THEN 'Juillet'
            WHEN 8 THEN 'Aôut'
            WHEN 9 THEN 'Septembre'
            WHEN 10 THEN 'Octobre'
            WHEN 11 THEN 'Novembre'
            WHEN 12 THEN 'Décembre'
            END as Mes, month(anneemois) from importation where marketeur = {value_marketeur} ORDER BY 1 ASC""")
        myresult_socasp = mycursor_socasp.fetchall()

        columns_socasp = ['Id', 'Origine', 'Provenance', 'Marketeur', 'Super', 'JetA1', 'Petrole', 'Gazoil', 'Annee', 'Mois', 'No']
        data_socasp = []

        for x_socasp in myresult_socasp:
            data_socasp.append(x_socasp)

        df_socasp = pd.DataFrame(data_socasp, columns=columns_socasp)
        df_socasp.to_excel(df_importation_excel)


    except Exception as Ex:
        print("%s"%(Ex))




def connectionDistribution(host_socasp, user_socasp, password_socasp, database_socasp):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        df_distribution_excel = './data/distribution.xlsx'

        mycursor_socasp = mydb_socasp.cursor()
        mycursor_socasp.execute("""select carbuId, marketeur, essence, jet, petrole, gazoil, year(anneemois), CASE MONTH(anneemois) 
            WHEN 1 THEN 'Janvier'
            WHEN 2 THEN 'Février'
            WHEN 3 THEN 'Mars'
            WHEN 4 THEN 'Avril'
            WHEN 5 THEN 'Mai'
            WHEN 6 THEN 'Juin'
            WHEN 7 THEN 'Juillet'
            WHEN 8 THEN 'Aôut'
            WHEN 9 THEN 'Septembre'
            WHEN 10 THEN 'Octobre'
            WHEN 11 THEN 'Novembre'
            WHEN 12 THEN 'Décembre'
            END as Mes, month(anneemois), anneemois, 
            
            CASE weekday(anneemois) 
            WHEN 0 THEN 'Lundi'
            WHEN 1 THEN 'Mardi'
            WHEN 2 THEN 'Mercredi'
            WHEN 3 THEN 'Jeudi'
            WHEN 4 THEN 'Vendredi'
            WHEN 5 THEN 'Samedi'
            WHEN 6 THEN 'Dimanche'
            END as Da, day(anneemois), weekday(anneemois) from distribution ORDER BY 1 ASC""")
        myresult_socasp = mycursor_socasp.fetchall()

        columns_socasp = ['Id', 'Marketeur', 'Super', 'JetA1', 'Petrole', 'Gazoil', 'Annee', 'Mois', 'No', 'Date', 'Jour Nom', 'Jour', 'Semaine']
        data_socasp = []

        for x_socasp in myresult_socasp:
            data_socasp.append(x_socasp)

        df_socasp = pd.DataFrame(data_socasp, columns=columns_socasp)
        df_socasp.to_excel(df_distribution_excel)


    except Exception as Ex:
        print("%s"%(Ex))



def insertImportation(host_socasp, user_socasp, password_socasp, database_socasp, origine, provenance, marketeur, essence, jet, petrole, gazoil, anneemois):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)
        mycursor = mydb_socasp.cursor()
        
        sql = "INSERT INTO importation (origine, provenance, marketeur, essence, jet, petrole, gazoil, anneemois) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"
        val = (origine, provenance, marketeur, essence, jet, petrole, gazoil, anneemois)

        mycursor.execute(sql, val)
        mydb_socasp.commit()


    except Exception as Ex:
        print("%s"%(Ex))



def insertDistribution(host_socasp, user_socasp, password_socasp, database_socasp, marketeur, essence, jet, petrole, gazoil, anneemois):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        mycursor = mydb_socasp.cursor()

        sql = "INSERT INTO distribution (marketeur, essence, jet, petrole, gazoil, anneemois) VALUES (%s, %s, %s, %s, %s, %s)"
        val = (marketeur, essence, jet, petrole, gazoil, anneemois)

        mycursor.execute(sql, val)
        mydb_socasp.commit()

    except Exception as Ex:
        print("%s"%(Ex))



def createUserSocasp(host_socasp, user_socasp, password_socasp, database_socasp, usernamesocasp, passwordsocasp):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp)

        mycursor = mydb_socasp.cursor()

        sqlCreateUser = "CREATE USER '%s'@'localhost' IDENTIFIED BY '%s';"%(usernamesocasp, passwordsocasp)
        mycursor.execute(sqlCreateUser)

    except Exception as Ex:
        print("%s"%(Ex))



def connectionUsername(host_socasp, user_socasp, password_socasp, database_user):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_user)

        df_user_excel = './data/utilisateurmotsdepasse.xlsx'

        mycursor_socasp = mydb_socasp.cursor()
        mycursor_socasp.execute("""select host, user, Insert_priv, Update_priv, Delete_priv, Create_priv from mysql.user""")
        myresult_socasp = mycursor_socasp.fetchall()

        columns_socasp = ['Hote', 'Utilisateur', "Insertion", "Modifier", "Supprimer", "Creer"]
        data_socasp = []

        for x_socasp in myresult_socasp:
            data_socasp.append(x_socasp)

        df_socasp = pd.DataFrame(data_socasp, columns=columns_socasp)
        df_socasp.to_excel(df_user_excel)


    except Exception as Ex:
        print("%s"%(Ex))



def exportUser(host_socasp, user_socasp, password_socasp, database_user):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_user)

        mycursor_socasp = mydb_socasp.cursor()
        mycursor_socasp.execute("select host, user from user")

        rows = mycursor_socasp.fetchall()
        thisdict = dict(rows)

        with open("credentials.json", "w") as final:
            json.dump(thisdict, final)

    except Exception as Ex:
        print("%s"%(Ex))


def betweenDate(host_socasp, user_socasp, password_socasp, database_socasp, first_year, second_year):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        mycursor_socasp = mydb_socasp.cursor()
        mycursor_socasp.execute(f"""select carbuId, origine, provenance, marketeur, essence, jet, petrole, gazoil, anneemois, year(anneemois), CASE MONTH(anneemois) 
            WHEN 1 THEN 'Janvier'
            WHEN 2 THEN 'Février'
            WHEN 3 THEN 'Mars'
            WHEN 4 THEN 'Avril'
            WHEN 5 THEN 'Mai'
            WHEN 6 THEN 'Juin'
            WHEN 7 THEN 'Juillet'
            WHEN 8 THEN 'Aôut'
            WHEN 9 THEN 'Septembre'
            WHEN 10 THEN 'Octobre'
            WHEN 11 THEN 'Novembre'
            WHEN 12 THEN 'Décembre'
            END as Mes, month(anneemois) from importation where anneemois between {first_year} and {second_year} ORDER BY 1 ASC""")
        myresult_socasp = mycursor_socasp.fetchall()

        columns_socasp = ['Id', 'Origine', 'Provenance', 'Marketeur', 'Super', 'JetA1', 'Petrole', 'Gazoil', 'Date', 'Annee', 'Mois', 'No']
        data_socasp = []

        for x_socasp in myresult_socasp:
            data_socasp.append(x_socasp)


        df_socasp = pd.DataFrame(data_socasp, columns=columns_socasp)
        df_socasp.to_excel('./data/between_importation_annee.xlsx')


    except Exception as Ex:
        print("%s"%(Ex))



def betweenImportationAnnee(host_socasp, user_socasp, password_socasp, database_socasp, first_year, second_year):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        mycursor_socasp = mydb_socasp.cursor()
        mycursor_socasp.execute(f"""select carbuId, origine, provenance, marketeur, essence, jet, petrole, gazoil, anneemois, year(anneemois), CASE MONTH(anneemois) 
            WHEN 1 THEN 'Janvier'
            WHEN 2 THEN 'Février'
            WHEN 3 THEN 'Mars'
            WHEN 4 THEN 'Avril'
            WHEN 5 THEN 'Mai'
            WHEN 6 THEN 'Juin'
            WHEN 7 THEN 'Juillet'
            WHEN 8 THEN 'Aôut'
            WHEN 9 THEN 'Septembre'
            WHEN 10 THEN 'Octobre'
            WHEN 11 THEN 'Novembre'
            WHEN 12 THEN 'Décembre'
            END as Mes, month(anneemois) from importation where year(anneemois) between {first_year} and {second_year} ORDER BY 1 ASC""")
        myresult_socasp = mycursor_socasp.fetchall()

        columns_socasp = ['Id', 'Origine', 'Provenance', 'Marketeur', 'Super', 'JetA1', 'Petrole', 'Gazoil', 'Submit', 'Annee', 'Mois', 'No']
        data_socasp = []

        for x_socasp in myresult_socasp:
            data_socasp.append(x_socasp)

        df_socasp = pd.DataFrame(data_socasp, columns=columns_socasp)
        df_socasp.to_excel('./data/between_importation_annee.xlsx')


    except Exception as Ex:
        print("%s"%(Ex))



def betweenImportationAnneeDate(host_socasp, user_socasp, password_socasp, database_socasp, first_year):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        mycursor_socasp = mydb_socasp.cursor()
        mycursor_socasp.execute(f"""select carbuId, origine, provenance, marketeur, essence, jet, petrole, gazoil, anneemois, year(anneemois), CASE MONTH(anneemois) 
            WHEN 1 THEN 'Janvier'
            WHEN 2 THEN 'Février'
            WHEN 3 THEN 'Mars'
            WHEN 4 THEN 'Avril'
            WHEN 5 THEN 'Mai'
            WHEN 6 THEN 'Juin'
            WHEN 7 THEN 'Juillet'
            WHEN 8 THEN 'Aôut'
            WHEN 9 THEN 'Septembre'
            WHEN 10 THEN 'Octobre'
            WHEN 11 THEN 'Novembre'
            WHEN 12 THEN 'Décembre'
            END as Mes, month(anneemois) from importation where year(anneemois) = {first_year} ORDER BY 1 ASC""")
        myresult_socasp = mycursor_socasp.fetchall()

        columns_socasp = ['Id', 'Origine', 'Provenance', 'Marketeur', 'Super', 'JetA1', 'Petrole', 'Gazoil', 'Submit', 'Annee', 'Mois', 'No']
        data_socasp = []

        for x_socasp in myresult_socasp:
            data_socasp.append(x_socasp)

        df_socasp = pd.DataFrame(data_socasp, columns=columns_socasp)
        df_socasp.to_excel('./data/unique_importation_annee.xlsx')


    except Exception as Ex:
        print("%s"%(Ex))


def betweenDistributionMarketeur(host_socasp, user_socasp, password_socasp, database_socasp, first_year, second_year):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        mycursor_socasp = mydb_socasp.cursor()
        mycursor_socasp.execute(f"""select carbuId, marketeur, essence, jet, petrole, gazoil, year(anneemois), CASE MONTH(anneemois) 
            WHEN 1 THEN 'Janvier'
            WHEN 2 THEN 'Février'
            WHEN 3 THEN 'Mars'
            WHEN 4 THEN 'Avril'
            WHEN 5 THEN 'Mai'
            WHEN 6 THEN 'Juin'
            WHEN 7 THEN 'Juillet'
            WHEN 8 THEN 'Aôut'
            WHEN 9 THEN 'Septembre'
            WHEN 10 THEN 'Octobre'
            WHEN 11 THEN 'Novembre'
            WHEN 12 THEN 'Décembre'
            END as Mes, month(anneemois) from distribution where year(anneemois) between {first_year} and {second_year} ORDER BY 1 ASC""")
        myresult_socasp = mycursor_socasp.fetchall()

        columns_socasp = ['Id', 'Marketeur', 'Super', 'JetA1', 'Petrole', 'Gazoil', 'Annee', 'Mois', 'No']
        data_socasp = []

        for x_socasp in myresult_socasp:
            data_socasp.append(x_socasp)

        df_socasp = pd.DataFrame(data_socasp, columns=columns_socasp)
        df_socasp.to_excel('./data/between_distribution_marketeur.xlsx')


    except Exception as Ex:
        print("%s"%(Ex))


    
def betweenDistributionAnnee(host_socasp, user_socasp, password_socasp, database_socasp, first_year, second_year):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        mycursor_socasp = mydb_socasp.cursor()
        mycursor_socasp.execute(f"""select carbuId, marketeur, essence, jet, petrole, gazoil, year(anneemois), CASE MONTH(anneemois) 
            WHEN 1 THEN 'Janvier'
            WHEN 2 THEN 'Février'
            WHEN 3 THEN 'Mars'
            WHEN 4 THEN 'Avril'
            WHEN 5 THEN 'Mai'
            WHEN 6 THEN 'Juin'
            WHEN 7 THEN 'Juillet'
            WHEN 8 THEN 'Aôut'
            WHEN 9 THEN 'Septembre'
            WHEN 10 THEN 'Octobre'
            WHEN 11 THEN 'Novembre'
            WHEN 12 THEN 'Décembre'
            END as Mes, month(anneemois) from distribution where year(anneemois) between {first_year} and {second_year} ORDER BY 1 ASC""")
        myresult_socasp = mycursor_socasp.fetchall()

        columns_socasp = ['Id', 'Marketeur', 'Super', 'JetA1', 'Petrole', 'Gazoil', 'Annee', 'Mois', 'No']
        data_socasp = []

        for x_socasp in myresult_socasp:
            data_socasp.append(x_socasp)

        df_socasp = pd.DataFrame(data_socasp, columns=columns_socasp)
        df_socasp.to_excel('./data/between_distribution_annee_one_two.xlsx')


    except Exception as Ex:
        print("%s"%(Ex))



def createTableOrigine(host_socasp, user_socasp, password_socasp, database_socasp):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        mycursor = mydb_socasp.cursor()
        
        mycursor.execute("create table if not exists origine (id int not null auto_increment primary key, origine varchar(255) not null, origine_data varchar(255) not null)")

    except Exception as Ex:
        print("%s"%(Ex))



def createTableProvenance(host_socasp, user_socasp, password_socasp, database_socasp):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        mycursor = mydb_socasp.cursor()
        
        mycursor.execute("create table if not exists provenance (id int not null auto_increment primary key, provenance varchar(255) not null, provenance_data varchar(255) not null)")

    except Exception as Ex:
        print("%s"%(Ex))




def createTableMarketeur(host_socasp, user_socasp, password_socasp, database_socasp):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        mycursor = mydb_socasp.cursor()
        
        mycursor.execute("create table if not exists marketeur (id int not null auto_increment primary key, marketeur varchar(255) not null, marketeur_data varchar(255) not null)")

    except Exception as Ex:
        print("%s"%(Ex))


def insertOrigine(host_socasp, user_socasp, password_socasp, database_socasp, origine, origine_data):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)
        mycursor = mydb_socasp.cursor()
        
        sql = "INSERT INTO origine (origine, origine_data) VALUES (%s, %s)"
        val = (origine, origine_data)

        mycursor.execute(sql, val)
        mydb_socasp.commit()

        print("Data had been correctly added to the database ...")

    except Exception as Ex:
        print("%s"%(Ex))

    

def insertProvenance(host_socasp, user_socasp, password_socasp, database_socasp, provenance, provenance_data):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)
        mycursor = mydb_socasp.cursor()
        
        sql = "INSERT INTO provenance (provenance, provenance_data) VALUES (%s, %s)"
        val = (provenance, provenance_data)

        mycursor.execute(sql, val)
        mydb_socasp.commit()

        print("Data had been correctly added to the database ...")

    except Exception as Ex:
        print("%s"%(Ex))


    
def insertMarketeur(host_socasp, user_socasp, password_socasp, database_socasp, marketeur, marketeur_data):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)
        mycursor = mydb_socasp.cursor()
        
        sql = "INSERT INTO marketeur (marketeur, marketeur_data) VALUES (%s, %s)"
        val = (marketeur, marketeur_data)

        mycursor.execute(sql, val)
        mydb_socasp.commit()

        print("Data had been correctly added to the database ...")

    except Exception as Ex:
        print("%s"%(Ex))




def connectionOrigine(host_socasp, user_socasp, password_socasp, database_socasp):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        df_importation_excel = './type/origine_data.xlsx'

        mycursor_socasp = mydb_socasp.cursor()
        mycursor_socasp.execute("""select origine, origine_data from origine""")
        myresult_socasp = mycursor_socasp.fetchall()

        columns_socasp = ['Origine', 'Origine Data']
        data_socasp = []

        for x_socasp in myresult_socasp:
            data_socasp.append(x_socasp)

        df_socasp = pd.DataFrame(data_socasp, columns=columns_socasp)
        df_socasp.to_excel(df_importation_excel)

    except Exception as Ex:
        print("%s"%(Ex))



def connectionProvenance(host_socasp, user_socasp, password_socasp, database_socasp):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        df_importation_excel = './type/provenance_data.xlsx'

        mycursor_socasp = mydb_socasp.cursor()
        mycursor_socasp.execute("""select provenance, provenance_data from provenance""")
        myresult_socasp = mycursor_socasp.fetchall()

        columns_socasp = ['Provenance', 'Provenance Data']
        data_socasp = []

        for x_socasp in myresult_socasp:
            data_socasp.append(x_socasp)

        df_socasp = pd.DataFrame(data_socasp, columns=columns_socasp)
        df_socasp.to_excel(df_importation_excel)

    except Exception as Ex:
        print("%s"%(Ex))




def connectionMarketeur(host_socasp, user_socasp, password_socasp, database_socasp):

    try:
        mydb_socasp = mysql.connector.connect(host = host_socasp, user = user_socasp, password =  password_socasp, database = database_socasp)

        df_importation_excel = './type/marketeur_data.xlsx'

        mycursor_socasp = mydb_socasp.cursor()
        mycursor_socasp.execute("""select marketeur, marketeur_data from marketeur""")
        myresult_socasp = mycursor_socasp.fetchall()

        columns_socasp = ['Marketeur', 'Marketeur Data']
        data_socasp = []

        for x_socasp in myresult_socasp:
            data_socasp.append(x_socasp)

        df_socasp = pd.DataFrame(data_socasp, columns=columns_socasp)
        df_socasp.to_excel(df_importation_excel)

    except Exception as Ex:
        print("%s"%(Ex))