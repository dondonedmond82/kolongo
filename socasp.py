import panel as pn
import hvplot.pandas
import warnings
import mysql.connector
import json
import datetime as dt
import os

from bokeh.io import curdoc
from bokeh.models import ColumnDataSource, CustomJS
from bokeh.models.widgets import DataTable, NumberFormatter, TableColumn, Button
from bokeh.layouts import column,row

from tkinter import *
from tkinter.filedialog import asksaveasfile

import ctypes
import pandas as pd
import numpy as np
from sklearn.cluster import KMeans, DBSCAN
from sklearn.preprocessing import StandardScaler

# import create_user as nomusername
import connexion as connect
import login as log
import go as gotodb

warnings.filterwarnings("ignore")
pn.extension('tabulator')
pn.extension('echarts')

# MySQL is optional for dashboard viewing. Local Excel files are the offline data source.
try:
    host_socasp, user_socasp, password_socasp, database_socasp, database_user = log.returnLogDetails()
except Exception as exc:
    print(f"SOCASP : configuration MySQL indisponible ; mode hors ligne activé ({exc}).")
    host_socasp = os.getenv("SOCASP_DB_HOST", "localhost")
    user_socasp = os.getenv("SOCASP_DB_USER", "")
    password_socasp = os.getenv("SOCASP_DB_PASSWORD", "")
    database_socasp = os.getenv("SOCASP_DB_NAME", "socasp")
    database_user = user_socasp

try:
    gotodb.connexion(host_socasp, user_socasp, password_socasp)
    MYSQL_AVAILABLE = True
except Exception as exc:
    MYSQL_AVAILABLE = False
    print(f"SOCASP : serveur MySQL indisponible ; affichage des données locales activé ({exc}).")

# Load the uploaded reference workbooks beside this script first; use legacy paths as fallback.
APP_DIR = os.path.dirname(os.path.abspath(__file__))

def _read_first_existing_excel(candidates, label):
    for candidate in candidates:
        path = candidate if os.path.isabs(candidate) else os.path.join(APP_DIR, candidate)
        if os.path.exists(path):
            print(f"SOCASP : fichier {label} chargé depuis {path}")
            return pd.read_excel(path)
    raise FileNotFoundError(f"Fichier Excel {label} introuvable : {candidates}")

df_importation = _read_first_existing_excel(
    ["importation.xlsx", os.path.join("data", "importation.xlsx")],
    "importation"
)
df = df_importation.copy()
df_distribution = _read_first_existing_excel(
    ["distribution.xlsx", os.path.join("data", "distribution.xlsx")],
    "distribution"
)
try:
    df_user = pd.read_excel(os.path.join(APP_DIR, "data", "utilisateurmotsdepasse.xlsx"))
except Exception:
    df_user = pd.DataFrame(columns=["Hote"])

df_origine = pd.read_excel("./type/origine_data.xlsx")
df_provenance = pd.read_excel("./type/provenance_data.xlsx")
df_marketeur = pd.read_excel("./type/marketeur_data.xlsx")

# ---------------------------------------------------------------------------
# Initialize SOCASP lookup tables automatically at startup.
# Works with existing tables whose text columns may be named Origine,
# Marketeur, Provenance, nom, ville, etc. Values are checked before insert,
# so restarting the application does not duplicate existing entries.
# ---------------------------------------------------------------------------
def initialize_socasp_lookup_tables():
    connection = None
    cursor = None

    def get_text_column(cursor, table_name, candidates):
        """Find an existing text column without assuming the table uses 'nom'."""
        cursor.execute(f"SHOW COLUMNS FROM `{table_name}`")
        columns = [row[0] for row in cursor.fetchall()]
        lower_map = {column.lower(): column for column in columns}

        for candidate in candidates:
            if candidate.lower() in lower_map:
                return lower_map[candidate.lower()]

        raise RuntimeError(
            f"La table MySQL '{table_name}' existe, mais aucune colonne "
            f"compatible n'a été trouvée. Colonnes existantes : {columns}. "
            f"Colonnes recherchées : {candidates}. "
            "Vérifiez la structure de cette table."
        )

    def clean_values(dataframe, column):
        if column not in dataframe.columns:
            raise ValueError(
                f"La colonne '{column}' est absente du fichier Excel correspondant."
            )
        values = dataframe[column].dropna().astype(str).str.strip()
        return sorted({value for value in values if value})

    def insert_if_missing(cursor, table_name, column_name, value):
        # Identifiers come only from SHOW COLUMNS and fixed table names above.
        query = (
            f"SELECT 1 FROM `{table_name}` "
            f"WHERE TRIM(`{column_name}`) = %s LIMIT 1"
        )
        cursor.execute(query, (value,))
        if cursor.fetchone() is None:
            cursor.execute(
                f"INSERT INTO `{table_name}` (`{column_name}`) VALUES (%s)",
                (value,),
            )

    try:
        connection = mysql.connector.connect(
            host=host_socasp,
            user=user_socasp,
            password=password_socasp,
            database=database_socasp,
        )
        cursor = connection.cursor()

        # Create these tables only when they do not already exist.
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS origine (
                id INT UNSIGNED NOT NULL AUTO_INCREMENT,
                nom VARCHAR(100) NOT NULL,
                PRIMARY KEY (id),
                UNIQUE KEY uq_origine_nom (nom)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS marketeur (
                id INT UNSIGNED NOT NULL AUTO_INCREMENT,
                nom VARCHAR(150) NOT NULL,
                PRIMARY KEY (id),
                UNIQUE KEY uq_marketeur_nom (nom)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS provenance (
                id INT UNSIGNED NOT NULL AUTO_INCREMENT,
                ville VARCHAR(150) NOT NULL,
                pays VARCHAR(150) NULL,
                PRIMARY KEY (id),
                UNIQUE KEY uq_provenance_ville_pays (ville, pays)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """)

        # Detect the actual text column in each existing table.
        origine_col = get_text_column(
            cursor, "origine", ["nom", "Origine", "type_origine", "libelle"]
        )
        marketeur_col = get_text_column(
            cursor, "marketeur", ["nom", "Marketeur", "marqueteur", "libelle"]
        )
        provenance_col = get_text_column(
            cursor, "provenance", ["ville", "Provenance", "nom", "libelle"]
        )

        for value in clean_values(df_origine, "Origine"):
            insert_if_missing(cursor, "origine", origine_col, value)

        for value in clean_values(df_marketeur, "Marketeur"):
            insert_if_missing(cursor, "marketeur", marketeur_col, value)

        for value in clean_values(df_provenance, "Provenance"):
            insert_if_missing(cursor, "provenance", provenance_col, value)

        connection.commit()
        print(
            "SOCASP : tables de référence vérifiées et initialisées "
            "sans réinsérer les valeurs existantes."
        )

    except Exception as exc:
        if connection and connection.is_connected():
            connection.rollback()
        print(
            "SOCASP : initialisation MySQL ignorée. "
            "Le tableau de bord continue avec les fichiers Excel locaux. "
            f"Détail : {exc}"
        )
    finally:
        if cursor:
            cursor.close()
        if connection and connection.is_connected():
            connection.close()


# Automatic initialization each time Panel starts this application.
initialize_socasp_lookup_tables()

df_importation = df_importation[['Origine', 'Provenance', 'Marketeur', 'Annee', 'Mois', 'Super', 'JetA1', 'Petrole', 'Gazoil']]
df_distribution = df_distribution[[col for col in ['No', 'Marketeur', 'Annee', 'Mois', 'Super', 'JetA1', 'Petrole', 'Gazoil'] if col in df_distribution.columns]]

df = df[[ 'No', 'Origine', 'Provenance', 'Marketeur', 'Annee', 'Mois', 'Super', 'JetA1', 'Petrole', 'Gazoil']]

df['Annee'] = df['Annee'].astype("str")

df_temp = df[['Origine', 'Provenance', 'Marketeur', 'Annee', 'Mois', 'Super', 'JetA1', 'Petrole', 'Gazoil']]
df_temp['Annee'] = df_temp['Annee'].astype("str")

# Default date range spans the available records so the initial view includes all data.
def _dataset_date_bounds(data):
    years = pd.to_numeric(data["Annee"], errors="coerce").dropna()
    months = pd.to_numeric(data["Mois"], errors="coerce").dropna() if "Mois" in data.columns else pd.Series(dtype=float)
    if years.empty:
        return dt.date(2000, 1, 1), dt.date.today()
    min_year, max_year = int(years.min()), int(years.max())
    min_month = int(months.min()) if not months.empty else 1
    max_month = int(months.max()) if not months.empty else 12
    return dt.date(min_year, min_month, 1), dt.date(max_year, max_month, 28)

IMPORTATION_DATE_MIN, IMPORTATION_DATE_MAX = _dataset_date_bounds(df_importation)
DISTRIBUTION_DATE_MIN, DISTRIBUTION_DATE_MAX = _dataset_date_bounds(df_distribution)



#########################################################################################################################################################

show_distribution_marketeur = pn.widgets.Button(name="Distribution par Marketeur", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
show_distribution_marketeur.on_click(lambda event: show_page("Page3"))

show_distribution_annuelle = pn.widgets.Button(name="Distribution Annuelle", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
show_distribution_annuelle.on_click(lambda event: show_page("Page4"))

#########################################################################################################################################################

create_user = pn.widgets.Button(name="Creer Utilisateur", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
create_user.on_click(lambda event: createUsername())

cancel_user = pn.widgets.Button(name="Annuler", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
cancel_user.on_click(lambda event: cancelUsername())

#########################################################################################################################################################

date_distribution_marketeur_annee_start = pn.widgets.DatePicker(name='Date du Debut', value=DISTRIBUTION_DATE_MIN)
date_distribution_marketeur_annee_end = pn.widgets.DatePicker(name='Date de la Fin', value=DISTRIBUTION_DATE_MAX)

date_distribution_annee_marketeur_start = pn.widgets.DatePicker(name='Date du Debut par Marketeur', value=DISTRIBUTION_DATE_MIN)
date_distribution_annee_marketeur_end   = pn.widgets.DatePicker(name='Date de la Fin par Marketeur', value=DISTRIBUTION_DATE_MAX)

#########################################################################################################################################################

show_UserName = pn.widgets.Button(name="Details Utilisateur", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
show_UserName.on_click(lambda event: show_page("Page7"))

#########################################################################################################################################################

insert_importation = pn.widgets.Button(name="Ajout Importation", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
insert_importation.on_click(lambda event: show_page("Page5"))

insert_distribution = pn.widgets.Button(name="Ajout Distribution", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
insert_distribution.on_click(lambda event: show_page("Page6"))

#########################################################################################################################################################

unique_annee_importation =  list(df_importation['Annee'].unique())

select_annee_importation_start = pn.widgets.Select(name="Année Importation", options=["All"] + unique_annee_importation, value="All")
select_annee_importation_end = pn.widgets.Select(name="Année Importation", options=["All"] + unique_annee_importation, value="All")


unique_marketeur_importation =  list(df_importation['Marketeur'].unique())

select_marketeur_importation_start = pn.widgets.Select(name="Marketeur Importation", options=["All"] + unique_marketeur_importation, value="All")
select_marketeur_importation_end = pn.widgets.Select(name="Marketeur Importation", options=["All"] + unique_marketeur_importation, value="All")

select_marketeur_annee = pn.widgets.Select(name="Année Importation", options=["All"] + unique_annee_importation, value="All")
select_marketeur_end = pn.widgets.Select(name="Année Importation", options=["All"] + unique_annee_importation, value="All")

#########################################################################################################################################################

date_importation = pn.widgets.DatePicker(name='Date', value=dt.date(2000, 1, 1))
date_distribution = pn.widgets.DatePicker(name='Date', value=dt.date(2000, 1, 1))

#########################################################################################################################################################

input_username = pn.widgets.TextInput(name=f"Nom d'Utilisateur", value='')
input_password = pn.widgets.PasswordInput(name=f"Mots de Passe", value='')
input_repassword = pn.widgets.PasswordInput(name=f"Re-Mots de Passe", value='')

#########################################################################################################################################################

input_essence = pn.widgets.TextInput(name="Essence", value='')
input_jet = pn.widgets.TextInput(name="JetA1", value='')
input_petrole = pn.widgets.TextInput(name="Petrole", value='')
input_gazoil = pn.widgets.TextInput(name="Gazoil", value='')

#########################################################################################################################################################

input_essence_distribution = pn.widgets.TextInput(name="Essence", value='')
input_jet_distribution = pn.widgets.TextInput(name="JetA1", value='')
input_petrole_distribution = pn.widgets.TextInput(name="Petrole", value='')
input_gazoil_distribution = pn.widgets.TextInput(name="Gazoil", value='')

#########################################################################################################################################################

JScode_fetch = """
var filename = '2025-2-22-8-5-16_importation_Tamoil.xlsx';

fetch('/excel/importation_produit/marketeur/2025-2-22-8-5-16_importation_Tamoil.xlsx', {cache: "no-store"}).then(response => response.blob())
                .then(blob => {
                    //addresses IE
                    if (navigator.msSaveBlob) {
                        navigator.msSaveBlob(blob, filename);
                    }

                    else {
                        var link = document.createElement("a");
                        link.href = URL.createObjectURL(blob);
                        
                        link.download = filename;
                        link.target = "_blank";
                        link.style.visibility = 'hidden';
                        link.dispatchEvent(new MouseEvent('click'))
                        URL.revokeObjectURL(url);
                    }
                    return response.text();
                });

"""


# Created Button for Export

Exportbutton = pn.widgets.Button(name="Telecharger", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
Exportbutton.js_on_click(code = JScode_fetch)


def clearOrigine():
    
    input_origine_one.value = ""
    input_origine_two.value = ""


def insertOrigine():

    origine_one = input_origine_one.value
    origine_one = f"{origine_one}"

    origine_two = input_origine_two.value
    origine_two = f"{origine_two}"

    try:
        if origine_one == origine_two:
        
            connect.insertOrigine(host_socasp, user_socasp, password_socasp, database_socasp, origine_one, origine_two)
            connect.connectionOrigine(host_socasp, user_socasp, password_socasp, database_socasp)

            clearOrigine()

        else:
            print(f"L'origine n'est pas la meme.")

    except Exception as Ex:
            print("%s"%(Ex))



def clearProvenance():
    
    input_provenance_one.value = ""
    input_provenance_two.value = ""



def insertProvenance():

    provenance_one = input_provenance_one.value
    provenance_one = f"{provenance_one}"

    provenance_two = input_provenance_two.value
    provenance_two = f"{provenance_two}"

    try:
        if provenance_one == provenance_two:
        
            connect.insertProvenance(host_socasp, user_socasp, password_socasp, database_socasp, provenance_one, provenance_two)
            connect.connectionProvenance(host_socasp, user_socasp, password_socasp, database_socasp)

            clearProvenance()

        else:
            print(f"La provenance n'est pas la meme.")

    except Exception as Ex:
            print("%s"%(Ex))



def clearMarketeur():
    
    input_marketeur_one.value = ""
    input_marketeur_two.value = ""



def insertMarketeur():

    marketeur_one = input_marketeur_one.value
    marketeur_one = f"{marketeur_one}"

    marketeur_two = input_marketeur_two.value
    marketeur_two = f"{marketeur_two}"

    try:
        if marketeur_one == marketeur_two:
        
            connect.insertMarketeur(host_socasp, user_socasp, password_socasp, database_socasp, marketeur_one, marketeur_two)
            connect.connectionMarketeur(host_socasp, user_socasp, password_socasp, database_socasp)

            clearMarketeur()

        else:
            print(f"Le Marketeur n'est pas la meme.")

    except Exception as Ex:
            print("%s"%(Ex))


#########################################################################################################################################################

input_origine_one = pn.widgets.TextInput(name="Origine", value='')
input_provenance_one = pn.widgets.TextInput(name="Provenance", value='')
input_marketeur_one = pn.widgets.TextInput(name="Marketeur", value='')

input_origine_two = pn.widgets.TextInput(name="Re-Origine", value='')
input_provenance_two = pn.widgets.TextInput(name="Re-Provenance", value='')
input_marketeur_two = pn.widgets.TextInput(name="Re-Marketeur", value='')

validate_origine = pn.widgets.Button(name="Valider", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
validate_provenance = pn.widgets.Button(name="Valider", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
validate_marketeur = pn.widgets.Button(name="Valider", button_type="primary", icon="clipboard-data", styles={"width": "100%"})

cancel_origine = pn.widgets.Button(name="Annuler", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
cancel_provenance = pn.widgets.Button(name="Annuler", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
cancel_marketeur = pn.widgets.Button(name="Annuler", button_type="primary", icon="clipboard-data", styles={"width": "100%"})


validate_origine.on_click(lambda event: insertOrigine())
validate_provenance.on_click(lambda event: insertProvenance())
validate_marketeur.on_click(lambda event: insertMarketeur())


cancel_origine.on_click(lambda event: clearOrigine())
cancel_provenance.on_click(lambda event: clearProvenance())
cancel_marketeur.on_click(lambda event: clearMarketeur())

#########################################################################################################################################################

unique_origine = list(df_origine['Origine'].unique())
unique_provenance =   list(df_provenance['Provenance'].unique())
unique_marketeur =  list(df_marketeur['Marketeur'].unique())

unique_annee =  list(df['Annee'].unique())

select_origine = pn.widgets.Select(name="Origine", options=unique_origine)
select_provenance = pn.widgets.Select(name="Provenance", options=unique_provenance)
select_marketeur = pn.widgets.Select(name="Marketeur", options=unique_marketeur)

select_marketeur_distribution_insert = pn.widgets.Select(name="Marketeur", options=unique_marketeur)

#########################################################################################################################################################

valider = pn.widgets.Button(name="Valider", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
annuler = pn.widgets.Button(name="Annuler", button_type="primary", icon="clipboard-data", styles={"width": "100%"})

valider_distribution = pn.widgets.Button(name="Valider", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
annuler_distribution = pn.widgets.Button(name="Annuler", button_type="primary", icon="clipboard-data", styles={"width": "100%"})

actualiser_distribution = pn.widgets.Button(name="Actualiser", button_type="primary", icon="clipboard-data", styles={"width": "100%"})

insert_username = pn.widgets.Button(name="Compte Utilisateur", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
insert_username.on_click(lambda event: show_page("Page8"))

#########################################################################################################################################################


unique_user =  list(df_user['Hote'].unique())
select_user = pn.widgets.Select(name="Hote", options=["All"] + unique_user, value="All")

@pn.depends(symbol_user=select_user)
def create_username(symbol_user):
    if symbol_user == "All":
        data = df_user
    else:
        data = df_user[df_user["Hote"] == symbol_user]
    return data.hvplot.table(width=1050)

#########################################################################################################################################################

def clearImportation():
    
    input_essence.value = ""
    input_jet.value = ""
    input_petrole.value = ""
    input_gazoil.value = ""


def insertImportation():

    origine = select_origine.value
    origine = f"{origine}"

    provenance = select_provenance.value
    provenance = f"{provenance}"

    marketeur = select_marketeur.value
    marketeur = f"{marketeur}"

    essence = input_essence.value
    essence = f"{essence}"

    jet = input_jet.value
    jet = f"{jet}"

    petrole = input_petrole.value
    petrole = f"{petrole}"

    gazoil = input_gazoil.value
    gazoil = f"{gazoil}"

    anneemois = date_importation.value
    anneemois = f"{anneemois}"

    try:
        connect.insertImportation(host_socasp, user_socasp, password_socasp, database_socasp, origine, provenance, marketeur, essence, jet, petrole, gazoil, anneemois)
        df_importation_excel =  pd.read_excel("./data/importation.xlsx")

        clearImportation()

    except Exception as Ex:
        print("%s"%(Ex))

#########################################################################################################################################################

def clearDistribution():
    
    input_essence_distribution.value = ""
    input_jet_distribution.value = ""
    input_petrole_distribution.value = ""
    input_gazoil_distribution.value = ""


def insertDistribution():

    marketeur_distribution = select_marketeur_distribution_insert.value
    essence_distribution = input_essence_distribution.value
    jet_distribution = input_jet_distribution.value
    petrole_distribution = input_petrole_distribution.value
    gazoil_distribution = input_gazoil_distribution.value
    anneemois_distribution = date_distribution.value

    try:
        connect.insertDistribution(host_socasp, user_socasp, password_socasp, database_socasp, marketeur_distribution, essence_distribution, jet_distribution, petrole_distribution, gazoil_distribution, anneemois_distribution)
        df_distribution_excel  = pd.read_excel("./data/distribution.xlsx")

        clearDistribution()

    except Exception as Ex:
        print("%s"%(Ex))

#########################################################################################################################################################

unique_marketeur =  list(df_importation['Marketeur'].unique())
unique_annee =  list(df_importation['Annee'].unique())

select_annee_importation = pn.widgets.Select(name="Marketeur", options=unique_marketeur)
select_importation_marketeur = pn.widgets.Select(name="Marketeur", options=["All"] + unique_marketeur, value="All")

select_importation_annee_start = pn.widgets.Select(name="Annee", options=["All"] + unique_annee, value="All")
select_importation_annee_end = pn.widgets.Select(name="Annee", options=["All"] + unique_annee, value="All")


def betweenDateImportation():
    """Filter local importation Excel data by the selected date range (offline-safe)."""
    data = df_importation.copy()
    data["Annee"] = pd.to_numeric(data["Annee"], errors="coerce")
    data["Mois"] = pd.to_numeric(data["Mois"], errors="coerce")

    start_date = date_importation_annee_marketeur_start.value
    end_date = date_importation_annee_marketeur_end.value

    if start_date is not None:
        data = data[
            (data["Annee"] > start_date.year)
            | ((data["Annee"] == start_date.year) & (data["Mois"] >= start_date.month))
        ]
    if end_date is not None:
        data = data[
            (data["Annee"] < end_date.year)
            | ((data["Annee"] == end_date.year) & (data["Mois"] <= end_date.month))
        ]

    data["Annee"] = data["Annee"].astype("Int64").astype(str)
    return data


#########################################################################################################################################################

def betweenDateImportationMarketeur():
    """Return all importation rows inside the selected date range."""
    data = df.copy()
    data["Annee"] = pd.to_numeric(data["Annee"], errors="coerce")
    data["Mois"] = pd.to_numeric(data["Mois"], errors="coerce")
    start_date = date_importation_annee_marketeur_start.value
    end_date = date_importation_annee_marketeur_end.value
    if start_date is not None:
        data = data[(data["Annee"] > start_date.year) |
                    ((data["Annee"] == start_date.year) & (data["Mois"] >= start_date.month))]
    if end_date is not None:
        data = data[(data["Annee"] < end_date.year) |
                    ((data["Annee"] == end_date.year) & (data["Mois"] <= end_date.month))]
    data["Annee"] = data["Annee"].astype("Int64").astype(str)
    return data


def betweenDateImportationAnnee():
    """Return all importation rows; the dropdowns apply optional year bounds."""
    return df.copy()



def betweenDateImportationAnnee():
    """Return local importation rows; annual dropdowns filter this data offline."""
    data = df_importation.copy()
    if "No" not in data.columns:
        data.insert(0, "No", range(1, len(data) + 1))
    columns = [
        col for col in
        ["No", "Origine", "Provenance", "Marketeur", "Annee", "Mois",
         "Super", "JetA1", "Petrole", "Gazoil"]
        if col in data.columns
    ]
    return data[columns]


unique_marketeur =  list(df_distribution['Marketeur'].unique())
unique_annee =  list(df_distribution['Annee'].unique())

select_annee_distribution = pn.widgets.Select(name="Marketeur", options=unique_marketeur)
select_distribution_marketeur = pn.widgets.Select(name="Marketeur", options=["All"] + unique_marketeur, value="All")

select_distribution_annee_start = pn.widgets.Select(name="Annee", options=unique_annee, value=unique_annee[0] if unique_annee else None)
select_distribution_annee_end = pn.widgets.Select(name="Annee", options=unique_annee, value=unique_annee[-1] if unique_annee else None)


#########################################################################################################################################################

def betweenDateDistributionAnnee():
    """Return all distribution rows; the dropdowns apply optional year bounds."""
    return df_distribution.copy()



def betweenDateDistributionMarketeur():
    """Return all distribution rows inside the selected date range."""
    data = df_distribution.copy()
    data["Annee"] = pd.to_numeric(data["Annee"], errors="coerce")
    data["Mois"] = pd.to_numeric(data["Mois"], errors="coerce")
    start_date = date_distribution_annee_marketeur_start.value
    end_date = date_distribution_annee_marketeur_end.value
    if start_date is not None:
        data = data[(data["Annee"] > start_date.year) |
                    ((data["Annee"] == start_date.year) & (data["Mois"] >= start_date.month))]
    if end_date is not None:
        data = data[(data["Annee"] < end_date.year) |
                    ((data["Annee"] == end_date.year) & (data["Mois"] <= end_date.month))]
    data["Annee"] = data["Annee"].astype("Int64").astype(str)
    return data



#########################################################################################################################################################


@pn.depends(symbol_marketeur_distribution = select_distribution_marketeur)

def create_distribution_annee_marketeur(symbol_marketeur_distribution):
    df_distribution_year_start_end_marketeur = betweenDateDistributionMarketeur()
    data = df_distribution_year_start_end_marketeur
    if symbol_marketeur_distribution != "All":
        data = data[data["Marketeur"] == symbol_marketeur_distribution]
    return data.hvplot.table(width=1050)


#########################################################################################################################################################

def carburantPivotImportationMarketeur(df_importation_marketeur):
    df_pivot = pd.pivot_table(data=df_importation_marketeur, index=['Marketeur', 'Annee'] ,values=['Super', 'JetA1', 'Petrole', 'Gazoil'], columns=['No', 'Mois'], aggfunc='sum', margins=True, margins_name='Cumul')
    
    df_pivot['Super'] = df_pivot['Super'].fillna(0)
    df_pivot['JetA1'] = df_pivot['JetA1'].fillna(0)
    df_pivot['Petrole'] = df_pivot['Petrole'].fillna(0)
    df_pivot['Gazoil'] = df_pivot['Gazoil'].fillna(0)

    df_pivot_stack = pd.DataFrame(df_pivot.stack(0))

    return df_pivot_stack


def carburantPivotImportationAnnee(df_importation_annee):
        df_pivot = pd.pivot_table(data=df_importation_annee, index=['Marketeur', 'Annee'] ,values=['Super', 'JetA1', 'Petrole', 'Gazoil'], columns=['No', 'Mois'], aggfunc='sum', margins=True, margins_name='Cumul')
    
        df_pivot['Super'] = df_pivot['Super'].fillna(0)
        df_pivot['JetA1'] = df_pivot['JetA1'].fillna(0)
        df_pivot['Petrole'] = df_pivot['Petrole'].fillna(0)
        df_pivot['Gazoil'] = df_pivot['Gazoil'].fillna(0)

        df_pivot_stack = pd.DataFrame(df_pivot.stack(0))

        return df_pivot_stack


def carburantPivotDistributionMarketeur(df_distribution_marketeur):
        df_pivot = pd.pivot_table(data=df_distribution_marketeur, index=['Marketeur', 'Annee'] ,values=['Super', 'JetA1', 'Petrole', 'Gazoil'], columns=['No', 'Mois'], aggfunc='sum', margins=True, margins_name='Cumul')
    
        df_pivot['Super'] = df_pivot['Super'].fillna(0)
        df_pivot['JetA1'] = df_pivot['JetA1'].fillna(0)
        df_pivot['Petrole'] = df_pivot['Petrole'].fillna(0)
        df_pivot['Gazoil'] = df_pivot['Gazoil'].fillna(0)

        df_pivot_stack = pd.DataFrame(df_pivot.stack(0))

        return df_pivot_stack


def carburantPivotDistributionAnnee(df_distribution_annee):
    df_pivot = pd.pivot_table(data=df_distribution_annee, index=['Marketeur', 'Annee'] ,values=['Super', 'JetA1', 'Petrole', 'Gazoil'], columns=['No', 'Mois'], aggfunc='sum', margins=True, margins_name='Cumul')
    
    df_pivot['Super'] = df_pivot['Super'].fillna(0)
    df_pivot['JetA1'] = df_pivot['JetA1'].fillna(0)
    df_pivot['Petrole'] = df_pivot['Petrole'].fillna(0)
    df_pivot['Gazoil'] = df_pivot['Gazoil'].fillna(0)

    df_pivot_stack = pd.DataFrame(df_pivot.stack(0))

    return df_pivot_stack


#########################################################################################################################################################

import datetime
    
year        = datetime.date.today().year     # the current year
month       = datetime.date.today().month    # the current month
date       = datetime.date.today().day      # the current day
hour        = datetime.datetime.now().hour   # the current hour
minute      = datetime.datetime.now().minute # the current minute
second     = datetime.datetime.now().second #the current second

def exportImportationMarketeur():

    df_importation_marketeur = betweenDateImportationMarketeur()

    if select_importation_marketeur.value != "All":
        df_importation_marketeur = df_importation_marketeur[
            df_importation_marketeur["Marketeur"] == select_importation_marketeur.value
        ]
    df_pivot_marketeur_socasp = carburantPivotImportationMarketeur(df_importation_marketeur)


    fichier_marketeur = f"./excel/importation_produit/marketeur/{year}-{month}-{date}-{hour}-{minute}-{second}_importation_{select_importation_marketeur.value}.xlsx"
    df_pivot_marketeur_socasp.to_excel(fichier_marketeur)


ExportBox_importation_marketeur = pn.widgets.Button(name="Sauvegarde Importation par Marketeur", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
ExportBox_importation_marketeur.on_click(lambda event: exportImportationMarketeur())



def exportImportationAnnee():

    df_importation_annee = betweenDateImportationAnnee()
    df_pivot_annee_socasp = carburantPivotImportationAnnee(df_importation_annee)

    fichier_annee = f"./excel/importation_produit/annee/{year}-{month}-{date}-{hour}-{minute}-{second}_importation_{select_importation_annee_start.value}_{select_importation_annee_end.value}.xlsx"
    df_pivot_annee_socasp.to_excel(fichier_annee)


ExportBox_importation_annee = pn.widgets.Button(name="Sauvegarde Importation Annuelle", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
ExportBox_importation_annee.on_click(lambda event: exportImportationAnnee())


def exportDistributionMarketeur():

    df_distribution_marketeur = betweenDateDistributionMarketeur()
    if select_distribution_marketeur.value != "All":
        df_distribution_marketeur = df_distribution_marketeur[
            df_distribution_marketeur["Marketeur"] == select_distribution_marketeur.value
        ]
    df_pivot_marketeur_socasp = carburantPivotDistributionMarketeur(df_distribution_marketeur)

    fichier_marketeur = f"./excel/distribution_produit/marketeur/{year}-{month}-{date}-{hour}-{minute}-{second}_distribution_{select_distribution_marketeur.value}.xlsx"
    df_pivot_marketeur_socasp.to_excel(fichier_marketeur)


ExportBox_distribution_marketeur = pn.widgets.Button(name="Sauvegarde Distribution par Marketeur", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
ExportBox_distribution_marketeur.on_click(lambda event: exportDistributionMarketeur())



def exportDistributionAnnee():

    df_distribution_annee = betweenDateDistributionAnnee()
    df_pivot_annee_socasp = carburantPivotDistributionAnnee(df_distribution_annee)

    fichier_annee = f"./excel/distribution_produit/annee/{year}-{month}-{date}-{hour}-{minute}-{second}_distribution_{select_distribution_annee_start.value}_{select_distribution_annee_end.value}.xlsx"
    df_pivot_annee_socasp.to_excel(fichier_annee)


ExportBox_distribution_annee = pn.widgets.Button(name="Sauvegarde Distribution Annuelle", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
ExportBox_distribution_annee.on_click(lambda event: exportDistributionAnnee())

#########################################################################################################################################################

@pn.depends(symbol_annee_distribution = select_distribution_annee_end)
def create_annee_distribution(symbol_annee_distribution):
    df_distribution_year_start_end = betweenDateDistributionAnnee()
    data = df_distribution_year_start_end.copy()
    years = pd.to_numeric(data["Annee"], errors="coerce")
    if select_distribution_annee_start.value != "All":
        data = data[years >= float(select_distribution_annee_start.value)]
        years = pd.to_numeric(data["Annee"], errors="coerce")
    if symbol_annee_distribution != "All":
        data = data[years <= float(symbol_annee_distribution)]
    return data.hvplot.table(width=1050)


ExportBox_marketeur = pn.widgets.Button(name="Sauvegarde Annuelle", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
ExportBox_marketeur.on_click(lambda event: betweenDateDistributionAnnee())

#########################################################################################################################################################

show_importation_marketeur = pn.widgets.Button(name="Importation par Marketeur", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
show_importation_marketeur.on_click(lambda event: show_page("Page1"))

show_importation_annuelle = pn.widgets.Button(name="Importation Annuelle", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
show_importation_annuelle.on_click(lambda event: show_page("Page2"))

insert_general = pn.widgets.Button(name="Ajout Details", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
insert_general.on_click(lambda event: show_page("Page9"))

#########################################################################################################################################################

date_importation_marketeur_annee_start = pn.widgets.DatePicker(name='Date du Debut', value=IMPORTATION_DATE_MIN)
date_importation_marketeur_annee_end = pn.widgets.DatePicker(name='Date de la Fin', value=IMPORTATION_DATE_MAX)

date_importation_annee_marketeur_start = pn.widgets.DatePicker(name='Date du Debut par Marketeur', value=IMPORTATION_DATE_MIN)
date_importation_annee_marketeur_end   = pn.widgets.DatePicker(name='Date de la Fin par Marketeur', value=IMPORTATION_DATE_MAX)

date_distribution_marketeur_annee_start = pn.widgets.DatePicker(name='Date du Debut', value=DISTRIBUTION_DATE_MIN)
date_distribution_marketeur_annee_end = pn.widgets.DatePicker(name='Date de la Fin', value=DISTRIBUTION_DATE_MAX)

date_distribution_annee_marketeur_start = pn.widgets.DatePicker(name='Date du Debut par Marketeur', value=DISTRIBUTION_DATE_MIN)
date_distribution_annee_marketeur_end   = pn.widgets.DatePicker(name='Date de la Fin par Marketeur', value=DISTRIBUTION_DATE_MAX)

#########################################################################################################################################################

valider = pn.widgets.Button(name="Valider", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
annuler = pn.widgets.Button(name="Annuler", button_type="primary", icon="clipboard-data", styles={"width": "100%"})

valider_distribution = pn.widgets.Button(name="Valider", button_type="primary", icon="clipboard-data", styles={"width": "100%"})
annuler_distribution = pn.widgets.Button(name="Annuler", button_type="primary", icon="clipboard-data", styles={"width": "100%"})

actualiser_distribution = pn.widgets.Button(name="Actualiser", button_type="primary", icon="clipboard-data", styles={"width": "100%"})

valider.on_click(lambda event: insertImportation())
annuler.on_click(lambda event: clearImportation())

valider_distribution.on_click(lambda event: insertDistribution())
annuler_distribution.on_click(lambda event: clearDistribution())

#########################################################################################################################################################

def cancelUsername():

    input_username.value = ""
    input_password.value = ""


def createUsername():

    usernamesocasp  = input_username.value
    passwordsocasp   = input_password.value
    
    try:
        connect.createUserSocasp(host_socasp, user_socasp, password_socasp, database_socasp, usernamesocasp, passwordsocasp)
        connect.insertUser(host_socasp, user_socasp, password_socasp, database_socasp, usernamesocasp, passwordsocasp)
        connect.connectionUser(host_socasp, user_socasp, password_socasp, database_socasp)

        cancelUsername()

    except Exception as Ex:
        print("%s"%(Ex))



@pn.depends(symbol_marketeur_importation = select_importation_marketeur)
def create_importation_annee_marketeur(symbol_marketeur_importation):
    df_importation_year_start_end_marketeur = betweenDateImportationMarketeur()
    data = df_importation_year_start_end_marketeur
    if symbol_marketeur_importation != "All":
        data = data[data["Marketeur"] == symbol_marketeur_importation]
    return data.hvplot.table(width=1050)



@pn.depends(symbol_annee_importation = select_importation_annee_end)
def create_annee_importation(symbol_annee_importation):
    df_importation_year_start_end = betweenDateImportationAnnee()
    data = df_importation_year_start_end.copy()
    years = pd.to_numeric(data["Annee"], errors="coerce")
    if select_importation_annee_start.value != "All":
        data = data[years >= float(select_importation_annee_start.value)]
        years = pd.to_numeric(data["Annee"], errors="coerce")
    if symbol_annee_importation != "All":
        data = data[years <= float(symbol_annee_importation)]
    return data.hvplot.table(width=1050)


#########################################################################################################################################################

CURRENT_USERNAME = ""

def create_welcome_header():
    """Create the persistent French welcome header for every authenticated interface."""
    safe_username = (
        CURRENT_USERNAME.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#39;")
    )
    return pn.pane.Markdown(
        f"<div style='background:#17148E;color:#ffffff;padding:16px 22px;"
        f"border-radius:10px;font-size:22px;font-weight:700;margin-bottom:16px;'>"
        f"Bienvenue, {safe_username}</div>",
        sizing_mode="stretch_width",
    )

def show_page(page_key):
    main_area.clear()
    main_area.append(create_welcome_header())
    main_area.append(mapping[page_key])
    
#########################################################################################################################################################

def CreatePage1():

    return pn.Column(
                    pn.pane.Markdown("## Importation des Données des Carburants par Marketeur"),
            pn.Row(
                pn.Column(
                    date_importation_annee_marketeur_start,
                ),

                pn.Column(
                    date_importation_annee_marketeur_end,
                ),

                pn.Column(
                    select_importation_marketeur,   
                )
                
            ),

            pn.Row(
                pn.Column(
                    create_importation_annee_marketeur,   
                )
            ),

            pn.Row(
                pn.Column(
                    ExportBox_importation_marketeur,   
                )
            ),
        )

def CreatePage2():

    return pn.Column(
                pn.pane.Markdown("## Importation des Données des Carburants par Année"),

            pn.Row(
                pn.Column(
                    select_importation_annee_start,
                ),

                pn.Column(
                    select_importation_annee_end,
                )
            ),

            pn.Row(
                pn.Column(
                    create_annee_importation,
                )
            ),

            pn.Row(
                pn.Column(
                    ExportBox_importation_annee,  
                )
            ),
        )


def CreatePage3():

    return pn.Column(

            pn.pane.Markdown("## Distribution des Données des Carburants par Marketeur"),

            pn.Row(
                pn.Column(
                    date_distribution_annee_marketeur_start,
                ),

                pn.Column(
                    date_distribution_annee_marketeur_end,
                ),

                pn.Column(
                    select_distribution_marketeur,   
                )
                
            ),

            pn.Row(
                pn.Column(
                    create_distribution_annee_marketeur,   
                )
            ),

            pn.Row(
                pn.Column(
                    ExportBox_distribution_marketeur,   
                )
            ),
        )

def CreatePage4():

    return pn.Column(

            pn.pane.Markdown("## Distribution des Données des Carburants par Année"),

            pn.Row(
                pn.Column(
                    select_distribution_annee_start,
                ),

                pn.Column(
                    select_distribution_annee_end,
                )
            ),

            pn.Row(
                pn.Column(
                    create_annee_distribution,
                )
            ),

            pn.Row(
                pn.Column(
                    ExportBox_distribution_annee,   
                )
            ),
        )


def CreatePage5():

        return pn.Column(

            pn.pane.Markdown("## Insertion des Données des Carburants par Importation"),

        pn.Column( 
            pn.Row(
                pn.Column(
                    pn.Row(date_importation),
                ),

                pn.Column(
                    pn.Row(select_origine),
                ),

                pn.Column(
                    pn.Row(select_provenance),
                )
            ),

            pn.Row(
                pn.Column(
                    pn.Row(select_marketeur),
                ),

                pn.Column(
                    pn.Row(input_essence),
                ),

                pn.Column(
                    pn.Row(input_jet),
                )
            ),

            pn.Row(
                pn.Column(
                    pn.Row(input_petrole),
                    pn.Row(valider),
                ),

                pn.Column(
                    pn.Row(input_gazoil),
                    pn.Row(annuler),
                ),
            ),
        ),
    )


def CreatePage6():
        return pn.Column(

            pn.pane.Markdown("## Insertion des Données des Carburants par Distribution"),

        pn.Column( 
            pn.Row(
                pn.Column(
                    pn.Row(date_distribution),
                ),

                pn.Column(
                    pn.Row(select_marketeur_distribution_insert),
                ),

                pn.Column(
                    pn.Row(input_essence_distribution),
                )
            ),

            pn.Row(
                pn.Column(
                    pn.Row(input_jet_distribution),
                    pn.Row(valider_distribution),
                ),

                pn.Column(
                    pn.Row(input_gazoil_distribution),
                    pn.Row(annuler_distribution),
                ),

                pn.Column(
                    pn.Row(input_petrole_distribution),
                )
            )
        )
    )


def CreatePage7():
    return pn.Column(

                pn.pane.Markdown("## Details de droits d'accés des Utilisateurs"),

                pn.Row(
                    pn.Column(
                        select_user,
                    )
                ),

                pn.Row(
                    pn.Column(
                        create_username,
                    )
                )
            )


def CreatePage8():
    return pn.Column(

            pn.pane.Markdown("## Creation de droits d'accés des Utilisateurs"),
            
            pn.Row(
                pn.Column(
                    input_username, create_user,
                ),

                pn.Column(
                    input_password, cancel_user,
                )
            )
        )


def CreatePage9():
    return pn.Column(

            pn.pane.Markdown("## Insertion des Informations telles que l'Origine, Provenance et Marketeur."),

            pn.Row(
                pn.Column(
                    input_origine_one, input_origine_two, validate_origine, cancel_origine,
                ),

                pn.Column(
                    input_provenance_one, input_provenance_two, validate_provenance, cancel_provenance,
                ),

                pn.Column(
                    input_marketeur_one, input_marketeur_two, validate_marketeur, cancel_marketeur,
                )
            )
        )

#########################################################################################################################################################


# ---------------------------------------------------------------------------
# SOCASP - Charte graphique inspirée du logo et tableau de bord analytique
# ---------------------------------------------------------------------------
# Palette SOCASP monochrome : noir et blanc uniquement.
pn.config.raw_css.append("""
:root {
    --socasp-green: #17148E;
    --socasp-red: #17148E;
    --socasp-bg: #ffffff;
    --socasp-ink: #17148E;
}
body, .bk-root { background: #ffffff; font-family: Inter, "Segoe UI", Arial, sans-serif; }
/* Override Bootstrap/Panel defaults so no blue remains in the menu/sidebar. */
:root {
    --bs-primary: #17148E;
    --bs-link-color: #17148E;
    --bs-link-hover-color: #17148E;
    --bs-btn-bg: #17148E;
    --bs-btn-border-color: #17148E;
}
#sidebar, .sidebar, .bk-sidebar, .pn-sidebar, .bk-panel-models-layout .sidebar,
.bk-panel-models-layout #sidebar, .bk-Column.socasp-sidebar {
    background-color: #ffffff !important;
    border-color: #ffffff !important;
}
.nav-link, .navbar .nav-link, .list-group-item { color: #17148E; }
.nav-link:hover, .navbar .nav-link:hover, .list-group-item:hover {
    color: #ffffff !important; background-color: #17148E !important;
}
.btn-primary, .btn-primary:focus, .btn-primary:active,
.bk-btn.bk-btn-primary, .bk-btn.bk-btn-primary:focus {
    background-color: #17148E !important; border-color: #17148E !important; color: #ffffff !important;
}
.btn-primary:hover, .bk-btn.bk-btn-primary:hover {
    background-color: #17148E !important; border-color: #17148E !important; color: #ffffff !important;
}
h2, h3 { color: var(--socasp-green); font-weight: 700; }
.socasp-hero {
    background: #17148E;
    color: white; padding: 32px 36px; border-radius: 16px;
    margin-bottom: 24px; box-shadow: 0 8px 24px rgba(0,0,0,.12);
}
.socasp-hero h2 { color: white; margin: 0 0 8px 0; font-size: 30px; }
.socasp-card {
    background: #ffffff; border: 1px solid #17148E; border-radius: 14px;
    padding: 20px; box-shadow: 0 4px 16px rgba(0,0,0,.06);
}
.socasp-sidebar {
    background: #17148E; border-radius: 14px; padding: 16px;
}
.socasp-sidebar .bk-btn {
    background: #17148E; color: white; border: 1px solid #17148E;
    font-weight: 650; text-align: left; border-radius: 9px;
    min-height: 44px; font-size: 14px;
}
.socasp-sidebar .bk-btn:hover, .socasp-sidebar .bk-btn:focus {
    background: #17148E; color: white; border-color: #17148E;
}
.bk-btn.bk-btn-primary { background-color: #17148E; border-color: #17148E; color: white; }
.bk-btn.bk-btn-primary:hover { background-color: #17148E; border-color: #17148E; color: white; }
.bk-btn.bk-btn-warning, .btn-warning { background-color: #17148E !important; border-color: #17148E !important; color: white !important; }
.socasp-explanation {
    background: #ffffff; border: 1px solid #17148E; border-left: 6px solid #17148E;
    border-radius: 12px; padding: 20px; color: #17148E;
}
""")

show_dashboard = pn.widgets.Button(
    name="Tableau de bord analytique",
    button_type="primary",
    icon="chart-line",
    styles={"width": "100%"}
)
show_dashboard.on_click(lambda event: show_page("Page10"))

def _analytics_frame(dataset_name):
    """Charge les sources ou relie les volumes par année pour une comparaison annuelle."""
    fuels = ["Super", "JetA1", "Petrole", "Gazoil"]
    if dataset_name == "Comparaison annuelle Importation / Distribution":
        def annual_totals(source):
            frame = source.copy()
            frame["Annee"] = pd.to_numeric(frame["Annee"], errors="coerce")
            for fuel in fuels:
                if fuel in frame.columns:
                    frame[fuel] = pd.to_numeric(frame[fuel], errors="coerce")
            if "Marketeur" in frame.columns:
                names = frame["Marketeur"].astype(str).str.strip().str.casefold()
                total_rows = frame[names.eq("total")]
                detail_rows = frame[~names.eq("total")]
                # Use explicit annual Total records when available; otherwise
                # sum marketeur detail rows. Never add totals and details together.
                years_with_total = set(pd.to_numeric(total_rows["Annee"], errors="coerce").dropna().tolist())
                detail_rows = detail_rows[~pd.to_numeric(detail_rows["Annee"], errors="coerce").isin(years_with_total)]
                frame = pd.concat([total_rows, detail_rows], ignore_index=True)
            cols = [c for c in fuels if c in frame.columns]
            return frame.dropna(subset=["Annee"]).groupby("Annee", as_index=False)[cols].sum()

        imp = annual_totals(df_importation).rename(
            columns={fuel: f"Importation_{fuel}" for fuel in fuels}
        )
        dist = annual_totals(df_distribution).rename(
            columns={fuel: f"Distribution_{fuel}" for fuel in fuels}
        )
        joined = imp.merge(dist, on="Annee", how="inner").sort_values("Annee")
        for fuel in fuels:
            if f"Importation_{fuel}" in joined.columns:
                joined[fuel] = joined[f"Importation_{fuel}"]
        joined["Marketeur"] = "Totaux annuels (sans double comptage)"
        return joined
    data = df_importation.copy() if dataset_name == "Importations" else df_distribution.copy()
    for col in ["Annee", "Mois", "Super", "JetA1", "Petrole", "Gazoil"]:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors="coerce")
    return data

analytics_dataset = pn.widgets.RadioButtonGroup(
    name="Jeu de données",
    options=["Importations", "Distributions", "Comparaison annuelle Importation / Distribution"],
    value="Importations",
    button_type="primary"
)
analytics_chart_type = pn.widgets.Select(
    name="Type de graphique",
    options={
        "Barres verticales (Bar)": "bar",
        "Barres horizontales (Barh)": "barh",
        "Nuage de points (Scatter)": "scatter",
        "Nuage de points avec clustering": "scatter_cluster",
        "Nuage de points avec polynôme": "scatter_poly",
        "Histogramme (Hist)": "hist",
        "Courbe (Line)": "line",
        "Carte de chaleur (Heatmap)": "heatmap",
        "Aire (Area)": "area",
        "Boîte à moustaches (Boxplot)": "box",
    },
    value="bar"
)
analytics_category = pn.widgets.Select(
    name="Catégorie",
    options=["All", "Marketeur", "Annee", "Mois", "Origine", "Provenance"],
    value="All"
)
FUEL_OPTIONS = ["Super", "Petrole", "JetA1", "Gazoil"]

analytics_measure = pn.widgets.Select(
    name="Mesure / carburant",
    options=FUEL_OPTIONS,
    value="Super"
)
analytics_x_axis = pn.widgets.Select(
    name="Axe X (nuage de points)",
    options=FUEL_OPTIONS,
    value="Petrole"
)

analytics_algorithm = pn.widgets.Select(
    name="Algorithme de clustering",
    options=["K-Means", "DBSCAN"],
    value="K-Means"
)
analytics_kmeans_clusters = pn.widgets.IntSlider(
    name="Nombre de clusters K-Means", start=2, end=10, value=6, step=1
)
analytics_dbscan_eps = pn.widgets.FloatSlider(
    name="DBSCAN — eps", start=0.1, end=2.0, value=0.5, step=0.1
)
analytics_dbscan_min_samples = pn.widgets.IntSlider(
    name="DBSCAN — min_samples", start=2, end=20, value=5, step=1
)

# Keep the two fuel selections different. When one dropdown changes, remove
# that fuel from the other dropdown's available options.
def _sync_fuel_dropdowns(event=None):
    x_value = analytics_x_axis.value
    y_value = analytics_measure.value

    if x_value == y_value:
        alternatives = [fuel for fuel in FUEL_OPTIONS if fuel != x_value]
        if alternatives:
            analytics_measure.value = alternatives[0]

    analytics_x_axis.options = [
        fuel for fuel in FUEL_OPTIONS if fuel != analytics_measure.value
    ]
    analytics_measure.options = [
        fuel for fuel in FUEL_OPTIONS if fuel != analytics_x_axis.value
    ]

analytics_x_axis.param.watch(_sync_fuel_dropdowns, "value")
analytics_measure.param.watch(_sync_fuel_dropdowns, "value")
_sync_fuel_dropdowns()


GRAPH_EXPLANATIONS = {
    "pie": "### Graphique en secteurs\nMontre la part de chaque catégorie dans le total sélectionné. Utile pour comparer la contribution des marketeurs ou des origines au volume total du carburant choisi.",
    "bar": "### Barres verticales\nCompare les volumes entre catégories. Les barres les plus hautes correspondent aux catégories dont le volume cumulé est le plus important.",
    "barh": "### Barres horizontales\nCompare les volumes par catégorie dans une présentation horizontale, pratique lorsque les noms des marketeurs ou des origines sont longs.",
    "scatter": "Nuage de points\nCompare les importations et distributions des carburants sélectionnés pour repérer leur relation.",
    "hist": "### Histogramme\nPrésente la distribution des valeurs du carburant choisi. Il aide à repérer les valeurs fréquentes, la dispersion et d'éventuelles valeurs atypiques.",
    "line": "### Courbe d'évolution\nVisualise l'évolution du volume selon la catégorie choisie. Pour une lecture temporelle, sélectionnez l'année ou le mois comme catégorie.",
    "heatmap": "### Carte de chaleur\nCompare les volumes selon la catégorie et l'année à l'aide d'une échelle de couleurs. Les couleurs les plus intenses signalent les combinaisons avec les volumes les plus élevés.",
    "area": "### Graphique en aires\nMontre l'évolution du volume et met visuellement en évidence son ampleur au fil des catégories ou des périodes.",
    "box": "### Boîte à moustaches\nRésume la distribution des volumes par catégorie : médiane, dispersion et valeurs potentiellement atypiques. Ce graphique est utile pour comparer la variabilité entre groupes."
    ,
    "scatter_cluster": "Nuage de points avec clustering\nRegroupe les observations selon leurs valeurs d’importation et de distribution afin d’identifier des profils similaires.",
    "scatter_poly": "Nuage de points avec polynôme\nAjoute une courbe de tendance polynomiale de degré 2 pour illustrer la relation entre les importations et les distributions des carburants sélectionnés."
}

# ---------------------------------------------------------------------------
# KPI cards: volumes cumulés par carburant, synchronisés avec le jeu de données
# sélectionné (Importations ou Distributions).
# ---------------------------------------------------------------------------
def _format_kpi_value(value):
    return f"{value:,.2f}".replace(",", " ").replace(".", ",")


@pn.depends(dataset=analytics_dataset)
def create_fuel_kpis(dataset):
    data = _analytics_frame(dataset)
    fuels = [
        ("Super", "Super"),
        ("JetA1", "Jet A1"),
        ("Gazoil", "Gazoil"),
        ("Petrole", "Pétrole"),
    ]

    cards = []
    for column, label in fuels:
        total = (
            pd.to_numeric(data[column], errors="coerce").fillna(0).sum()
            if column in data.columns else 0
        )
        cards.append(
            pn.Column(
                pn.pane.Markdown(
                    f"<div style='font-size:13px;font-weight:600;color:#17148E'>{label}</div>",
                    sizing_mode="stretch_width",
                ),
                pn.pane.Markdown(
                    f"<div style='font-size:clamp(14px, 1.5vw, 22px);font-weight:750;color:#17148E;line-height:1.2;overflow-wrap:anywhere'>{_format_kpi_value(total)}</div>",
                    sizing_mode="stretch_width",
                ),
                pn.pane.Markdown(
                    f"<div style='font-size:11px;color:#666'>Volume total</div>",
                    sizing_mode="stretch_width",
                ),
                styles={
                    "background": "#ffffff",
                    "border": "1px solid #17148E",
                    "border-radius": "10px",
                    "padding": "clamp(6px, 1vw, 12px)",
                    "box-shadow": "0 3px 12px rgba(0,0,0,0.05)",
                    "min-width": "0",
                    "flex": "1 1 0",
                    "overflow-wrap": "anywhere",
                },
                sizing_mode="stretch_width",
                min_width=0,
            )
        )

    total_all = sum(
        pd.to_numeric(data[c], errors="coerce").fillna(0).sum()
        for c, _ in fuels if c in data.columns
    )
    cards.append(
        pn.Column(
            pn.pane.Markdown(
                "<div style='font-size:13px;font-weight:600;color:#17148E'>Total général</div>",
                sizing_mode="stretch_width",
            ),
            pn.pane.Markdown(
                f"<div style='font-size:clamp(14px, 1.5vw, 22px);font-weight:750;color:#17148E;line-height:1.2;overflow-wrap:anywhere'>{_format_kpi_value(total_all)}</div>",
                sizing_mode="stretch_width",
            ),
            pn.pane.Markdown(
                "<div style='font-size:11px;color:#666'>Tous les carburants</div>",
                sizing_mode="stretch_width",
            ),
            styles={
                "background": "#f5f5ff",
                "border": "1px solid #17148E",
                "border-radius": "10px",
                "padding": "clamp(6px, 1vw, 12px)",
                "box-shadow": "0 3px 12px rgba(0,0,0,0.05)",
                "min-width": "0",
                "flex": "1 1 0",
                "overflow-wrap": "anywhere",
            },
            sizing_mode="stretch_width",
            min_width=0,
        )
    )
    # Keep all KPI cards horizontal and distribute them across the full window.
    # Flex items shrink equally so fixed card widths cannot cause overflow.
    return pn.FlexBox(
        *cards,
        sizing_mode="stretch_width",
        flex_wrap="nowrap",
        gap="8px",
        align_items="stretch",
        styles={"width": "100%", "min-width": "0", "overflow": "hidden"},
    )


@pn.depends(
    dataset=analytics_dataset,
    chart_type=analytics_chart_type,
    category=analytics_category,
    measure=analytics_measure,
    x_axis=analytics_x_axis,
    algorithm=analytics_algorithm,
    k_clusters=analytics_kmeans_clusters,
    dbscan_eps=analytics_dbscan_eps,
    dbscan_min_samples=analytics_dbscan_min_samples,
)
def create_analytics_chart(dataset, chart_type, category, measure, x_axis):
    data = _analytics_frame(dataset)
    comparison_mode = dataset == "Comparaison annuelle Importation / Distribution"
    if data.empty:
        return pn.pane.Alert("Aucune donnée disponible pour ce jeu de données.", alert_type="warning")

    if category == "All":
        category = "Marketeur" if "Marketeur" in data.columns else "Annee"

    if measure not in data.columns:
        return pn.pane.Alert(f"La colonne « {measure} » est absente du jeu de données.", alert_type="warning")

    value_column = f"Distribution_{measure}" if comparison_mode else measure
    if value_column not in data.columns:
        return pn.pane.Alert(f"La colonne « {value_column} » est absente du jeu de données.", alert_type="warning")
    data[value_column] = pd.to_numeric(data[value_column], errors="coerce")
    data = data.dropna(subset=[value_column])
    if data.empty:
        return pn.pane.Alert(f"Aucune valeur numérique exploitable pour « {value_column} ».", alert_type="warning")

    category_exists = category in data.columns
    plot = None

    if chart_type == "hist":
        # Choose a data-adaptive number of bins using the Freedman–Diaconis rule.
        # Fall back to Sturges when the IQR is zero or the sample is very small.
        hist_values = pd.to_numeric(data[measure], errors="coerce").replace(
            [np.inf, -np.inf], np.nan
        ).dropna()
        n_values = len(hist_values)
        if n_values <= 1:
            hist_bins = 1
        else:
            q1, q3 = hist_values.quantile([0.25, 0.75])
            iqr = q3 - q1
            data_range = hist_values.max() - hist_values.min()
            if iqr > 0 and data_range > 0:
                bin_width = 2 * iqr / (n_values ** (1 / 3))
                estimated_bins = int(np.ceil(data_range / bin_width)) if bin_width > 0 else 0
            else:
                estimated_bins = int(np.ceil(np.log2(n_values) + 1))
            # Avoid unreadable histograms with too few or too many bins.
            hist_bins = max(5, min(50, estimated_bins or 10))

        plot = data.hvplot.hist(
            y=measure, bins=hist_bins, height=350, width=720, responsive=False,
            title=f"Distribution des valeurs — {measure} ({hist_bins} classes)",
            xlabel=f"Valeur de {measure}", ylabel="Fréquence", grid=True,
            color="#17148E"
        )
    elif chart_type in ("scatter", "scatter_cluster", "scatter_poly"):
        if x_axis not in data.columns:
            return pn.pane.Alert(f"La colonne « {x_axis} » est absente.", alert_type="warning")

        scatter_data = data.copy()
        x_col = f"Importation_{x_axis}" if comparison_mode else x_axis
        y_col = f"Distribution_{measure}" if comparison_mode else measure
        if x_col not in scatter_data.columns or y_col not in scatter_data.columns:
            return pn.pane.Alert("Les colonnes nécessaires à la comparaison ne sont pas disponibles.", alert_type="warning")
        scatter_data[x_col] = pd.to_numeric(scatter_data[x_col], errors="coerce")
        scatter_data[y_col] = pd.to_numeric(scatter_data[y_col], errors="coerce")
        scatter_data = scatter_data.replace([np.inf, -np.inf], np.nan).dropna(subset=[x_col, y_col])
        x_axis_plot, y_axis_plot = x_col, y_col

        if scatter_data.empty:
            return pn.pane.Alert("Aucune paire de valeurs numériques pour le nuage de points.", alert_type="warning")

        if chart_type == "scatter_cluster":
            xy = scatter_data[[x_axis_plot, y_axis_plot]].astype(float).to_numpy()
            distinct_points = np.unique(xy, axis=0)
            if len(distinct_points) < 2:
                return pn.pane.Alert("Le clustering nécessite au moins deux paires de valeurs distinctes.", alert_type="warning")
            scaler = StandardScaler()
            xy_scaled = scaler.fit_transform(xy)
            centers_data = pd.DataFrame(columns=[x_axis_plot, y_axis_plot, "Centre"])
            if analytics_algorithm.value == "DBSCAN":
                model = DBSCAN(
                    eps=float(analytics_dbscan_eps.value),
                    min_samples=int(analytics_dbscan_min_samples.value)
                )
                labels = model.fit_predict(xy_scaled)
                scatter_data["Cluster"] = [
                    "Bruit / atypique" if label == -1 else f"Cluster {label + 1}"
                    for label in labels
                ]
                n_found = len(set(labels) - {-1})
                title = f"DBSCAN — {n_found} clusters, {int((labels == -1).sum())} points atypiques"
            else:
                n_clusters = min(int(analytics_kmeans_clusters.value), len(distinct_points), len(scatter_data))
                model = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
                labels = model.fit_predict(xy_scaled)
                scatter_data["Cluster"] = [f"Cluster {label + 1}" for label in labels]
                centers = scaler.inverse_transform(model.cluster_centers_)
                centers_data = pd.DataFrame(centers, columns=[x_axis_plot, y_axis_plot])
                centers_data["Centre"] = "Centre K-Means"
                title = f"K-Means — {n_clusters} clusters"
            plot = scatter_data.hvplot.scatter(
                x=x_axis_plot, y=y_axis_plot, by="Cluster",
                height=350, width=720, responsive=False, title=title,
                xlabel=f"Importation — {x_axis}" if comparison_mode else (f"Importation / {x_axis}" if dataset == "Importations" else f"Distribution / {x_axis}"),
                ylabel=f"Distribution — {measure}" if comparison_mode else f"Volume — {measure}",
                grid=True, size=75, alpha=0.8, legend="right"
            )
            if not centers_data.empty:
                plot = plot * centers_data.hvplot.scatter(
                    x=x_axis_plot, y=y_axis_plot, by="Centre",
                    height=350, width=720, responsive=False,
                    size=180, marker="x", line_width=3, color="#00FF00", legend="right"
                )
        elif chart_type == "scatter_poly":
            plot = scatter_data.hvplot.scatter(
                x=x_axis_plot, y=y_axis_plot, height=350, width=720, responsive=False,
                title=f"Tendance polynomiale — {y_col} en fonction de {x_col}",
                xlabel=f"Importation — {x_axis}" if comparison_mode else x_axis,
                ylabel=f"Distribution — {measure}" if comparison_mode else measure,
                grid=True, size=75, alpha=0.75, color="#17148E", legend=False
            )
            if scatter_data[x_axis_plot].nunique() < 3:
                return pn.pane.Alert("La courbe polynomiale de degré 2 nécessite au moins trois valeurs distinctes sur l'axe X.", alert_type="warning")
            ordered = scatter_data.sort_values(x_axis_plot)
            coefficients = np.polyfit(ordered[x_axis_plot].astype(float), ordered[y_axis_plot].astype(float), deg=2)
            x_curve = np.linspace(ordered[x_axis_plot].min(), ordered[x_axis_plot].max(), 150)
            curve_data = pd.DataFrame({x_axis_plot: x_curve, y_axis_plot: np.polyval(coefficients, x_curve)})
            curve = curve_data.hvplot.line(
                x=x_axis_plot, y=y_axis_plot, height=350, width=720, responsive=False,
                xlabel=x_axis_plot, ylabel=y_axis_plot, grid=True,
                color="#d62728", line_width=3, legend=False
            )
            plot = plot * curve
        else:
            plot = scatter_data.hvplot.scatter(
                x=x_axis_plot, y=y_axis_plot, height=350, width=720, responsive=False,
                title=f"{y_col} en fonction de {x_col}",
                xlabel=f"Importation — {x_axis}" if comparison_mode else x_axis,
                ylabel=f"Distribution — {measure}" if comparison_mode else measure,
                grid=True, size=75, alpha=0.75, color="#17148E", legend=False
            )
    elif chart_type == "heatmap":
        if "Annee" not in data.columns or "Mois" not in data.columns:
            return pn.pane.Alert("La carte de chaleur nécessite les colonnes Année et Mois.", alert_type="warning")
        heat_x = category if category_exists and category != "Annee" else "Mois"
        heat_data = data.dropna(subset=[heat_x, "Annee", measure])
        heat_data = heat_data.groupby([heat_x, "Annee"], as_index=False)[measure].sum()
        if heat_data.empty:
            return pn.pane.Alert("Aucune donnée disponible pour construire la carte de chaleur.", alert_type="warning")
        plot = heat_data.hvplot.heatmap(
            x=heat_x, y="Annee", C=measure, aggregate="sum",
            height=350, width=720, responsive=False,
            title=f"Carte de chaleur — {measure} par {heat_x} et année",
            xlabel=heat_x, ylabel="Année", colorbar=True,
            cmap=["#17148E", "#17148E"]
        )
    else:
        if not category_exists:
            return pn.pane.Alert(
                f"La catégorie « {category} » n'est pas disponible dans {dataset.lower()}. Choisissez une autre catégorie.",
                alert_type="warning"
            )
        grouped = (
            data.dropna(subset=[category])
                .groupby(category, dropna=False)[measure]
                .sum()
                .reset_index()
                .sort_values(measure, ascending=False)
        )
        if grouped.empty:
            return pn.pane.Alert("Aucune donnée à afficher avec les paramètres sélectionnés.", alert_type="warning")

        common = dict(height=350, width=720, responsive=False, title=f"{measure} par {category}", grid=True)
        if chart_type == "pie":
            plot = grouped.hvplot.pie(
                y=measure, by=category, title=f"Répartition de {measure} par {category}",
                height=350, width=720, responsive=False, legend="right"
            )
        elif chart_type == "bar":
            plot = grouped.hvplot.bar(
                x=category, y=measure, **common, xlabel=category,
                ylabel=f"Volume de {measure}", rot=35, color="#17148E"
            )
        elif chart_type == "barh":
            plot = grouped.hvplot.barh(
                x=category, y=measure, **common, xlabel=f"Volume de {measure}",
                ylabel=category, color="#17148E"
            )
        elif chart_type == "line":
            line_data = grouped.sort_values(category)
            plot = line_data.hvplot.line(
                x=category, y=measure, **common, xlabel=category,
                ylabel=f"Volume de {measure}", marker="circle", color="#17148E"
            )
        elif chart_type == "area":
            area_data = grouped.sort_values(category)
            plot = area_data.hvplot.area(
                x=category, y=measure, **common, xlabel=category,
                ylabel=f"Volume de {measure}", alpha=0.65, color="#17148E"
            )
        elif chart_type == "box":
            plot = data.hvplot.box(
                y=measure, by=category, height=350, width=720, responsive=False,
                title=f"Dispersion de {measure} par {category}",
                xlabel=category, ylabel=f"Volume de {measure}", grid=True,
                color="#17148E"
            )

    if plot is None:
        return pn.pane.Alert("Type de graphique non reconnu.", alert_type="warning")

    explanation_text = GRAPH_EXPLANATIONS.get(
        chart_type, "### Lecture du graphique\\nExplorez les données à l'aide des paramètres."
    )
    if chart_type in ("scatter", "scatter_cluster", "scatter_poly"):
        corr_x = f"Importation_{x_axis}" if comparison_mode else x_axis
        corr_y = f"Distribution_{measure}" if comparison_mode else measure
        corr_frame = data[[corr_x, corr_y]].apply(pd.to_numeric, errors="coerce").replace(
            [np.inf, -np.inf], np.nan
        ).dropna()
        if len(corr_frame) >= 2 and corr_frame[corr_x].nunique() > 1 and corr_frame[corr_y].nunique() > 1:
            r_value = corr_frame[corr_x].corr(corr_frame[corr_y])
            strength = "forte" if abs(r_value) >= 0.7 else "modérée" if abs(r_value) >= 0.4 else "faible"
            direction = "positive" if r_value > 0 else "négative" if r_value < 0 else "nulle"
            explanation_text += (
                f"\\n\\n**Corrélation de Pearson :** r = {r_value:.3f} ({direction}, intensité {strength}), "
                f"calculée sur {len(corr_frame)} paire(s). La corrélation décrit une association, pas une causalité."
            )
        else:
            explanation_text += (
                "\\n\\n**Corrélation non calculable :** il faut au moins deux paires valides et une variation sur les deux axes."
            )
        if comparison_mode and len(corr_frame) < 10:
            explanation_text += (
                f"\\n\\n**Prudence :** seulement {len(corr_frame)} année(s) commune(s) sont disponibles. "
                "Le résultat est exploratoire et ne doit pas être généralisé."
            )
        if chart_type == "scatter_cluster":
            if analytics_algorithm.value == "K-Means":
                explanation_text += (
                    f"\\n\\n**K-Means :** répartit les observations en {int(analytics_kmeans_clusters.value)} groupes demandés "
                    "en minimisant la distance aux centres. Les croix vertes représentent les centres des groupes. "
                    "Les variables sont standardisées avant le calcul pour limiter l'effet des échelles différentes."
                )
            else:
                explanation_text += (
                    f"\\n\\n**DBSCAN :** détecte les groupes selon la densité, sans fixer à l'avance leur nombre. "
                    f"Paramètres : eps = {analytics_dbscan_eps.value:.1f}, min_samples = {int(analytics_dbscan_min_samples.value)}. "
                    "Les points « Bruit / atypique » ne sont rattachés à aucun groupe dense."
                )
    explanation = pn.pane.Markdown(explanation_text, sizing_mode="stretch_width")
    explanation_card = pn.Column(
        explanation,
        pn.pane.Markdown(
            f"**Jeu de données :** {dataset}  \n"
            f"**Indicateur :** {measure}  \n"
            f"**Catégorie :** {category}",
            styles={"color": "#17148E", "font-size": "13px"}
        ),
        styles={"background": "#ffffff", "border": "1px solid #17148E", "border-left": "5px solid #17148E",
                "border-radius": "10px", "padding": "16px"},
        width=435,
        sizing_mode="fixed"
    )
    return pn.Row(
        pn.Column(plot, sizing_mode="fixed", min_width=600, width=720),
        explanation_card,
        sizing_mode="stretch_width",
        min_height=440,
        margin=(8, 0, 12, 0)
    )

def CreatePage10():
    hero = pn.pane.HTML(
        '<div class="socasp-hero">'
        '<h2>Tableau de bord analytique SOCASP</h2>'
        '<div>Analyse interactive des importations et distributions de carburants.</div>'
        '</div>',
        sizing_mode="stretch_width"
    )
    controls = pn.Card(
        pn.Row(analytics_dataset, analytics_chart_type, sizing_mode="stretch_width"),
        pn.Row(analytics_category, analytics_measure, analytics_x_axis, sizing_mode="stretch_width"),
        pn.Row(analytics_algorithm, analytics_kmeans_clusters, sizing_mode="stretch_width"),
        pn.Row(analytics_dbscan_eps, analytics_dbscan_min_samples, sizing_mode="stretch_width"),
        title="Paramètres d'analyse",
        collapsed=False,
        styles={"background": "white", "border-radius": "12px", "border": "1px solid #17148E"}
    )
    note = pn.pane.Markdown(
        "**Mode d'emploi :** choisissez le jeu de données, le type de graphique, la catégorie et le carburant. "
        "Le graphique et son explication en français se mettent à jour automatiquement. En comparaison annuelle, les lignes « Total » sont exclues pour éviter le double comptage.",
        styles={"color": "#17148E"}
    )
    kpi_heading = pn.pane.Markdown("### Indicateurs clés de performance (KPI)")
    return pn.Column(
        hero,
        kpi_heading,
        create_fuel_kpis,
        controls,
        create_analytics_chart,
        note,
        sizing_mode="stretch_width",
        min_width=1000,
        margin=(0, 0, 24, 0),
    )

# ---------------------------------------------------------------------------

mapping = { 
    "Page1": CreatePage1(),
    "Page2": CreatePage2(),
    "Page3": CreatePage3(),
    "Page4": CreatePage4(),
    "Page5": CreatePage5(),
    "Page6": CreatePage6(),
    "Page7": CreatePage7(),
    "Page8": CreatePage8(),
    "Page9": CreatePage9(),
    "Page10": CreatePage10(),
}

create_user_menu = pn.widgets.Button(
    name="Créer Utilisateur", button_type="primary", icon="user-plus",
    styles={"width": "100%"}
)
add_information_menu = pn.widgets.Button(
    name="Ajout Informations", button_type="primary", icon="database",
    styles={"width": "100%"}
)
create_user_menu.on_click(lambda event: show_page("Page8"))
add_information_menu.on_click(lambda event: show_page("Page9"))

logout = pn.widgets.Button(
    name="Déconnexion",
    button_type="primary",
    icon="clipboard-data",
    styles={"width": "100%"},
)

def _logout_to_login(event=None):
    global CURRENT_USERNAME
    # This is an application-level logout: return to the login form.
    # It does not clear Panel's HTTP Basic Auth session because this app uses its own form.
    CURRENT_USERNAME = ""
    sidebar.visible = False
    login_username.value = ""
    login_password.value = ""
    login_message.object = ""
    main_area.clear()
    main_area.append(login_screen)
    template.title = "SOCASP — Connexion"

logout.on_click(_logout_to_login)

# Top navigation tabs in the main body. The sidebar keeps only the dashboard shortcut and logout.
main_tabs = pn.Tabs(
    ("Importation par Marketeur", mapping["Page1"]),
    ("Importation Annuelle", mapping["Page2"]),
    ("Distribution par Marketeur", mapping["Page3"]),
    ("Distribution Annuelle", mapping["Page4"]),
    ("Ajout Importation", mapping["Page5"]),
    ("Ajout Distribution", mapping["Page6"]),
    ("Droits Utilisateurs", mapping["Page7"]),
    sizing_mode="stretch_width",
    tabs_location="above",
    styles={"width": "100%"},
)

def open_dashboard_tab(event=None):
    # The dashboard remains accessible from the sidebar, not in the top tab bar.
    main_area.clear()
    main_area.append(create_welcome_header())
    main_area.append(mapping["Page10"])

show_dashboard.on_click(open_dashboard_tab)

sidebar = pn.Column(
    pn.pane.Markdown("## Menu", styles={"color": "#17148E"}),
    show_dashboard,
    create_user_menu,
    add_information_menu,
    logout,
    styles={"width": "100%", "padding": "18px", "background": "#ffffff", "border-radius": "14px", "border": "1px solid #eeeeee"},
)
# Custom SOCASP login page inspired by the supplied screenshot.
# Launch WITHOUT --basic-auth so this in-app login page is shown first.
logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
credentials_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials.json")

login_username = pn.widgets.TextInput(
    name="Nom d'utilisateur", placeholder="Nom d'utilisateur",
    sizing_mode="stretch_width", height=42,
)
login_password = pn.widgets.PasswordInput(
    name="Mot de passe", placeholder="Mot de passe",
    sizing_mode="stretch_width", height=42,
)
login_message = pn.pane.Markdown("", styles={"color": "#b91c1c", "font-size": "13px"})
login_button = pn.widgets.Button(
    name="LOGIN", button_type="primary", sizing_mode="stretch_width", height=42
)

def _load_login_credentials():
    try:
        with open(credentials_path, "r", encoding="utf-8") as credential_file:
            credentials = json.load(credential_file)
        if not isinstance(credentials, dict) or not credentials:
            raise ValueError("credentials.json doit contenir un objet JSON username: password.")
        return {str(username): str(password) for username, password in credentials.items()}
    except Exception as exc:
        login_message.object = (
            "**Configuration de connexion invalide.** Vérifiez credentials.json. "
            + str(exc).replace("<", "&lt;").replace(">", "&gt;")
        )
        return {}

def _show_login(event=None):
    main_area.clear()
    main_area.append(login_screen)
    sidebar.visible = False

def _login(event=None):
    global CURRENT_USERNAME
    credentials = _load_login_credentials()
    username = login_username.value.strip()
    password = login_password.value
    if username in credentials and password and credentials[username] == password:
        login_message.object = ""
        CURRENT_USERNAME = username
        login_username.value = ""
        login_password.value = ""
        sidebar.visible = True
        main_area.clear()
        main_area.append(create_welcome_header())
        main_area.append(main_tabs)
        template.title = f"SOCASP — Bienvenue, {username}"
    else:
        login_message.object = "Identifiant ou mot de passe incorrect."

login_button.on_click(_login)
login_password.param.watch(lambda event: setattr(login_message, "object", ""), "value")
login_username.param.watch(lambda event: setattr(login_message, "object", ""), "value")

login_logo = (
    pn.pane.PNG(logo_path, width=150, height=100, sizing_mode="fixed")
    if os.path.exists(logo_path)
    else pn.pane.Markdown(
        "<div style='font-size:30px;font-weight:800;color:#17148E;text-align:center'>SOCASP</div>",
        width=220, height=75, sizing_mode="fixed",
    )
)
login_screen = pn.Column(
    pn.Spacer(height=20),
    pn.Row(login_logo, align="center"),
    pn.pane.Markdown(
        "<div style='text-align:center;font-size:16px;color:#111827'>"
        "Login to access your application</div>",
        sizing_mode="stretch_width",
    ),
    login_username,
    login_password,
    login_message,
    login_button,
    width=366,
    max_width=366,
    styles={
        "background": "#ffffff",
        "border": "1px solid #dddddd",
        "padding": "38px 32px 48px 32px",
        "position": "fixed",
        "top": "50%",
        "left": "50%",
        "transform": "translate(-50%, -50%)",
        "margin": "0",
        "box-sizing": "border-box",
        "z-index": "1000",
        "max-height": "95vh",
        "overflow-y": "auto",
    },
    sizing_mode="fixed",
)
main_area = pn.Column(
    login_screen,
    styles={
        "width": "100%",
        "min-height": "100vh",
        "background": "#fafafa",
        "padding": "0",
        "box-sizing": "border-box",
    },
    sizing_mode="stretch_width",
)
sidebar.visible = False

template = pn.template.BootstrapTemplate(
    title="SOCASP — Connexion",
    sidebar=[sidebar],
    main=[main_area],
    # footer=[sidebar],
    header_background="#ffffff",
    site="Société Centrafricaine de stockage des Produits Pétroliers", logo=(logo_path if os.path.exists(logo_path) else None),
    sidebar_width=310,
    busy_indicator=None,
)

template.servable()