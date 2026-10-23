import panel as pn
import pandas as pd
import hvplot.pandas
import warnings
import mysql.connector
import json
import datetime as dt
import os


def returnLogDetails():

    host_socasp = "192.168.1.143"
    user_socasp = 'dondonedmond'
    password_socasp = 'toor'
    database_socasp = 'socasp'

    database_user = 'mysql'

    return host_socasp, user_socasp, password_socasp, database_socasp, database_user