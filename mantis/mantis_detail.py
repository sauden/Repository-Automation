__author__ = 'i20764'
import os,shutil,fileinput
import config.config
import dbconnect.dbconnection
from openpyxl import load_workbook
import pandas as pd
import csv
import glob
import re


db = dbconnect.dbconnection.dbconnect()

def getMantisDetail():
    cur = db.cursor()
    getDet = cur.execute('SELECT a.id,(select note from mantis.mantis_bugnote_text_table where id in (select max(id) from mantis.mantis_bugnote_table where bug_id = a.id )) as mantis_note from mantis.mantis_bug_table a where a.handler_id in ( ''204 '', ''330 '', ''366 '', ''374 '', ''402 '') and a.status= ''50 '' and a.project_id =  ''8''')
    getNote = cur.fetchall()
    cur.close()
    return getNote

def createdir(mantisid):
    try:
        if not os._exists(mantisid):
            os.mkdir(config.config.path+mantisid)
            print(config.config.path+mantisid)
            print('directory: '+mantisid+' created')
            wd = config.config.path+mantisid
    except:
        print('Folder already exists')
        wd=1

    return wd

def get_dir():
    dir_list = os.listdir(config.config.template_path)
    return dir_list
















