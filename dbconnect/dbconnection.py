__author__ = 'i20764'
#import mysql.connector
import MySQLdb
import config.config

def dbconnect():
    try:
        db = MySQLdb.connect(database=config.config.db_name,host=config.config.server,
                                 user=config.config.user_name,password=config.config.password)
        print('Connection Established')
    except ConnectionError :
        print('Could not connect to database')
    return db