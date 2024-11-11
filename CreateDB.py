import csv
import sqlite3
import PySimpleGUI as sg
import traceback
import sys
import os

dbname = 'test.db'  # データベース名
csv_file = '楽曲データ.csv'  # CSVファイル名

conn = sqlite3.connect('test.db', isolation_level=None, check_same_thread=False) # データベースに接続（自動コミット機能ON）
cursor = conn.cursor()  # カーソルオブジェクトを作成


try: 

    # テーブル作成のSQL文(基盤)
    sql = """CREATE TABLE music_master (
        Title TEXT,
        Artist TEXT,
        Score DOUBLE,
        Last_Rank INT,
        Last_Number INT,
        On_Chart INT,
        Unique_id TEXT,
        PRIMARY KEY (Unique_id)
    );"""
    
    # DBのフォーマット設定
    cursor.execute(sql)
    conn.commit()

    with open('楽曲データ.csv', encoding='utf-8') as f:
        next(f)
        csv_reader = csv.reader(f)
        for row in csv_reader:
            cursor.execute("INSERT OR REPLACE INTO music_master VALUES (?, ?, ?, ?, ?, ?, ?)", (row[0], row[1], row[2], row[3], row[4], row[5], row[6]))

    cursor.execute('UPDATE music_master SET Score = 0;') #Scoreの値を全消し
    cursor.execute('''DELETE FROM music_master WHERE Last_Number = '' OR Last_Number = 0;''') #Last_Numberの値が0もしくは空文字の場合消す
    cursor.execute('''DELETE FROM music_master WHERE Last_Number = "ベストヒットえひめ";''')

    conn.close()

except Exception as e:
    with open('error.log', 'a') as f:
        traceback.print_exc(file=f)
