import streamlit as st
import sqlite3
import pandas as pd
import GetData
import CreateDB

# データベース接続設定
dbname = 'test.db'
conn = sqlite3.connect(dbname, isolation_level=None)  # 自動コミットモードで接続
cursor = conn.cursor()  # カーソルオブジェクトを作成

# セッション状態の設定
if 'database_initialized' not in st.session_state:
    st.session_state['database_initialized'] = False

# st.session_state['df'] は初期化せず、generate_data() 内で定義する

HaruyaPath = st.file_uploader('明屋書店のデータを選択してください', type='xlsx')

# st.successを表示するプレースホルダーを作成
success_message = st.empty()

# リロード時に実行する関数
def reload_process():
    st.info("ページがリロードされました。")
    
    st.session_state['df'] = pd.read_sql('''
        SELECT * FROM music_master
        WHERE Score IS NOT NULL AND Score != ''
        ORDER BY Score DESC
        LIMIT 60
    ''', conn)
    

def initialize_database():
    # データベースを初期化する
    # 例: CreateDBの利用
    # CreateDB.initialize()  # これが必要であれば実装
    st.session_state['database_initialized'] = True

def generate_data():
    if HaruyaPath is None:
        st.success("明屋書店のデータなしでランキングを作成します")
    else:
        GetData.WebGetThisWeekRank(HaruyaPath)

    # データをデータフレームとして取得
    df = pd.read_sql('''
    SELECT * FROM music_master
    WHERE Score IS NOT NULL AND Score != ''
    ORDER BY Score DESC
    LIMIT 60''', conn)
    return df

def update_database(edited_df):
    # データベースを更新
    for index, row in edited_df.iterrows():
        cursor.execute('''
        UPDATE music_master
        SET Title = ?, Artist = ?, Score = ?
        WHERE Unique_id = ?
        ''', (row['Title'], row['Artist'], row['Score'], row['Unique_id']))
    conn.commit()

if st.button('データ生成'):
    if not st.session_state['database_initialized']:
        initialize_database()
    st.session_state['df'] = generate_data()

if st.session_state['df'] is not None:
    # 編集可能なデータエディタを表示
    edited_df = st.data_editor(st.session_state['df'].copy().head(20), num_rows="dynamic")  # st.session_state['df'] を直接編集しないようにコピーを渡す

    # 編集を反映
    if edited_df is not None:
        update_database(edited_df)
        st.session_state['df'] = edited_df  # 編集後のデータフレームをst.session_state['df']に保存
        reload_process()
        success_message.success("データが更新されました！")

    if st.button('リロードする'):
        update_database(edited_df)
        st.session_state['df'] = edited_df
        success_message.success("データが更新されました！")