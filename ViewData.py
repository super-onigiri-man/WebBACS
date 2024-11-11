import streamlit as st
import pandas as pd
import sqlite3

# データベース接続
dbname = 'test.db'
conn = sqlite3.connect(dbname, isolation_level=None)  # 自動コミットモードで接続
cursor = conn.cursor()  # カーソルオブジェクトを作成


# データがない場合、サンプルデータを挿入
cursor.execute('SELECT COUNT(*) FROM music_master')
if cursor.fetchone()[0] == 0:
    sample_data = [
        ("Song A", "Artist A", 95),
        ("Song B", "Artist B", 90),
        ("Song C", "Artist C", 85),
        ("Song D", "Artist D", 80),
    ]
    cursor.executemany('INSERT INTO music_master (Title, Artist, Score) VALUES (?, ?, ?)', sample_data)
    conn.commit()

# データの取得
df = pd.read_sql_query('''
SELECT * FROM music_master
WHERE Score IS NOT NULL AND Score != ''
ORDER BY Score DESC
LIMIT 60
''', conn)

# Streamlitによるデータ編集
st.title('音楽ランキング編集アプリ')

# データフレームの編集
edited_df = st.data_editor(df)

# 編集を確認するためのボタン
if st.button('データを確認'):
    st.write('編集後のデータ:')
    st.dataframe(edited_df)
    
    # データベースを更新
    for index, row in edited_df.iterrows():
        cursor.execute('''
        UPDATE music_master
        SET Title = ?, Artist = ?, Score = ?
        WHERE Unique_id = ?
        ''', (row['Title'], row['Artist'], row['Score'], row['Unique_id']))
    conn.commit()
    
    # データのリロード
    st.rerun()

# 接続を閉じる
conn.close()