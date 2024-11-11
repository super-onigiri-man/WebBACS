import streamlit as st
import pandas as pd

# セッションステートを使用してデータを永続化
if 'data' not in st.session_state:
    st.session_state['data'] = {
        '名前': ['Alice', 'Bob', 'Charlie', 'David'],
        '年齢': [24, 30, 22, 35],
        '都市': ['New York', 'Los Angeles', 'Chicago', 'Houston']
    }

df = pd.DataFrame(st.session_state['data'])

st.title('データフレーム編集アプリ')

# データフレームの表示
st.write('現在のデータフレーム')
st.dataframe(df)

# 編集フォームの作成
st.subheader('データを編集')

name = st.selectbox('名前を選択', df['名前'])
age = st.number_input('年齢を入力', min_value=0, max_value=120, value=int(df[df['名前'] == name]['年齢']))
city = st.text_input('都市を入力', df[df['名前'] == name]['都市'].values[0])

if st.button('更新'):
    # データを更新
    idx = df.index[df['名前'] == name][0]
    st.session_state['data']['年齢'][idx] = age
    st.session_state['data']['都市'][idx] = city
    
    # データ更新後にデータフレームを再生成
    df = pd.DataFrame(st.session_state['data'])
    raise st.rerun()