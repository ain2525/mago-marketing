import streamlit as st
import pandas as pd

# ページ設定
st.set_page_config(page_title="まごころサポート分析", layout="wide")
st.title("📊 まごころサポート：広告×商談 分析アプリ")

# データの読み込み関数
def load_data(file):
    try:
        if file.name.endswith('.csv'):
            try:
                return pd.read_csv(file)
            except:
                return pd.read_csv(file, encoding='shift-jis')
    except:
        return None
    return pd.read_excel(file)

# サイドバー
st.sidebar.header("⚙️ 設定")
cpa_limit = st.sidebar.number_input("許容CPA（円）", value=15000, step=1000)
meeting_target = st.sidebar.slider("目標商談化率（%）", 0, 30, 10)

# アップロード
col1, col2 = st.columns(2)
with col1:
    meta_file = st.file_uploader("📂 Meta広告実績", type=['xlsx', 'csv'])
with col2:
    hs_file = st.file_uploader("📂 HubSpotデータ", type=['xlsx', 'csv'])

st.divider()

# 分析実行
if meta_file and hs_file:
    df_meta = load_data(meta_file)
    df_hs = load_data(hs_file)

    if df_meta is not None and df_hs is not None:
        # 列名のゆらぎ吸収
        meta_cols = df_meta.columns.astype(str)
        name_col = next((c for c in meta_cols if '名前' in c or 'Name' in c), None)
        spend_col = next((c for c in meta_cols if '消化金額' in c or 'Amount' in c), None)
        res_col = next((c for c in meta_cols if '結果' in c or 'Results' in c), None)

        hs_cols = df_hs.columns.astype(str)
        utm_col = next((c for c in hs_cols if 'UTM' in c or 'Content' in c), None)
        deal_col = next((c for c in hs_cols if '商談' in c or 'Deal' in c), None)

        if name_col and spend_col and res_col and utm_col:
            # キー作成
            df_meta['key'] = df_meta[name_col].astype(str).str.extract(r'(bn\d+)')[0]
            df_hs['key'] = df_hs[utm_col].astype(str).str.strip()
            
            # 集計
            meta_agg = df_meta.groupby('key')[[spend_col, res_col]].sum().reset_index()
            
            # 商談カウント
            if deal_col:
                hs_deals = df_hs[df_hs[deal_col].fillna('').astype(str).str.contains('あり|TRUE|Yes', case=False)]
                deal_counts = hs_deals.groupby('key').size().reset_index(name='商談数')
            else:
                deal_counts = pd.DataFrame({'key': [], '商談数': []})

            # 結合
            result = pd.merge(meta_agg, deal_counts, on='key', how='left').fillna(0)
            
            # 計算
            result['CPA'] = (result[spend_col] / result[res_col]).replace([float('inf')], 0).astype(int)
            result['商談化率'] = (result['商談数'] / result[res_col] * 100).round(1)

            # 判定
            def judge(row):
                if row['商談化率'] >= meeting_target and row['CPA'] <= cpa_limit: return "🏆 勝ち"
                if row['商談化率'] >= meeting_target: return "🟡 質良"
                if row['CPA'] <= cpa_limit: return "🥈 CPA良"
                return "🛑 停止"
            
            result['判定'] = result.apply(judge, axis=1)
            
            # 表示
            st.subheader("分析結果")
            st.dataframe(result.style.applymap(lambda x: 'background-color: #d4edda' if '勝ち' in str(x) else '', subset=['判定']))
            
        else:
            st.error("必要な列が見つかりませんでした。列名を確認してください。")
    else:
        st.error("ファイルの読み込みに失敗しました。")
