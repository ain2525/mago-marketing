import streamlit as st
import pandas as pd
import altair as alt

# --- ページ設定 ---
st.set_page_config(page_title="まごころ分析 v3.1", layout="wide")
st.title("📊 まごころサポート：広告×商談 分析ダッシュボード")

# --- エラーハンドリング付きデータ読み込み ---
def load_data(file):
    try:
        if file.name.endswith('.csv'):
            try:
                return pd.read_csv(file)
            except:
                return pd.read_csv(file, encoding='shift-jis')
    except Exception as e:
        st.error(f"ファイル読み込みエラー: {e}")
        return None
    return pd.read_excel(file)

# --- サイドバー ---
st.sidebar.header("⚙️ 判定基準")
cpa_limit = st.sidebar.number_input("許容CPA（円）", value=15000, step=1000)
meeting_target = st.sidebar.slider("目標商談化率（%）", 0, 30, 10)

# --- アップロード ---
col1, col2 = st.columns(2)
with col1:
    meta_file = st.file_uploader("📂 Meta広告実績", type=['xlsx', 'csv'])
with col2:
    hs_file = st.file_uploader("📂 HubSpotデータ", type=['xlsx', 'csv'])

st.divider()

# --- メイン処理 ---
if meta_file and hs_file:
    # 1. データ読み込み
    df_meta = load_data(meta_file)
    df_hs = load_data(hs_file)

    if df_meta is not None and df_hs is not None:
        try:
            # 2. 列名の自動検出（安全策）
            meta_cols = df_meta.columns.astype(str)
            name_col = next((c for c in meta_cols if '名前' in c or 'Name' in c), None)
            spend_col = next((c for c in meta_cols if '消化金額' in c or 'Amount' in c), None)
            res_col = next((c for c in meta_cols if '結果' in c or 'Results' in c), None)

            hs_cols = df_hs.columns.astype(str)
            utm_col = next((c for c in hs_cols if 'UTM' in c or 'Content' in c), None)
            
            # 詳細分析用の列（なければ無視する安全設計）
            connect_col = next((c for c in hs_cols if '接続' in c), None)
            deal_col = next((c for c in hs_cols if '商談' in c and '予定' not in c), None)
            deal_plan_col = next((c for c in hs_cols if '商談' in c and '予定' in c), None)
            attr_col = next((c for c in hs_cols if '属性' in c), None)
            stage_col = next((c for c in hs_cols if 'ステージ' in c), None)

            # 必須カラムのチェック
            if not (name_col and spend_col and res_col and utm_col):
                st.error(f"必要な列が見つかりません。\n検出された列: Meta={name_col}, {spend_col}, {res_col} / HS={utm_col}")
                st.stop()

            # 3. データの結合キー作成
            # regexで抽出できなかった場合(NaN)は、そのままの値を使うように変更（エラー回避）
            df_meta['key'] = df_meta[name_col].astype(str).str.extract(r'(bn\d+)')[0]
            df_meta.loc[df_meta['key'].isna(), 'key'] = df_meta[name_col].astype(str) # 救済措置

            df_hs['key'] = df_hs[utm_col].astype(str).str.strip()

            # 4. 集計処理
            meta_agg = df_meta.groupby('key')[[spend_col, res_col]].sum().reset_index()

            if deal_col:
                # 'あり' 'True' 'Yes' などが含まれる行をカウント
                hs_deals = df_hs[df_hs[deal_col].fillna('').astype(str).str.contains('あり|TRUE|Yes', case=False)]
                deal_counts = hs_deals.groupby('key').size().reset_index(name='商談数')
            else:
                deal_counts = pd.DataFrame({'key': [], '商談数': []})

            # 結合
            result = pd.merge(meta_agg, deal_counts, on='key', how='left').fillna(0)

            # 指標計算（ゼロ除算エラー回避）
            result['CPA'] = result.apply(lambda x: int(x[spend_col] / x[res_col]) if x[res_col] > 0 else 0, axis=1)
            result['商談化率'] = result.apply(lambda x: x['商談数'] / x[res_col] if x[res_col] > 0 else 0, axis=1)

            # 判定ロジック
            def judge(row):
                rate_pct = row['商談化率'] * 100
                if rate_pct >= meeting_target and row['CPA'] <= cpa_limit and row['商談数'] > 0: return "🏆 勝ち"
                if rate_pct >= meeting_target and row['商談数'] > 0: return "🟡 質良"
                if row['CPA'] <= cpa_limit: return "🥈 CPA良"
                return "🛑 停止"
            
            result['判定'] = result.apply(judge, axis=1)

            # --- 表示パート ---

            # サマリー
            total_spend = result[spend_col].sum()
            total_leads = result[res_col].sum()
            total_meetings = result['商談数'].sum()
            avg_cpa = int(total_spend / total_leads) if total_leads > 0 else 0
            avg_rate = (total_meetings / total_leads * 100) if total_leads > 0 else 0

            st.subheader("📈 実績サマリー")
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("総消化金額", f"¥{total_spend:,.0f}")
            k2.metric("総リード数", f"{int(total_leads)}件")
            k3.metric("総商談数", f"{int(total_meetings)}件", delta=f"{avg_rate:.1f}%")
            k4.metric("平均CPA", f"¥{avg_cpa:,.0f}")

            tab1, tab2 = st.tabs(["📊 バナー別成績", "📋 リード詳細リスト"])

            with tab1:
                # エラー回避: データがある場合のみグラフ描画
                if not result.empty and result[res_col].sum() > 0:
                    chart_data = result[result[res_col] > 0].copy()
                    chart_data['商談化率(%)'] = chart_data['商談化率'] * 100
                    
                    # Altairグラフ
                    base = alt.Chart(chart_data).encode(
                        x=alt.X('CPA', title='CPA (円)'),
                        y=alt.Y('商談化率(%)', title='商談化率 (%)'),
                        tooltip=['key', 'CPA', '商談化率(%)', '消化金額']
                    )
                    points = base.mark_circle(size=100).encode(color='判定')
                    st.altair_chart(points.interactive(), use_container_width=True)
                else:
                    st.info("グラフを表示するデータがありません")

                # テーブル表示
                display_df = result.copy()
                display_df = display_df.rename(columns={'key': 'バナーID', spend_col: '消化金額', res_col: 'リード数'})
                
                st.dataframe(
                    display_df[['判定', 'バナーID', '消化金額', 'リード数', 'CPA', '商談数', '商談化率']].sort_values('消化金額', ascending=False),
                    column_config={
                        "消化金額": st.column_config.NumberColumn(format="¥%d"),
                        "CPA": st.column_config.NumberColumn(format="¥%d"),
                        "商談化率": st.column_config.NumberColumn(format="%.1f%%"),
                    },
                    use_container_width=True,
                    hide_index=True
                )

            with tab2:
                st.subheader("リード詳細分析")
                # 詳細データの結合
                detail_df = df_hs.copy()
                detail_df = pd.merge(detail_df, result[['key', '判定']], on='key', how='left')
                
                # 表示項目の整理
                cols_map = {'key': '流入バナー', '判定': '判定'}
                if connect_col: cols_map[connect_col] = '接続'
                if deal_col: cols_map[deal_col] = '商談有無'
                if deal_plan_col: cols_map[deal_plan_col] = '商談予定'
                if attr_col: cols_map[attr_col] = '属性'
                if stage_col: cols_map[stage_col] = 'ステージ'

                # 存在する列だけ抽出
                available_cols = [c for c in cols_map.keys() if c in detail_df.columns]
                
                # バナー絞り込み
                banner_filter = st.multiselect("バナーで絞り込み", options=detail_df['key'].unique())
                if banner_filter:
                    detail_df = detail_df[detail_df['key'].isin(banner_filter)]

                st.dataframe(
                    detail_df[available_cols].rename(columns=cols_map).fillna('-'),
                    use_container_width=True,
                    hide_index=True
                )

        except Exception as e:
            st.error(f"処理中にエラーが発生しました: {e}")
            st.warning("Excelの列名が変わっていないか、データに空行が含まれていないか確認してください。")

else:
    st.info("👆 ファイルをアップロードしてください")
