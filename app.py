import streamlit as st
import pandas as pd
import altair as alt

# --- ページ設定 ---
st.set_page_config(page_title="まごころサポート分析 v3", layout="wide")
st.title("📊 まごころサポート：広告×商談 分析ダッシュボード")

# --- データ読み込み関数 ---
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

# --- サイドバー設定 ---
st.sidebar.header("⚙️ 判定基準の設定")
cpa_limit = st.sidebar.number_input("許容CPA（円）", value=15000, step=1000)
meeting_target = st.sidebar.slider("目標商談化率（%）", 0, 30, 10)

# --- ファイルアップロード ---
col1, col2 = st.columns(2)
with col1:
    meta_file = st.file_uploader("📂 Meta広告実績", type=['xlsx', 'csv'])
with col2:
    hs_file = st.file_uploader("📂 HubSpotデータ", type=['xlsx', 'csv'])

st.divider()

# --- 分析実行 ---
if meta_file and hs_file:
    df_meta = load_data(meta_file)
    df_hs = load_data(hs_file)

    if df_meta is not None and df_hs is not None:
        # 列名の特定（ゆらぎ吸収）
        meta_cols = df_meta.columns.astype(str)
        name_col = next((c for c in meta_cols if '名前' in c or 'Name' in c), None)
        spend_col = next((c for c in meta_cols if '消化金額' in c or 'Amount' in c), None)
        res_col = next((c for c in meta_cols if '結果' in c or 'Results' in c), None)

        hs_cols = df_hs.columns.astype(str)
        utm_col = next((c for c in hs_cols if 'UTM' in c or 'Content' in c), None)
        
        # 追加したい詳細項目（存在チェック）
        connect_col = next((c for c in hs_cols if '接続' in c), None)
        deal_col = next((c for c in hs_cols if '商談' in c and '予定' not in c), None) # 商談有無
        deal_plan_col = next((c for c in hs_cols if '商談' in c and '予定' in c), None) # 商談予定
        attr_col = next((c for c in hs_cols if '属性' in c), None)
        stage_col = next((c for c in hs_cols if 'ステージ' in c), None)

        if name_col and spend_col and res_col and utm_col:
            # 1. データ結合キーの作成
            df_meta['key'] = df_meta[name_col].astype(str).str.extract(r'(bn\d+)')[0]
            df_hs['key'] = df_hs[utm_col].astype(str).str.strip()
            
            # 2. バナー別集計（ROI分析用）
            meta_agg = df_meta.groupby('key')[[spend_col, res_col]].sum().reset_index()
            
            # 商談数のカウント（HubSpot側）
            if deal_col:
                # 'あり'や'TRUE'を含むものを商談としてカウント
                hs_deals = df_hs[df_hs[deal_col].fillna('').astype(str).str.contains('あり|TRUE|Yes', case=False)]
                deal_counts = hs_deals.groupby('key').size().reset_index(name='商談数')
            else:
                deal_counts = pd.DataFrame({'key': [], '商談数': []})

            # 結合
            result = pd.merge(meta_agg, deal_counts, on='key', how='left').fillna(0)
            
            # 計算
            result['CPA'] = (result[spend_col] / result[res_col]).replace([float('inf')], 0).fillna(0).astype(int)
            result['商談化率'] = (result['商談数'] / result[res_col]).fillna(0)

            # 判定ロジック
            def judge(row):
                rate_pct = row['商談化率'] * 100
                if rate_pct >= meeting_target and row['CPA'] <= cpa_limit and row['商談数'] > 0: return "🏆 勝ち"
                if rate_pct >= meeting_target and row['商談数'] > 0: return "🟡 質良(CPA高)"
                if row['CPA'] <= cpa_limit: return "🥈 CPA良"
                return "🛑 停止推奨"
            
            result['判定'] = result.apply(judge, axis=1)

            # --- 📊 ダッシュボード表示 ---

            # 全体サマリー
            total_spend = result[spend_col].sum()
            total_leads = result[res_col].sum()
            total_meetings = result['商談数'].sum()
            avg_cpa = int(total_spend / total_leads) if total_leads > 0 else 0
            avg_rate = (total_meetings / total_leads * 100) if total_leads > 0 else 0

            st.subheader("📈 全体実績サマリー")
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("総消化金額", f"¥{total_spend:,.0f}")
            k2.metric("総リード数", f"{int(total_leads)}件")
            k3.metric("総商談数", f"{int(total_meetings)}件", delta=f"{avg_rate:.1f}%")
            k4.metric("平均CPA", f"¥{avg_cpa:,.0f}")

            st.divider()

            # タブで表示を切り替え
            tab1, tab2 = st.tabs(["📊 バナー別成績 (ROI)", "📋 リード詳細リスト (質)"])

            with tab1:
                st.subheader("バナー別 パフォーマンス")
                
                # エラー回避のためのデータクレンジング
                chart_data = result[result[res_col] > 0].copy() # リード0件は除外
                chart_data['商談化率(%)'] = chart_data['商談化率'] * 100
                
                # 散布図 (Altair)
                if not chart_data.empty:
                    chart = alt.Chart(chart_data).mark_circle(size=100).encode(
                        x=alt.X('CPA', title='CPA (円)'),
                        y=alt.Y('商談化率(%)', title='商談化率 (%)'),
                        color='判定',
                        tooltip=['key', 'CPA', '商談化率(%)', '消化金額']
                    ).interactive()
                    st.altair_chart(chart, use_container_width=True)
                else:
                    st.info("データ不足のためグラフを表示できません")

                # テーブル表示
                display_df = result.copy()
                display_df = display_df.rename(columns={'key': 'バナーID', spend_col: '消化金額', res_col: 'リード数'})
                
                st.dataframe(
                    display_df[['判定', 'バナーID', '消化金額', 'リード数', 'CPA', '商談数', '商談化率']].sort_values('消化金額', ascending=False),
                    column_config={
                        "消化金額": st.column_config.NumberColumn(format="¥%d"),
                        "CPA": st.column_config.NumberColumn(format="¥%d"), # カンマ区切り
                        "商談化率": st.column_config.NumberColumn(format="%.1f%%"), # %表示
                    },
                    use_container_width=True,
                    hide_index=True
                )
            
            with tab2:
                st.subheader("リード属性・質の詳細分析")
                st.caption("どのバナーから「どんな人（属性・ステージ）」が来ているかを確認します")
                
                # HubSpotデータをベースに、バナー情報を結合
                detail_df = df_hs.copy()
                # バナーごとのCPAや判定を紐付ける
                detail_df = pd.merge(detail_df, result[['key', '判定', 'CPA']], on='key', how='left')
                
                # 表示する列を選択（存在するものだけ）
                cols_to_show = ['key', '判定']
                # 指定された項目を追加
                target_cols_map = {
                    '属性': attr_col,
                    '接続': connect_col,
                    '商談有無': deal_col,
                    '商談予定': deal_plan_col,
                    'ステージ': stage_col
                }
                
                # 列名を分かりやすくリネームして表示リストに追加
                rename_dict = {'key': '流入バナー'}
                for label, col_name in target_cols_map.items():
                    if col_name:
                        cols_to_show.append(col_name)
                        rename_dict[col_name] = label
                
                # フィルタリング機能
                filter_banner = st.multiselect("バナーで絞り込む", options=detail_df['key'].unique())
                if filter_banner:
                    detail_df = detail_df[detail_df['key'].isin(filter_banner)]

                # テーブル表示
                st.dataframe(
                    detail_df[cols_to_show].rename(columns=rename_dict).fillna('-'),
                    use_container_width=True,
                    hide_index=True
                )

            # アクション提案
            st.divider()
            st.subheader("🤖 AIアクション提案")
            winners = result[result['判定'].str.contains("勝ち")]['key'].tolist()
            if winners:
                st.success(f"**【予算増額！】** : 「{'、'.join(winners)}」は最強です。CPAも安く、商談にも繋がっています。")
            else:
                st.info("圧倒的な勝ちパターンはまだありません。CPAが安いバナーのLPを見直すか、商談率が高いバナーの入札を強めましょう。")

        else:
            st.error("必要なデータ列が見つかりません。ファイルの中身（列名）を確認してください。")
    else:
        st.error("ファイル読み込みエラー")
