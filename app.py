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
        else:
            return pd.read_excel(file)
    except Exception as e:
        st.error(f"ファイル読み込みエラー: {e}")
        return None

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
        try:
            # 列名の特定（ゆらぎ吸収）
            meta_cols = list(df_meta.columns)
            name_col = next((c for c in meta_cols if '名前' in str(c) or 'Name' in str(c)), None)
            spend_col = next((c for c in meta_cols if '消化金額' in str(c) or 'Amount' in str(c)), None)
            res_col = next((c for c in meta_cols if '結果' in str(c) or 'Results' in str(c) or 'リード' in str(c)), None)

            hs_cols = list(df_hs.columns)
            utm_col = next((c for c in hs_cols if 'UTM' in str(c) or 'Content' in str(c)), None)
            connect_col = next((c for c in hs_cols if '接続' in str(c)), None)
            deal_col = next((c for c in hs_cols if '商談' in str(c) and '予定' not in str(c)), None)
            deal_plan_col = next((c for c in hs_cols if '商談' in str(c) and '予定' in str(c)), None)
            attr_col = next((c for c in hs_cols if '属性' in str(c)), None)
            stage_col = next((c for c in hs_cols if 'ステージ' in str(c)), None)

            if not all([name_col, spend_col, res_col, utm_col]):
                st.error(f"必要な列が見つかりません。\n\nMeta: 広告名={name_col}, 消化金額={spend_col}, 結果={res_col}\nHubSpot: UTM={utm_col}")
                st.stop()

            # 1. データ結合キーの作成
            df_meta['key'] = df_meta[name_col].astype(str).str.extract(r'(bn\d+)', expand=False)
            df_hs['key'] = df_hs[utm_col].astype(str).str.strip()
            
            # キーがない行を除外
            df_meta = df_meta[df_meta['key'].notna()]
            df_hs = df_hs[df_hs['key'].notna()]

            # 2. バナー別集計（ROI分析用）
            meta_agg = df_meta.groupby('key').agg({
                spend_col: 'sum',
                res_col: 'sum'
            }).reset_index()
            
            # 数値型に変換（エラー回避）
            meta_agg[spend_col] = pd.to_numeric(meta_agg[spend_col], errors='coerce').fillna(0)
            meta_agg[res_col] = pd.to_numeric(meta_agg[res_col], errors='coerce').fillna(0)

            # 商談数のカウント
            if deal_col:
                hs_deals = df_hs[df_hs[deal_col].fillna('').astype(str).str.contains('あり|TRUE|Yes|true', case=False)]
                deal_counts = hs_deals.groupby('key').size().reset_index(name='商談数')
            else:
                deal_counts = pd.DataFrame({'key': [], '商談数': []})

            # 結合
            result = pd.merge(meta_agg, deal_counts, on='key', how='left')
            result['商談数'] = result['商談数'].fillna(0).astype(int)
            
            # 計算
            result['CPA'] = result.apply(
                lambda x: int(x[spend_col] / x[res_col]) if x[res_col] > 0 else 0, 
                axis=1
            )
            result['商談化率'] = result.apply(
                lambda x: (x['商談数'] / x[res_col]) if x[res_col] > 0 else 0,
                axis=1
            )

            # 判定ロジック
            def judge(row):
                rate_pct = row['商談化率'] * 100
                if rate_pct >= meeting_target and row['CPA'] <= cpa_limit and row['商談数'] > 0: 
                    return "🏆 勝ち"
                if rate_pct >= meeting_target and row['商談数'] > 0: 
                    return "🟡 質良(CPA高)"
                if row['CPA'] <= cpa_limit and row['CPA'] > 0: 
                    return "🥈 CPA良"
                return "🛑 停止推奨"
            
            result['判定'] = result.apply(judge, axis=1)

            # --- 📊 ダッシュボード表示 ---
            total_spend = result[spend_col].sum()
            total_leads = result[res_col].sum()
            total_meetings = result['商談数'].sum()
            avg_cpa = int(total_spend / total_leads) if total_leads > 0 else 0
            avg_rate = (total_meetings / total_leads * 100) if total_leads > 0 else 0

            st.subheader("📈 全体実績サマリー")
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("総消化金額", f"¥{int(total_spend):,}")
            k2.metric("総リード数", f"{int(total_leads)}件")
            k3.metric("総商談数", f"{int(total_meetings)}件", delta=f"{avg_rate:.1f}%")
            k4.metric("平均CPA", f"¥{avg_cpa:,}")

            st.divider()

            # タブで表示を切り替え
            tab1, tab2 = st.tabs(["📊 バナー別成績 (ROI)", "📋 リード詳細リスト (質)"])

            with tab1:
                st.subheader("バナー別 パフォーマンス")
                
                # グラフ用データ
                chart_data = result[result[res_col] > 0].copy()
                chart_data['商談化率(%)'] = chart_data['商談化率'] * 100
                
                if len(chart_data) > 0:
                    chart = alt.Chart(chart_data).mark_circle(size=150).encode(
                        x=alt.X('CPA:Q', title='CPA (円)', scale=alt.Scale(zero=False)),
                        y=alt.Y('商談化率(%):Q', title='商談化率 (%)'),
                        color=alt.Color('判定:N', legend=alt.Legend(title="判定")),
                        size=alt.Size(spend_col, legend=None),
                        tooltip=['key', 'CPA', '商談化率(%)', spend_col, '商談数']
                    ).properties(height=400).interactive()
                    st.altair_chart(chart, use_container_width=True)

                # テーブル表示
                display_df = result.copy()
                display_df['商談化率'] = (display_df['商談化率'] * 100).round(1)
                display_df = display_df.rename(columns={
                    'key': 'バナーID', 
                    spend_col: '消化金額', 
                    res_col: 'リード数'
                })
                
                st.dataframe(
                    display_df[['判定', 'バナーID', '消化金額', 'リード数', 'CPA', '商談数', '商談化率']].sort_values('消化金額', ascending=False),
                    use_container_width=True,
                    hide_index=True
                )
            
            with tab2:
                st.subheader("リード属性・質の詳細分析")
                
                detail_df = df_hs.copy()
                detail_df = pd.merge(detail_df, result[['key', '判定', 'CPA']], on='key', how='left')
                
                cols_to_show = ['key', '判定']
                rename_dict = {'key': '流入バナー'}
                
                target_cols = {
                    '属性': attr_col,
                    '接続': connect_col,
                    '商談有無': deal_col,
                    '商談予定': deal_plan_col,
                    'ステージ': stage_col
                }
                
                for label, col in target_cols.items():
                    if col and col in detail_df.columns:
                        cols_to_show.append(col)
                        rename_dict[col] = label
                
                if len(detail_df) > 0:
                    filter_banner = st.multiselect("バナーで絞り込む", options=sorted(detail_df['key'].unique().tolist()))
                    if filter_banner:
                        detail_df = detail_df[detail_df['key'].isin(filter_banner)]

                    st.dataframe(
                        detail_df[cols_to_show].rename(columns=rename_dict).fillna('-'),
                        use_container_width=True,
                        hide_index=True
                    )

            # アクション提案
            st.divider()
            st.subheader("🤖 AIアクション提案")
            winners = result[result['判定'].str.contains("勝ち", na=False)]['key'].tolist()
            if winners:
                st.success(f"**【予算増額！】** {', '.join(winners)} は最強パターン。CPAも安く商談に繋がっています。")
            else:
                st.info("現状、圧倒的な勝ちパターンはなし。CPA優秀バナーのLP改善か、商談率高バナーの予算増を検討してください。")

        except Exception as e:
            st.error(f"処理エラー: {e}")
            st.code(str(e))

else:
    st.info("👆 2つのファイルをアップロードすると分析が始まります")
