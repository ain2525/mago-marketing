import streamlit as st
import pandas as pd
import altair as alt

# --- ページ設定 ---
st.set_page_config(page_title="まごころサポート分析 v4", layout="wide")
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
cpa_limit = st.sidebar.number_input("許容CPA（円）", value=10000, step=1000)
connect_target = st.sidebar.slider("目標接続率（%）", 0, 100, 50)
meeting_target = st.sidebar.slider("目標商談化率（%）", 0, 50, 18)

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
            # === Meta側：消化金額の取得用 ===
            meta_cols = list(df_meta.columns)
            name_col = next((c for c in meta_cols if '名前' in str(c) or 'Name' in str(c)), None)
            spend_col = next((c for c in meta_cols if '消化金額' in str(c) or 'Amount' in str(c)), None)

            # === HubSpot側：すべての分析基準 ===
            hs_cols = list(df_hs.columns)
            utm_col = next((c for c in hs_cols if 'UTM' in str(c) or 'Content' in str(c)), None)
            connect_col = next((c for c in hs_cols if '接続' in str(c)), None)
            deal_col = next((c for c in hs_cols if '商談' in str(c) and '予定' not in str(c)), None)
            deal_plan_col = next((c for c in hs_cols if '商談' in str(c) and '予定' in str(c)), None)
            attr_col = next((c for c in hs_cols if '属性' in str(c)), None)
            stage_col = next((c for c in hs_cols if 'ステージ' in str(c) or 'Stage' in str(c)), None)

            if not all([name_col, spend_col, utm_col]):
                st.error(f"必要な列が見つかりません。Meta: {name_col}/{spend_col}, HubSpot: {utm_col}")
                st.stop()

            # === 1. データ結合キーの作成 ===
            df_meta['key'] = df_meta[name_col].astype(str).str.extract(r'(bn\d+)', expand=False)
            df_hs['key'] = df_hs[utm_col].astype(str).str.strip()
            
            # キーがない行を除外
            df_meta = df_meta[df_meta['key'].notna()]
            df_hs = df_hs[df_hs['key'].notna()]

            # === 2. Meta側の消化金額集計 ===
            meta_spend = df_meta.groupby('key')[spend_col].sum().reset_index()
            meta_spend[spend_col] = pd.to_numeric(meta_spend[spend_col], errors='coerce').fillna(0)

            # === 3. HubSpot側でリード数・接続・商談をカウント ===
            hs_summary = df_hs.groupby('key').agg(
                リード数=('key', 'size')
            ).reset_index()

            # 接続数（「接続」列が「あり」「TRUE」などの場合）
            if connect_col:
                connect_df = df_hs[df_hs[connect_col].fillna('').astype(str).str.contains('あり|TRUE|Yes|true|済', case=False, na=False)]
                connect_count = connect_df.groupby('key').size().reset_index(name='接続数')
                hs_summary = pd.merge(hs_summary, connect_count, on='key', how='left')
            else:
                hs_summary['接続数'] = 0
            
            # 商談実施数
            if deal_col:
                deal_df = df_hs[df_hs[deal_col].fillna('').astype(str).str.contains('あり|TRUE|Yes|true|済', case=False, na=False)]
                deal_count = deal_df.groupby('key').size().reset_index(name='商談実施数')
                hs_summary = pd.merge(hs_summary, deal_count, on='key', how='left')
            else:
                hs_summary['商談実施数'] = 0

            # 商談予約数
            if deal_plan_col:
                plan_df = df_hs[df_hs[deal_plan_col].fillna('').astype(str).str.contains('あり|TRUE|Yes|true|済|予定', case=False, na=False)]
                plan_count = plan_df.groupby('key').size().reset_index(name='商談予約数')
                hs_summary = pd.merge(hs_summary, plan_count, on='key', how='left')
            else:
                hs_summary['商談予約数'] = 0

            # 欠損値を0埋め
            hs_summary = hs_summary.fillna(0)
            hs_summary['接続数'] = hs_summary['接続数'].astype(int)
            hs_summary['商談実施数'] = hs_summary['商談実施数'].astype(int)
            hs_summary['商談予約数'] = hs_summary['商談予約数'].astype(int)

            # === 4. Meta消化金額と結合 ===
            result = pd.merge(hs_summary, meta_spend, on='key', how='left')
            result[spend_col] = result[spend_col].fillna(0)

            # === 5. 指標計算 ===
            result['CPA'] = result.apply(
                lambda x: int(x[spend_col] / x['リード数']) if x['リード数'] > 0 else 0,
                axis=1
            )
            result['接続率'] = result.apply(
                lambda x: (x['接続数'] / x['リード数'] * 100) if x['リード数'] > 0 else 0,
                axis=1
            )
            result['商談化率'] = result.apply(
                lambda x: ((x['商談実施数'] + x['商談予約数']) / x['リード数'] * 100) if x['リード数'] > 0 else 0,
                axis=1
            )

            # === 6. 判定ロジック ===
            def judge(row):
                conditions_met = 0
                if row['CPA'] > 0 and row['CPA'] <= cpa_limit:
                    conditions_met += 1
                if row['接続率'] >= connect_target:
                    conditions_met += 1
                if row['商談化率'] >= meeting_target:
                    conditions_met += 1
                
                if conditions_met == 3:
                    return "🏆 最優秀"
                elif conditions_met == 2:
                    return "🥇 優秀"
                elif conditions_met == 1:
                    return "🟡 要改善"
                else:
                    return "🛑 停止推奨"
            
            result['判定'] = result.apply(judge, axis=1)

            # === 7. 全体サマリー ===
            total_spend = result[spend_col].sum()
            total_leads = result['リード数'].sum()
            total_connect = result['接続数'].sum()
            total_deal = result['商談実施数'].sum()
            total_plan = result['商談予約数'].sum()
            avg_cpa = int(total_spend / total_leads) if total_leads > 0 else 0
            avg_connect = (total_connect / total_leads * 100) if total_leads > 0 else 0
            avg_meeting = ((total_deal + total_plan) / total_leads * 100) if total_leads > 0 else 0

            st.subheader("📈 全体実績サマリー")
            k1, k2, k3 = st.columns(3)
            k1.metric("総消化金額", f"¥{int(total_spend):,}")
            k1.metric("総リード数", f"{int(total_leads)}件")
            
            k2.metric("接続数", f"{int(total_connect)}件", delta=f"{avg_connect:.1f}%")
            k2.metric("平均CPA", f"¥{avg_cpa:,}")
            
            k3.metric("商談実施数", f"{int(total_deal)}件")
            k3.metric("商談予約数", f"{int(total_plan)}件", delta=f"化率{avg_meeting:.1f}%")

            st.divider()

            # === 8. タブ表示 ===
            tab1, tab2 = st.tabs(["📊 バナー別パフォーマンス", "📋 リード詳細リスト"])

            with tab1:
                st.subheader("バナー別 総合評価")
                
                # グラフ
                chart_data = result[result['リード数'] > 0].copy()
                if len(chart_data) > 0:
                    chart = alt.Chart(chart_data).mark_circle(size=200).encode(
                        x=alt.X('CPA:Q', title='CPA (円)', scale=alt.Scale(zero=False)),
                        y=alt.Y('商談化率:Q', title='商談化率 (%)'),
                        color=alt.Color('判定:N', legend=alt.Legend(title="判定")),
                        size=alt.Size('リード数:Q', legend=None),
                        tooltip=['key', 'CPA', '接続率', '商談化率', 'リード数', '判定']
                    ).properties(height=400).interactive()
                    st.altair_chart(chart, use_container_width=True)

                # テーブル
                display_df = result.copy()
                display_df = display_df.rename(columns={
                    'key': 'バナーID',
                    spend_col: '消化金額'
                })
                display_df['接続率'] = display_df['接続率'].round(1)
                display_df['商談化率'] = display_df['商談化率'].round(1)
                
                st.dataframe(
                    display_df[['判定', 'バナーID', '消化金額', 'リード数', 'CPA', '接続率', '商談化率', '商談実施数', '商談予約数']].sort_values('消化金額', ascending=False),
                    use_container_width=True,
                    hide_index=True
                )

            with tab2:
                st.subheader("リード属性・質の詳細")
                
                detail_df = df_hs.copy()
                detail_df = pd.merge(detail_df, result[['key', '判定', 'CPA']], on='key', how='left')
                
                cols_to_show = ['key', '判定', 'CPA']
                rename_dict = {'key': 'バナー'}
                
                target_cols = {
                    '属性': attr_col,
                    '接続': connect_col,
                    '商談': deal_col,
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

            # === 9. AIアクション提案 ===
            st.divider()
            st.subheader("🤖 AIによる評価と推奨アクション")
            
            best = result[result['判定'] == "🏆 最優秀"]['key'].tolist()
            good = result[result['判定'] == "🥇 優秀"]['key'].tolist()
            stop = result[result['判定'] == "🛑 停止推奨"]['key'].tolist()
            
            if best:
                st.success(f"**【予算集中！】** {', '.join(best)} → CPA・接続率・商談化率すべて基準クリア。予算を最大化してください。")
            if good:
                st.info(f"**【改善余地あり】** {', '.join(good)} → 2つの指標は合格。残り1つを改善すれば最優秀に。")
            if stop:
                st.warning(f"**【停止検討】** {', '.join(stop)} → 3指標すべて基準未達。予算を優秀バナーに振り替えてください。")

        except Exception as e:
            st.error(f"処理エラー: {e}")
            import traceback
            st.code(traceback.format_exc())

else:
    st.info("👆 2つのファイルをアップロードすると分析が始まります")
