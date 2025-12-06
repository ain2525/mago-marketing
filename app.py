import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime, timedelta

# --- ページ設定 ---
st.set_page_config(page_title="まごころサポート分析 v10", layout="wide")
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
            # === Meta側：列の特定 ===
            meta_cols = list(df_meta.columns)
            name_col = next((c for c in meta_cols if '名前' in str(c) or 'Name' in str(c)), None)
            spend_col = next((c for c in meta_cols if '消化金額' in str(c) or 'Amount' in str(c) or '費用' in str(c)), None)
            date_col_meta = next((c for c in meta_cols if '日' in str(c) or 'Date' in str(c) or '開始' in str(c)), None)

            # === HubSpot側：列の特定 ===
            hs_cols = list(df_hs.columns)
            utm_col = next((c for c in hs_cols if 'UTM' in str(c) or 'Content' in str(c)), None)
            connect_col = next((c for c in hs_cols if '接続' in str(c)), None)
            deal_col = next((c for c in hs_cols if '商談' in str(c)), None)
            date_col_hs = next((c for c in hs_cols if '作成日' in str(c) or 'Created' in str(c) or '日付' in str(c)), None)

            if not all([name_col, spend_col, utm_col]):
                st.error(f"必要な列が見つかりません。\nMeta: 広告名={name_col}, 消化金額={spend_col}\nHubSpot: UTM={utm_col}")
                st.stop()

            # === 日付列の変換 ===
            if date_col_meta:
                df_meta[date_col_meta] = pd.to_datetime(df_meta[date_col_meta], errors='coerce')
            if date_col_hs:
                df_hs[date_col_hs] = pd.to_datetime(df_hs[date_col_hs], errors='coerce')

            # === 期間フィルター（サイドバーに移動） ===
            st.sidebar.markdown("---")
            st.sidebar.subheader("📅 分析期間の設定")
            filter_enabled = st.sidebar.checkbox("期間で絞り込む", value=False)
            
            if filter_enabled:
                if date_col_hs:
                    min_date_hs = df_hs[date_col_hs].min()
                    max_date_hs = df_hs[date_col_hs].max()
                    start_date = st.sidebar.date_input("開始日", value=min_date_hs if pd.notna(min_date_hs) else datetime.now() - timedelta(days=30))
                    end_date = st.sidebar.date_input("終了日", value=max_date_hs if pd.notna(max_date_hs) else datetime.now())
                else:
                    start_date = st.sidebar.date_input("開始日", value=datetime.now() - timedelta(days=30))
                    end_date = st.sidebar.date_input("終了日", value=datetime.now())
                
                start_datetime = pd.to_datetime(start_date)
                end_datetime = pd.to_datetime(end_date) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
                
                # フィルタリング実行
                if date_col_meta:
                    df_meta = df_meta[(df_meta[date_col_meta] >= start_datetime) & (df_meta[date_col_meta] <= end_datetime)]
                if date_col_hs:
                    df_hs = df_hs[(df_hs[date_col_hs] >= start_datetime) & (df_hs[date_col_hs] <= end_datetime)]
                
                st.sidebar.success(f"📊 {start_date} 〜 {end_date}")

            # === デバッグ情報 ===
            st.sidebar.markdown("---")
            st.sidebar.subheader("🔍 デバッグ情報")
            st.sidebar.write(f"**Meta広告データ:** {len(df_meta)}行")
            st.sidebar.write(f"**HubSpotデータ:** {len(df_hs)}行")
            
            if deal_col:
                st.sidebar.write(f"**商談列名:** `{deal_col}`")
                deal_values = df_hs[deal_col].fillna('(空白)').astype(str).value_counts()
                st.sidebar.write("**商談列の値:**")
                st.sidebar.dataframe(deal_values, use_container_width=True)

            # === 1. データ結合キーの作成 ===
            df_meta['key'] = df_meta[name_col].astype(str).str.extract(r'(bn\d+)', expand=False)
            df_hs['key'] = df_hs[utm_col].astype(str).str.strip()
            
            df_meta = df_meta[df_meta['key'].notna()]
            df_hs = df_hs[df_hs['key'].notna()]

            # === 2. Meta側の消化金額集計（フィルター後） ===
            meta_spend = df_meta.groupby('key')[spend_col].sum().reset_index()
            meta_spend[spend_col] = pd.to_numeric(meta_spend[spend_col], errors='coerce').fillna(0)
            
            st.sidebar.markdown("---")
            st.sidebar.write("**Meta消化金額（バナー別）:**")
            st.sidebar.dataframe(meta_spend.rename(columns={'key': 'バナー', spend_col: '消化金額'}), use_container_width=True)

            # === 3. HubSpot側でリード数・接続・商談をカウント ===
            hs_summary = df_hs.groupby('key').agg(
                リード数=('key', 'size')
            ).reset_index()

            # 接続数
            if connect_col:
                connect_df = df_hs[df_hs[connect_col].fillna('').astype(str).str.contains('あり|TRUE|Yes|true|済', case=False, na=False)]
                connect_count = connect_df.groupby('key').size().reset_index(name='接続数')
                hs_summary = pd.merge(hs_summary, connect_count, on='key', how='left')
            else:
                hs_summary['接続数'] = 0
            
            # 商談実施数・予約数
            if deal_col:
                df_hs['商談_normalized'] = df_hs[deal_col].fillna('').astype(str).str.lower().str.strip()
                
                deal_done = df_hs[
                    (df_hs['商談_normalized'].str.contains('あり|済|完了|実施|done|yes|true', case=False, na=False)) &
                    (~df_hs['商談_normalized'].str.contains('予約|予定|scheduled', case=False, na=False))
                ]
                deal_done_count = deal_done.groupby('key').size().reset_index(name='商談実施数')
                
                deal_plan = df_hs[df_hs['商談_normalized'].str.contains('予約|予定|scheduled', case=False, na=False)]
                deal_plan_count = deal_plan.groupby('key').size().reset_index(name='商談予約数')
                
                st.sidebar.write(f"✅ 商談実施: **{len(deal_done)}件**")
                st.sidebar.write(f"📅 商談予約: **{len(deal_plan)}件**")
                
                hs_summary = pd.merge(hs_summary, deal_done_count, on='key', how='left')
                hs_summary = pd.merge(hs_summary, deal_plan_count, on='key', how='left')
            else:
                hs_summary['商談実施数'] = 0
                hs_summary['商談予約数'] = 0

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
                cpa_ok = row['CPA'] > 0 and row['CPA'] <= cpa_limit
                connect_ok = row['接続率'] >= connect_target
                meeting_ok = row['商談化率'] >= meeting_target
                
                conditions_met = sum([cpa_ok, connect_ok, meeting_ok])
                
                if conditions_met == 3:
                    return "🏆 最優秀"
                elif conditions_met == 2 and meeting_ok:
                    return "🥇 優秀"
                elif conditions_met == 2:
                    return "🟡 要改善"
                elif conditions_met == 1 and meeting_ok:
                    return "🟡 要改善"
                else:
                    return "🛑 停止推奨"
            
            result['判定'] = result.apply(judge, axis=1)

            # === 7. 全体サマリー（枠サイズ統一版） ===
            total_spend = result[spend_col].sum()
            total_leads = result['リード数'].sum()
            total_connect = result['接続数'].sum()
            total_deal = result['商談実施数'].sum()
            total_plan = result['商談予約数'].sum()
            avg_cpa = int(total_spend / total_leads) if total_leads > 0 else 0
            avg_connect = (total_connect / total_leads * 100) if total_leads > 0 else 0
            avg_meeting = ((total_deal + total_plan) / total_leads * 100) if total_leads > 0 else 0

            st.subheader("📈 全体実績サマリー")

            cols = st.columns(6)

            with cols[0]:
                st.markdown("""
                <div style='border: 1px solid #ddd; border-radius: 8px; padding: 16px; background-color: #f8f9fa; min-height: 100px;'>
                    <p style='margin: 0; font-size: 0.85rem; color: #666;'>総消化金額</p>
                    <p style='margin: 8px 0 0 0; font-size: 1.4rem; font-weight: bold;'>¥{:,}</p>
                </div>
                """.format(int(total_spend)), unsafe_allow_html=True)

            with cols[1]:
                st.markdown("""
                <div style='border: 1px solid #ddd; border-radius: 8px; padding: 16px; background-color: #f8f9fa; min-height: 100px;'>
                    <p style='margin: 0; font-size: 0.85rem; color: #666;'>総リード数</p>
                    <p style='margin: 8px 0 0 0; font-size: 1.4rem; font-weight: bold;'>{}件</p>
                </div>
                """.format(int(total_leads)), unsafe_allow_html=True)

            with cols[2]:
                st.markdown("""
                <div style='border: 1px solid #ddd; border-radius: 8px; padding: 16px; background-color: #f8f9fa; min-height: 100px;'>
                    <p style='margin: 0; font-size: 0.85rem; color: #666;'>接続数</p>
                    <p style='margin: 8px 0 0 0; font-size: 1.4rem; font-weight: bold;'>{}件</p>
                    <p style='margin: 4px 0 0 0; font-size: 0.75rem; color: #28a745;'>接続率 {:.1f}%</p>
                </div>
                """.format(int(total_connect), avg_connect), unsafe_allow_html=True)

            with cols[3]:
                st.markdown("""
                <div style='border: 1px solid #ddd; border-radius: 8px; padding: 16px; background-color: #f8f9fa; min-height: 100px;'>
                    <p style='margin: 0; font-size: 0.85rem; color: #666;'>平均CPA</p>
                    <p style='margin: 8px 0 0 0; font-size: 1.4rem; font-weight: bold;'>¥{:,}</p>
                </div>
                """.format(avg_cpa), unsafe_allow_html=True)

            with cols[4]:
                st.markdown("""
                <div style='border: 1px solid #ddd; border-radius: 8px; padding: 16px; background-color: #f8f9fa; min-height: 100px;'>
                    <p style='margin: 0; font-size: 0.85rem; color: #666;'>商談実施数</p>
                    <p style='margin: 8px 0 0 0; font-size: 1.4rem; font-weight: bold;'>{}件</p>
                </div>
                """.format(int(total_deal)), unsafe_allow_html=True)

            with cols[5]:
                st.markdown("""
                <div style='border: 1px solid #ddd; border-radius: 8px; padding: 16px; background-color: #f8f9fa; min-height: 100px;'>
                    <p style='margin: 0; font-size: 0.85rem; color: #666;'>商談予約数</p>
                    <p style='margin: 8px 0 0 0; font-size: 1.4rem; font-weight: bold;'>{}件</p>
                    <p style='margin: 4px 0 0 0; font-size: 0.75rem; color: #28a745;'>商談化率 {:.1f}%</p>
                </div>
                """.format(int(total_plan), avg_meeting), unsafe_allow_html=True)

            st.divider()

            # === 8. バナー別評価表（上に移動） ===
            st.subheader("📋 バナー別 評価表")
            
            display_df = result.copy()
            display_df = display_df.rename(columns={
                'key': 'バナーID',
                spend_col: '消化金額'
            })
            
            display_df['消化金額_表示'] = display_df['消化金額'].apply(lambda x: f"{int(x):,}")
            display_df['CPA_表示'] = display_df['CPA'].apply(lambda x: f"{int(x):,}")
            display_df['接続率_表示'] = display_df['接続率'].apply(lambda x: f"{x:.1f}%")
            display_df['商談化率_表示'] = display_df['商談化率'].apply(lambda x: f"{x:.1f}%")
            
            show_df = display_df[['判定', 'バナーID', '消化金額_表示', 'リード数', 'CPA_表示', '接続率_表示', '商談化率_表示', '商談実施数', '商談予約数']].copy()
            show_df.columns = ['判定', 'バナーID', '消化金額', 'リード数', 'CPA', '接続率', '商談化率', '商談実施数', '商談予約数']
            show_df = show_df.sort_values(by='商談化率', ascending=False, key=lambda x: display_df['商談化率'])
            
            def highlight_row(row):
                判定 = row['判定']
                if 判定 == "🏆 最優秀":
                    color = 'background-color: #d4edda'
                elif "優秀" in 判定:
                    color = 'background-color: #d1ecf1'
                elif "要改善" in 判定:
                    color = 'background-color: #fff3cd'
                elif "停止" in 判定:
                    color = 'background-color: #f8d7da'
                else:
                    color = ''
                return [color] * len(row)
            
            st.dataframe(
                show_df.style.apply(highlight_row, axis=1),
                use_container_width=True,
                hide_index=True
            )

            # === 9. AIアクション提案 ===
            st.divider()
            st.subheader("🤖 AIによる評価と推奨アクション")
            
            best = result[result['判定'] == "🏆 最優秀"]['key'].tolist()
            good = result[result['判定'].str.contains("優秀", na=False)]['key'].tolist()
            improve = result[result['判定'].str.contains("要改善", na=False)]['key'].tolist()
            stop = result[result['判定'] == "🛑 停止推奨"]['key'].tolist()
            
            if best:
                st.success(f"**【予算集中！】** {', '.join(best)} → CPA・接続率・商談化率すべて基準クリア。予算を最大化してください。")
            if good:
                good_filtered = [b for b in good if b not in best]
                if good_filtered:
                    st.info(f"**【有望株】** {', '.join(good_filtered)} → 商談化率は目標達成。CPA or 接続率を改善すれば最優秀に。")
            if improve:
                st.warning(f"**【要分析】** {', '.join(improve)} → LP改善や接続体制の見直しを検討してください。")
            if stop:
                st.error(f"**【停止検討】** {', '.join(stop)} → 予算を優秀バナーに振り替えてください。")

            # === 10. バナー別総合評価（最下部に移動） ===
            st.divider()
            st.subheader("📊 バナー別 総合評価（分布図）")
            
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

        except Exception as e:
            st.error(f"処理エラー: {e}")
            import traceback
            st.code(traceback.format_exc())

else:
    st.info("👆 2つのファイルをアップロードすると分析が始まります")
