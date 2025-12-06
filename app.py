import streamlit as st
import pandas as pd
import altair as alt

# ページ設定
st.set_page_config(page_title="まごころサポート分析 v2", layout="wide")
st.title("📊 まごころサポート：広告×商談 分析ダッシュボード")

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
st.sidebar.header("⚙️ 判定基準の設定")
cpa_limit = st.sidebar.number_input("許容CPA（円）", value=15000, step=1000)
meeting_target = st.sidebar.slider("目標商談化率（%）", 0, 30, 10)

# アップロードエリア
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
            # キー作成と結合
            df_meta['key'] = df_meta[name_col].astype(str).str.extract(r'(bn\d+)')[0]
            df_hs['key'] = df_hs[utm_col].astype(str).str.strip()
            
            # 集計
            meta_agg = df_meta.groupby('key')[[spend_col, res_col]].sum().reset_index()
            
            if deal_col:
                hs_deals = df_hs[df_hs[deal_col].fillna('').astype(str).str.contains('あり|TRUE|Yes', case=False)]
                deal_counts = hs_deals.groupby('key').size().reset_index(name='商談数')
            else:
                deal_counts = pd.DataFrame({'key': [], '商談数': []})

            result = pd.merge(meta_agg, deal_counts, on='key', how='left').fillna(0)
            
            # 計算
            result['CPA'] = (result[spend_col] / result[res_col]).replace([float('inf')], 0).fillna(0).astype(int)
            result['商談化率'] = (result['商談数'] / result[res_col]).fillna(0) # %計算は表示時に行う

            # 判定ロジック
            def judge(row):
                rate_pct = row['商談化率'] * 100
                if rate_pct >= meeting_target and row['CPA'] <= cpa_limit and row['商談数'] > 0: return "🏆 勝ち"
                if rate_pct >= meeting_target and row['商談数'] > 0: return "🟡 質良(CPA高)"
                if row['CPA'] <= cpa_limit: return "🥈 CPA良(商談低)"
                return "🛑 停止推奨"
            
            result['判定'] = result.apply(judge, axis=1)
            
            # --- 📊 ダッシュボード表示 ---

            # 1. 全体サマリー (KPI)
            total_spend = result[spend_col].sum()
            total_leads = result[res_col].sum()
            total_meetings = result['商談数'].sum()
            avg_cpa = int(total_spend / total_leads) if total_leads > 0 else 0
            avg_rate = (total_meetings / total_leads * 100) if total_leads > 0 else 0

            st.subheader("📈 全体実績サマリー")
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric("総消化金額", f"¥{total_spend:,}")
            kpi2.metric("総リード数", f"{int(total_leads)}件")
            kpi3.metric("総商談数", f"{int(total_meetings)}件", delta=f"{avg_rate:.1f}%")
            kpi4.metric("平均CPA", f"¥{avg_cpa:,}")

            st.divider()

            # 2. メイン分析テーブル (整形済み)
            st.subheader("🔍 バナー別 詳細分析")
            
            # 表示用データの整理
            display_df = result.copy()
            display_df = display_df.rename(columns={
                'key': 'バナーID',
                spend_col: '消化金額',
                res_col: 'リード数'
            })
            
            # データフレーム表示（フォーマット指定）
            st.dataframe(
                display_df[['判定', 'バナーID', '消化金額', 'リード数', 'CPA', '商談数', '商談化率']].sort_values('消化金額', ascending=False),
                column_config={
                    "消化金額": st.column_config.NumberColumn(format="¥%d"),
                    "CPA": st.column_config.NumberColumn(format="¥%d"),
                    "商談化率": st.column_config.NumberColumn(format="%.1f%%"), # ここで%表示
                },
                use_container_width=True,
                hide_index=True
            )

            # 3. グラフ分析 (散布図)
            st.subheader("💠 ポートフォリオ分析 (横軸:CPA × 縦軸:商談化率)")
            st.caption("右上にあるほど「安く・質の良い」最強のバナーです")
            
            chart_data = result.copy()
            chart_data['商談化率(%)'] = chart_data['商談化率'] * 100
            
            chart = alt.Chart(chart_data).mark_circle(size=100).encode(
                x=alt.X('CPA', title='CPA (低いほど良い)'),
                y=alt.Y('商談化率(%)', title='商談化率 (高いほど良い)'),
                color='判定',
                tooltip=['key', 'CPA', '商談化率(%)', '消化金額']
            ).interactive()
            
            st.altair_chart(chart, use_container_width=True)

            # 4. アクション提案
            st.subheader("🤖 AIアクション提案")
            winners = result[result['判定'].str.contains("勝ち")]['key'].tolist()
            stops = result[result['判定'].str.contains("停止")]['key'].tolist()
            
            if winners:
                st.success(f"**【予算増額！】** : 「{'、'.join(winners)}」はCPA・商談率ともに基準クリアです。予算を寄せて件数を最大化しましょう。")
            else:
                st.info("現在、完璧な「勝ちパターン」は見つかっていません。CPAが安いバナーのLPを改善するか、商談率が高いバナーの入札を調整しましょう。")
                
            if stops:
                st.error(f"**【停止検討】** : 「{'、'.join(stops)}」は成果が出ていません。クリエイティブを差し替えましょう。")

        else:
            st.error("必要な列が見つかりませんでした。Metaデータの「広告の名前」、HubSpotの「UTM Content」などを確認してください。")
    else:
        st.error("ファイルの読み込みに失敗しました。")
