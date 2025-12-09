import streamlit as st
import pandas as pd
import altair as alt
import os
import json
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
from PIL import Image

# =========================================================================
# 【１】設定とスプレッドシート書き込み関数
# =========================================================================

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1dJwYYK-koOgU0V9z83hfz-Wjjjl_UNbl_N6eHQk5OmI/edit"
KPI_SHEET_INDEX = 0

def write_analysis_to_sheet(analysis_data, spreadsheet_url, sheet_index):
    import numpy as np
    from datetime import datetime
    import streamlit as st
    import gspread

    status_container = st.sidebar.empty()
    status_container.info("スプレッドシートへ書き込み中...")

    try:
        client = gspread.service_account_from_dict(
            st.secrets["google_sheets"]
        )

        workbook = client.open_by_url(spreadsheet_url)
        sheet = workbook.get_worksheet(sheet_index)

        if sheet is None:
            raise ValueError(f"sheet_index {sheet_index} が存在しません")

        def normalize(v):
            if isinstance(v, (np.integer,)):
                return int(v)
            if isinstance(v, (np.floating,)):
                return float(v)
            return v

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        data_to_write = [
            now,
            analysis_data.get("ファイル名", ""),
            normalize(analysis_data.get("総リード数", 0)),
            normalize(analysis_data.get("平均CPA", 0)),
            normalize(analysis_data.get("総消化金額", 0)),
            normalize(round(float(analysis_data.get("商談化率", 0)), 2)),
        ]

        data_to_write = [str(v) if not isinstance(v, (int, float)) else v for v in data_to_write]

        sheet.append_row(data_to_write)
        status_container.success("✅ KPIをスプレッドシートに反映しました")

    except Exception as e:
        status_container.error(f"❌ 書き込み失敗: {e}")


# =========================================================================
# 【２】アプリのメイン処理
# =========================================================================

st.set_page_config(page_title="Meta広告×セールスダッシュボード", layout="wide")

try:
    logo = Image.open("logo.png")
    st.markdown("<div style='text-align: center;'>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.image(logo, use_column_width=True)
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("<h1 style='text-align: center; margin-top: 20px;'>Meta広告×セールスダッシュボード</h1>", unsafe_allow_html=True)
except:
    st.markdown("<h1 style='text-align: center;'>Meta広告×セールスダッシュボード</h1>", unsafe_allow_html=True)

st.markdown("---")

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

st.sidebar.header("判定基準の設定")
cpa_limit = st.sidebar.number_input("許容CPA（円）", value=10000, step=1000)
connect_target = st.sidebar.slider("目標接続率（%）", 0, 100, 50)
meeting_target = st.sidebar.slider("目標商談化率（%）", 0, 50, 18)

st.sidebar.markdown("---")
st.sidebar.subheader("分析期間の設定")

col1, col2 = st.columns(2)
with col1:
    meta_file = st.file_uploader("Meta広告実績", type=['xlsx', 'csv'])
with col2:
    hs_file = st.file_uploader("HubSpotデータ", type=['xlsx', 'csv'])

st.markdown("---")

if meta_file and hs_file:
    df_meta = load_data(meta_file)
    df_hs = load_data(hs_file)

    if df_meta is not None and df_hs is not None:
        try:
            meta_cols = list(df_meta.columns)
            hs_cols = list(df_hs.columns)
            
            # === Meta側：列の特定（優先順位：広告の名前 > 広告名 > 広告セット名 > キャンペーン名）===
            name_col = None
            for pattern in ['広告の名前', '広告名', '広告セット名', 'キャンペーン名', 'Ad name', 'Ad set name', 'Campaign name', '名前', 'Name']:
                name_col = next((c for c in meta_cols if pattern in str(c)), None)
                if name_col:
                    break
            
            spend_col = next((c for c in meta_cols if '消化金額' in str(c)), None)
            if spend_col is None:
                spend_col = next((c for c in meta_cols if 'Amount' in str(c) or '費用' in str(c) or 'Spent' in str(c)), None)
            
            date_col_meta = next((c for c in meta_cols if 'レポート開始日' in str(c) or '開始日' in str(c)), None)

            # === HubSpot側：列の特定 ===
            utm_col = next((c for c in hs_cols if 'UTM Content' in str(c)), None)
            if utm_col is None:
                utm_col = next((c for c in hs_cols if 'UTM' in str(c) or 'Content' in str(c)), None)
            
            attr_col = next((c for c in hs_cols if '属性' in str(c)), None)
            date_col_hs = next((c for c in hs_cols if '作成日' in str(c)), None)
            
            # 列が見つからない場合のエラー表示
            if not all([name_col, spend_col, utm_col]):
                st.error(f"必要な列が見つかりません。")
                st.write(f"Meta: 広告名={name_col}, 消化金額={spend_col}")
                st.write(f"HubSpot: UTM={utm_col}")
                
                with st.expander("Meta列名一覧"):
                    st.write(meta_cols)
                with st.expander("HubSpot列名一覧"):
                    st.write(hs_cols)
                st.stop()

            # === 日付列の変換 ===
            if date_col_meta:
                df_meta[date_col_meta] = pd.to_datetime(df_meta[date_col_meta], errors='coerce')
            if date_col_hs:
                df_hs[date_col_hs] = pd.to_datetime(df_hs[date_col_hs], errors='coerce')

            # === 期間フィルター ===
            filter_enabled = st.sidebar.checkbox("期間で絞り込む", value=False)

            if filter_enabled:
                period_preset = st.sidebar.radio(
                    "期間を選択",
                    options=["今月", "先月", "カスタム"],
                    index=2
                )
                
                now = datetime.now()
                
                if period_preset == "今月":
                    start_date = datetime(now.year, now.month, 1).date()
                    end_date = now.date()
                    st.sidebar.info(f"📅 今月: {start_date} ~ {end_date}")
                    
                elif period_preset == "先月":
                    first_day_this_month = datetime(now.year, now.month, 1)
                    last_day_last_month = first_day_this_month - timedelta(days=1)
                    first_day_last_month = datetime(last_day_last_month.year, last_day_last_month.month, 1)
                    start_date = first_day_last_month.date()
                    end_date = last_day_last_month.date()
                    st.sidebar.info(f"📅 先月: {start_date} ~ {end_date}")
                    
                else:
                    if date_col_hs:
                        min_date_hs = df_hs[date_col_hs].min()
                        max_date_hs = df_hs[date_col_hs].max()
                        start_date = st.sidebar.date_input("開始日", value=min_date_hs if pd.notna(min_date_hs) else datetime.now() - timedelta(days=30))
                        end_date = st.sidebar.date_input("終了日", value=max_date_hs if pd.notna(max_date_hs) else datetime.now())
                    else:
                        start_date = st.sidebar.date_input("開始日", value=datetime.now() - timedelta(days=30))
                        end_date = st.sidebar.date_input("終了日", value=datetime.now())
                    st.sidebar.info(f"📅 カスタム: {start_date} ~ {end_date}")

                start_datetime = pd.to_datetime(start_date)
                end_datetime = pd.to_datetime(end_date) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

                if date_col_meta:
                    before_count = len(df_meta)
                    df_meta = df_meta[(df_meta[date_col_meta] >= start_datetime) & (df_meta[date_col_meta] <= end_datetime)]
                    st.sidebar.write(f"Meta: {before_count}行 → {len(df_meta)}行")
                if date_col_hs:
                    before_count = len(df_hs)
                    df_hs = df_hs[(df_hs[date_col_hs] >= start_datetime) & (df_hs[date_col_hs] <= end_datetime)]
                    st.sidebar.write(f"HubSpot: {before_count}行 → {len(df_hs)}行")

            # === デバッグ情報（期間フィルターの後） ===
            st.sidebar.markdown("---")
            st.sidebar.subheader("🔍 検出された列")
            st.sidebar.write(f"広告名: `{name_col}`")
            st.sidebar.write(f"消化金額: `{spend_col}`")
            st.sidebar.write(f"Meta日付: `{date_col_meta}`")
            st.sidebar.write(f"UTM: `{utm_col}`")

            st.sidebar.markdown("---")
            st.sidebar.subheader("デバッグ情報")
            st.sidebar.write(f"Meta広告データ: {len(df_meta)}行")
            st.sidebar.write(f"HubSpotデータ: {len(df_hs)}行")

            # === 1. データ結合キーの作成 ===
            df_meta['key'] = df_meta[name_col].astype(str).str.extract(r'(bn\d+)', expand=False)
            df_hs['key'] = df_hs[utm_col].astype(str).str.strip()

            st.sidebar.write(f"Meta（キー抽出前）: {len(df_meta)}行")
            st.sidebar.write(f"HubSpot（キー抽出前）: {len(df_hs)}行")
            
            df_meta = df_meta[df_meta['key'].notna()]
            df_hs = df_hs[df_hs['key'].notna()]
            
            st.sidebar.write(f"Meta（キー抽出後）: {len(df_meta)}行")
            st.sidebar.write(f"HubSpot（キー抽出後）: {len(df_hs)}行")
            
            st.sidebar.write("抽出されたバナーID:")
            st.sidebar.write(sorted(df_meta['key'].unique()))

            # === 2. Meta側の消化金額集計 ===
            meta_spend = df_meta.groupby('key')[spend_col].sum().reset_index()
            meta_spend[spend_col] = pd.to_numeric(meta_spend[spend_col], errors='coerce').fillna(0)

            st.sidebar.markdown("---")
            st.sidebar.write("📊 Meta消化金額（バナー別）:")
            st.sidebar.dataframe(meta_spend.rename(columns={'key': 'バナー', spend_col: '消化金額'}), use_container_width=True)
            
            total_meta_spend = meta_spend[spend_col].sum()
            st.sidebar.write(f"Meta消化金額合計: ¥{int(total_meta_spend):,}")

            # === 3. HubSpot側でリード数・接続・商談・法人をカウント ===
            hs_summary = df_hs.groupby('key').agg(
                リード数=('key', 'size')
            ).reset_index()
            
            st.sidebar.markdown("---")
            st.sidebar.write("📊 HubSpotリード数（バナー別）:")
            st.sidebar.dataframe(hs_summary, use_container_width=True)
            st.sidebar.write(f"HubSpotリード数合計: {hs_summary['リード数'].sum()}件")

            # 接続数
            call_result_col = next((c for c in hs_cols if 'コールの成果' in str(c)), None)
            if call_result_col:
                connect_df = df_hs[df_hs[call_result_col].fillna('').astype(str).str.contains('あり', case=False, na=False)]
                connect_count = connect_df.groupby('key').size().reset_index(name='接続数')
                hs_summary = pd.merge(hs_summary, connect_count, on='key', how='left')
                st.sidebar.write(f"接続列: `{call_result_col}` → {len(connect_df)}件")
            else:
                hs_summary['接続数'] = 0
                st.sidebar.warning("⚠️ 「コールの成果」列が見つかりません")

            # 商談実施数
            first_meeting_col = next((c for c in hs_cols if '初回商談日' in str(c)), None)
            second_meeting_col = next((c for c in hs_cols if '再商談日' in str(c)), None)

            deal_done_df = pd.DataFrame()

            if first_meeting_col:
                df_hs[first_meeting_col] = pd.to_datetime(df_hs[first_meeting_col], errors='coerce')
                first_deal = df_hs[df_hs[first_meeting_col].notna()][['key']].copy()
                deal_done_df = pd.concat([deal_done_df, first_deal])
                st.sidebar.write(f"初回商談日あり: {df_hs[first_meeting_col].notna().sum()}件")

            if second_meeting_col:
                df_hs[second_meeting_col] = pd.to_datetime(df_hs[second_meeting_col], errors='coerce')
                second_deal = df_hs[df_hs[second_meeting_col].notna()][['key']].copy()
                deal_done_df = pd.concat([deal_done_df, second_deal])
                st.sidebar.write(f"再商談日あり: {df_hs[second_meeting_col].notna().sum()}件")

            if len(deal_done_df) > 0:
                deal_done_count = deal_done_df.groupby('key').size().reset_index(name='商談実施数')
                hs_summary = pd.merge(hs_summary, deal_done_count, on='key', how='left')
                st.sidebar.write(f"✅ 商談実施数: {len(deal_done_df.drop_duplicates())}件")
            else:
                hs_summary['商談実施数'] = 0
                st.sidebar.warning("⚠️ 商談日付列が見つかりません")

            # 商談予約数
            stage_col = next((c for c in hs_cols if '取引ステージ' in str(c)), None)
            if stage_col:
                deal_plan_df = df_hs[df_hs[stage_col].fillna('').astype(str).str.contains('商談予定', case=False, na=False)]
                deal_plan_count = deal_plan_df.groupby('key').size().reset_index(name='商談予約数')
                hs_summary = pd.merge(hs_summary, deal_plan_count, on='key', how='left')
                st.sidebar.write(f"商談予約: {len(deal_plan_df)}件")
            else:
                hs_summary['商談予約数'] = 0
                st.sidebar.warning("⚠️ 「取引ステージ」列が見つかりません")

            # 進捗ステータス別カウント
            if stage_col:
                status_mapping = {
                    '新規リード': ['新規リード'],
                    '進捗中': ['FS対応中：社内検討', 'FS対応中：検討', 'FS対応中：見込み', 'FS対応中：申込'],
                    '商談予定': ['商談予定'],
                    'ナーチャリング': ['ナーチャリング'],
                    '保留・NG': ['低温リスト', '完全NG', 'NG対象'],
                    '契約': ['規約完了']
                }
                
                for status_name, keywords in status_mapping.items():
                    pattern = '|'.join(keywords)
                    status_df = df_hs[df_hs[stage_col].fillna('').astype(str).str.contains(pattern, case=False, na=False)]
                    status_count = status_df.groupby('key').size().reset_index(name=status_name)
                    hs_summary = pd.merge(hs_summary, status_count, on='key', how='left')
            else:
                for status_name in ['新規リード', '進捗中', '商談予定', 'ナーチャリング', '保留・NG', '契約']:
                    hs_summary[status_name] = 0

            # 法人数
            if attr_col:
                corp_df = df_hs[
                    (df_hs[attr_col].fillna('').astype(str).str.contains('法人', case=False, na=False)) &
                    (~df_hs[attr_col].fillna('').astype(str).str.contains('社員', case=False, na=False))
                ]
                corp_count = corp_df.groupby('key').size().reset_index(name='法人数')
                hs_summary = pd.merge(hs_summary, corp_count, on='key', how='left')
                st.sidebar.write(f"法人数: {len(corp_df)}件")
            else:
                hs_summary['法人数'] = 0

            hs_summary = hs_summary.fillna(0)
            hs_summary['接続数'] = hs_summary['接続数'].astype(int)
            hs_summary['商談実施数'] = hs_summary['商談実施数'].astype(int)
            hs_summary['商談予約数'] = hs_summary['商談予約数'].astype(int)
            hs_summary['法人数'] = hs_summary['法人数'].astype(int)
            
            for status_name in ['新規リード', '進捗中', '商談予定', 'ナーチャリング', '保留・NG', '契約']:
                hs_summary[status_name] = hs_summary[status_name].astype(int)

            # === 4. Meta消化金額と結合 ===
            result = pd.merge(hs_summary, meta_spend, on='key', how='outer')
            result[spend_col] = result[spend_col].fillna(0)
            result['リード数'] = result['リード数'].fillna(0).astype(int)
            
            # 進捗ステータス列もNaNを0で埋める
            for col in ['接続数', '商談実施数', '商談予約数', '法人数', '新規リード', '進捗中', '商談予定', 'ナーチャリング', '保留・NG', '契約']:
                if col in result.columns:
                    result[col] = result[col].fillna(0).astype(int)

            # === 5. 指標計算 ===
            total_spend = result[spend_col].sum()
            total_leads = result['リード数'].sum()
            total_connect = result['接続数'].sum()
            total_deal = result['商談実施数'].sum()
            total_plan = result['商談予約数'].sum()
            total_corp = result['法人数'].sum()

            result['CPA'] = result.apply(
                lambda x: int(x[spend_col] / x['リード数']) if x['リード数'] > 0 else 0, axis=1
            )
            result['接続率'] = result.apply(
                lambda x: (x['接続数'] / x['リード数'] * 100) if x['リード数'] > 0 else 0, axis=1
            )
            result['商談化率'] = result.apply(
                lambda x: ((x['商談実施数'] + x['商談予約数']) / x['リード数'] * 100) if x['リード数'] > 0 else 0, axis=1
            )
            result['法人率'] = result.apply(
                lambda x: (x['法人数'] / x['リード数'] * 100) if x['リード数'] > 0 else 0, axis=1
            )

            # === 6. 判定ロジック ===
            def judge(row):
                cpa_ok = row['CPA'] > 0 and row['CPA'] <= cpa_limit
                connect_ok = row['接続率'] >= connect_target
                meeting_ok = row['商談化率'] >= meeting_target

                conditions_met = sum([cpa_ok, connect_ok, meeting_ok])

                if conditions_met == 3:
                    return "最優秀"
                elif conditions_met == 2 and meeting_ok:
                    return "優秀"
                elif conditions_met == 2:
                    return "要改善"
                elif conditions_met == 1 and meeting_ok:
                    return "要改善"
                else:
                    return "停止推奨"

            result['判定'] = result.apply(judge, axis=1)

            # === 7. 全体サマリー KPI計算 ===
            avg_cpa = int(total_spend / total_leads) if total_leads > 0 else 0
            avg_connect = (total_connect / total_leads * 100) if total_leads > 0 else 0
            avg_meeting = ((total_deal + total_plan) / total_leads * 100) if total_leads > 0 else 0
            avg_corp = (total_corp / total_leads * 100) if total_leads > 0 else 0

            st.subheader("全体実績サマリー")

            cols_row1 = st.columns([1, 1, 1, 1])

            with cols_row1[0]:
                st.markdown(f"""
                <div style='background-color: rgb(64, 180, 200); border-radius: 12px; padding: 24px; color: white; height: 140px; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; margin-bottom: 16px;'>
                    <div style='font-size: 0.85rem; font-weight: 400; opacity: 0.95; margin-bottom: 12px;'>総消化金額</div>
                    <div style='font-size: 1.6rem; font-weight: 700; line-height: 1.2;'>¥{int(total_spend):,}</div>
                </div>
                """, unsafe_allow_html=True)

            with cols_row1[1]:
                st.markdown(f"""
                <div style='background-color: rgb(64, 180, 200); border-radius: 12px; padding: 24px; color: white; height: 140px; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; margin-bottom: 16px;'>
                    <div style='font-size: 0.85rem; font-weight: 400; opacity: 0.95; margin-bottom: 12px;'>総リード数</div>
                    <div style='font-size: 1.6rem; font-weight: 700; line-height: 1.2;'>{int(total_leads)}件</div>
                </div>
                """, unsafe_allow_html=True)

            with cols_row1[2]:
                st.markdown(f"""
                <div style='background-color: rgb(64, 180, 200); border-radius: 12px; padding: 24px; color: white; height: 140px; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; margin-bottom: 16px;'>
                    <div style='font-size: 0.85rem; font-weight: 400; opacity: 0.95; margin-bottom: 8px;'>接続数</div>
                    <div style='font-size: 1.6rem; font-weight: 700; line-height: 1.2; margin-bottom: 4px;'>{int(total_connect)}件</div>
                    <div style='font-size: 0.75rem; font-weight: 400; opacity: 0.9;'>接続率 {avg_connect:.1f}%</div>
                </div>
                """, unsafe_allow_html=True)

            with cols_row1[3]:
                st.markdown(f"""
                <div style='background-color: rgb(64, 180, 200); border-radius: 12px; padding: 24px; color: white; height: 140px; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; margin-bottom: 16px;'>
                    <div style='font-size: 0.85rem; font-weight: 400; opacity: 0.95; margin-bottom: 12px;'>平均CPA</div>
                    <div style='font-size: 1.6rem; font-weight: 700; line-height: 1.2;'>¥{avg_cpa:,}</div>
                </div>
                """, unsafe_allow_html=True)

            cols_row2 = st.columns([1, 1, 1, 1])

            with cols_row2[0]:
                st.markdown(f"""
                <div style='background-color: rgb(64, 180, 200); border-radius: 12px; padding: 24px; color: white; height: 140px; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center;'>
                    <div style='font-size: 0.85rem; font-weight: 400; opacity: 0.95; margin-bottom: 12px;'>商談実施数</div>
                    <div style='font-size: 1.6rem; font-weight: 700; line-height: 1.2;'>{int(total_deal)}件</div>
                </div>
                """, unsafe_allow_html=True)

            with cols_row2[1]:
                st.markdown(f"""
                <div style='background-color: rgb(64, 180, 200); border-radius: 12px; padding: 24px; color: white; height: 140px; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center;'>
                    <div style='font-size: 0.85rem; font-weight: 400; opacity: 0.95; margin-bottom: 8px;'>商談予約数</div>
                    <div style='font-size: 1.6rem; font-weight: 700; line-height: 1.2; margin-bottom: 4px;'>{int(total_plan)}件</div>
                    <div style='font-size: 0.75rem; font-weight: 400; opacity: 0.9;'>商談化率 {avg_meeting:.1f}%</div>
                </div>
                """, unsafe_allow_html=True)

            with cols_row2[2]:
                st.markdown(f"""
                <div style='background-color: rgb(64, 180, 200); border-radius: 12px; padding: 24px; color: white; height: 140px; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center;'>
                    <div style='font-size: 0.85rem; font-weight: 400; opacity: 0.95; margin-bottom: 8px;'>法人数</div>
                    <div style='font-size: 1.6rem; font-weight: 700; line-height: 1.2; margin-bottom: 4px;'>{int(total_corp)}件</div>
                    <div style='font-size: 0.75rem; font-weight: 400; opacity: 0.9;'>法人率 {avg_corp:.1f}%</div>
                </div>
                """, unsafe_allow_html=True)

            with cols_row2[3]:
                st.markdown("""
                <div style='height: 140px;'></div>
                """, unsafe_allow_html=True)

            st.markdown("---")

            summary_data = {
                'ファイル名': meta_file.name + ' & ' + hs_file.name,
                '総リード数': total_leads,
                '平均CPA': avg_cpa,
                '総消化金額': int(total_spend),
                '商談化率': avg_meeting
            }

            if st.button("✅ KPI分析結果をスプレッドシートに反映！", help="このボタンで分析結果が全員共有のスプレッドシートに記録されます。"):
                write_analysis_to_sheet(summary_data, SPREADSHEET_URL, KPI_SHEET_INDEX)

            st.markdown("---")

            # === 8. バナー別評価表 ===
            st.subheader("バナー別 評価表")

            display_df = result.copy()
            display_df = display_df.rename(columns={
                'key': 'バナーID',
                spend_col: '消化金額'
            })

            display_df['消化金額_表示'] = display_df['消化金額'].apply(lambda x: f"{int(x):,}")
            display_df['CPA_表示'] = display_df['CPA'].apply(lambda x: f"{int(x):,}")
            display_df['接続率_表示'] = display_df['接続率'].apply(lambda x: f"{x:.1f}%")
            display_df['商談化率_表示'] = display_df['商談化率'].apply(lambda x: f"{x:.1f}%")
            display_df['法人率_表示'] = display_df['法人率'].apply(lambda x: f"{x:.1f}%")

            show_df = display_df[['判定', 'バナーID', '消化金額_表示', 'リード数', 'CPA_表示', '接続率_表示', '商談化率_表示', '法人率_表示', '接続数', '商談実施数', '商談予約数', '法人数']].copy()
            show_df.columns = ['判定', 'バナーID', '消化金額', 'リード数', 'CPA', '接続率', '商談化率', '法人率', '接続数', '商談実施数', '商談予約数', '法人数']

            judgment_order = {"最優秀": 0, "優秀": 1, "要改善": 2, "停止推奨": 3}

            show_df['判定_rank'] = show_df['判定'].map(judgment_order)
            show_df['バナーID_num'] = show_df['バナーID'].str.extract(r'(\d+)').astype(int)

            show_df = show_df.sort_values(by=['判定_rank', 'バナーID_num'], ascending=[True, False])
            show_df = show_df.drop(columns=['判定_rank', 'バナーID_num'])

            def highlight_row(row):
                判定 = row['判定']
                if 判定 == "最優秀":
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

            # === 9. バナー別進捗状況テーブル ===
            st.markdown("---")
            st.subheader("バナー別 進捗状況")

            progress_df = display_df[['バナーID', 'リード数', '新規リード', '進捗中', '商談予定', 'ナーチャリング', '保留・NG', '契約']].copy()
            
            # リード数が0のバナーを除外
            progress_df = progress_df[progress_df['リード数'] > 0].drop(columns=['リード数'])
            
            progress_df['バナーID_num'] = progress_df['バナーID'].str.extract(r'(\d+)').astype(int)
            progress_df = progress_df.sort_values(by=['バナーID_num'], ascending=[False])
            progress_df = progress_df.drop(columns=['バナーID_num'])

            # 0を空白に置換
            progress_df_display = progress_df.fillna(0).replace(0, '').replace(0.0, '')

            st.dataframe(
                progress_df_display, 
                use_container_width=True,
                hide_index=True
            )

            # === 10. アクション提案 ===
            st.markdown("---")
            st.subheader("推奨アクション")

            best = result[result['判定'] == "最優秀"]['key'].tolist()
            good = result[result['判定'].str.contains("優秀", na=False)]['key'].tolist()
            improve = result[result['判定'].str.contains("要改善", na=False)]['key'].tolist()
            stop = result[result['判定'] == "停止推奨"]['key'].tolist()

            if best:
                st.success(f"【予算集中】 {', '.join(best)} → CPA・接続率・商談化率すべて基準クリア")
            if good:
                good_filtered = [b for b in good if b not in best]
                if good_filtered:
                    st.info(f"【有望株】 {', '.join(good_filtered)} → 商談化率は目標達成。CPA or 接続率を改善すれば最優秀に")
            if improve:
                st.warning(f"【要分析】 {', '.join(improve)} → LP改善や接続体制の見直しを検討")
            if stop:
                st.error(f"【停止検討】 {', '.join(stop)} → 予算を優秀バナーに振り替え")

            # === 11. 分布図 ===
            st.markdown("---")
            st.subheader("バナー別 パフォーマンス分布")

            chart_data = result[result['リード数'] > 0].copy()
            if len(chart_data) > 0:
                chart = alt.Chart(chart_data).mark_circle(size=200).encode(
                    x=alt.X('CPA:Q', title='CPA (円)', scale=alt.Scale(zero=False)),
                    y=alt.Y('商談化率:Q', title='商談化率 (%)'),
                    color=alt.Color('判定:N', legend=alt.Legend(title="判定"), scale=alt.Scale(
                        domain=['最優秀', '優秀', '要改善', '停止推奨'],
                        range=['#28a745', '#17a2b8', '#ffc107', '#dc3545']
                    )),
                    size=alt.Size('リード数:Q', legend=None),
                    tooltip=['key', 'CPA', '接続率', '商談化率', '法人率', 'リード数', '判定']
                ).properties(height=450).interactive()
                st.altair_chart(chart, use_container_width=True)

        except Exception as e:
            st.error(f"処理エラー: {e}")
            import traceback
            st.code(traceback.format_exc())

else:
    st.info("Meta広告実績とHubSpotデータをアップロードしてください")
