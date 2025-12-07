import streamlit as st
import pandas as pd
import altair as alt
import os # ★追加: 環境変数 (Secrets) を読み込むため
import json # ★追加: JSON文字列をPythonオブジェクトに変換するため
import gspread # ★追加: Google Sheetsにアクセスするため
from google.oauth2.service_account import Credentials # ★追加: サービスアカウント認証のため
from datetime import datetime, timedelta
from PIL import Image

# =========================================================================
# 【１】スプレッドシート書き込み関数 (完全版)
# =========================================================================

# 【組み込み済み】提供されたスプレッドシートURL
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1dJwYYK-koOgU0V9z83hfz-Wjjjl_UNbl_N6eHQk5OmI/edit"
# 【設定】書き込み先のシートインデックス (1: 2枚目のシートを意味。1枚目の場合は0にしてください)
KPI_SHEET_INDEX = 1 

def write_analysis_to_sheet(analysis_data, spreadsheet_url, sheet_index):
    """
    計算されたKPI分析結果をGoogleスプレッドシートに書き込む関数
    """
    
    # メッセージ表示用のコンテナ
    status_container = st.sidebar.empty() 
    status_container.info("スプレッドシートへの接続準備中...")
    
    try:
        # 認証処理（GitHub Secretsから認証情報を取得）
        creds_json = os.environ.get("GOOGLE_SHEETS_CREDENTIALS")
        if not creds_json:
            status_container.error("❌ 認証情報が見つかりません。GitHub Secretsを確認してください。")
            return

        creds_dict = json.loads(creds_json)
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        
        # スプレッドシートを開く
        workbook = client.open_by_url(spreadsheet_url)
        sheet = workbook.get_worksheet(sheet_index) 

        # データをリスト形式に整形 (ヘッダー順: 記録日時, ファイル名, 総リード数, 平均CPA, 総消化金額, 商談化率)
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        data_to_write = [
            now,
            analysis_data.get('ファイル名', 'N/A'),
            analysis_data.get('総リード数', 'N/A'),
            analysis_data.get('平均CPA', 'N/A'),
            analysis_data.get('総消化金額', 'N/A'),
            analysis_data.get('商談化率', 'N/A')
        ]
        
        sheet.append_row(data_to_write)
        status_container.success("✅ 分析結果をKPIサマリーシートに反映しました！")
        
    except Exception as e:
        status_container.error(f"❌ 書き込み失敗。エラー: {e}")
        # デバッグのために詳細ログを出力
        st.code(f"書き込みエラー詳細: {e}")

# =========================================================================
# 【２】既存のアプリメイン処理
# =========================================================================

# --- ページ設定 ---
st.set_page_config(page_title="Meta広告×セールスダッシュボード", layout="wide")

# --- ロゴ表示（既存コード） ---
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

# --- データ読み込み関数（既存コード） ---
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

# --- サイドバー設定（既存コード） ---
st.sidebar.header("判定基準の設定")
cpa_limit = st.sidebar.number_input("許容CPA（円）", value=10000, step=1000)
connect_target = st.sidebar.slider("目標接続率（%）", 0, 100, 50)
meeting_target = st.sidebar.slider("目標商談化率（%）", 0, 50, 18)

st.sidebar.markdown("---")
st.sidebar.subheader("分析期間の設定")

# --- ファイルアップロード（既存コード） ---
col1, col2 = st.columns(2)
with col1:
    meta_file = st.file_uploader("Meta広告実績", type=['xlsx', 'csv'])
with col2:
    hs_file = st.file_uploader("HubSpotデータ", type=['xlsx', 'csv'])

st.markdown("---")

# --- 分析実行（既存コード） ---
if meta_file and hs_file:
    df_meta = load_data(meta_file)
    df_hs = load_data(hs_file)

    if df_meta is not None and df_hs is not None:
        try:
            # === Meta側：列の特定 === (中略)
            meta_cols = list(df_meta.columns)
            name_col = next((c for c in meta_cols if '名前' in str(c) or 'Name' in str(c)), None)
            spend_col = next((c for c in meta_cols if '消化金額' in str(c) or 'Amount' in str(c) or '費用' in str(c)), None)
            date_col_meta = next((c for c in meta_cols if '日' in str(c) or 'Date' in str(c) or '開始' in str(c)), None)

            # === HubSpot側：列の特定 === (中略)
            hs_cols = list(df_hs.columns)
            utm_col = next((c for c in hs_cols if 'UTM' in str(c) or 'Content' in str(c)), None)
            connect_col = next((c for c in hs_cols if '接続' in str(c)), None)
            deal_col = next((c for c in hs_cols if '商談' in str(c)), None)
            attr_col = next((c for c in hs_cols if '属性' in str(c)), None)
            date_col_hs = next((c for c in hs_cols if '作成日' in str(c) or 'Created' in str(c) or '日付' in str(c)), None)

            if not all([name_col, spend_col, utm_col]):
                st.error(f"必要な列が見つかりません。\nMeta: 広告名={name_col}, 消化金額={spend_col}\nHubSpot: UTM={utm_col}")
                st.stop()

            # === 日付列の変換 & 期間フィルター & デバッグ情報 === (中略)

            # --- 既存の分析ロジックの実行開始 ---
            # ... (既存のコード: データ結合、集計、判定ロジックなど) ...

            # === 1. データ結合キーの作成 ===
            df_meta['key'] = df_meta[name_col].astype(str).str.extract(r'(bn\d+)', expand=False)
            df_hs['key'] = df_hs[utm_col].astype(str).str.strip()
            df_meta = df_meta[df_meta['key'].notna()]
            df_hs = df_hs[df_hs['key'].notna()]

            # === 2. Meta側の消化金額集計 ===
            meta_spend = df_meta.groupby('key')[spend_col].sum().reset_index()
            meta_spend[spend_col] = pd.to_numeric(meta_spend[spend_col], errors='coerce').fillna(0)
            
            # ... (中略: サイドバーへの Meta消化金額表示) ...

            # === 3. HubSpot側でリード数・接続・商談・法人をカウント ===
            hs_summary = df_hs.groupby('key').agg(リード数=('key', 'size')).reset_index()

            # ... (中略: 接続数、商談実施数・予約数、法人数 の計算と結合) ...
            
            # === 4. Meta消化金額と結合 ===
            result = pd.merge(hs_summary, meta_spend, on='key', how='left')
            result[spend_col] = result[spend_col].fillna(0)

            # === 5. 指標計算 ===
            total_spend = result[spend_col].sum() # ★この行を移動・追加 (全体サマリー計算前に必要)
            total_leads = result['リード数'].sum() # ★この行を移動・追加 (全体サマリー計算前に必要)
            total_connect = result['接続数'].sum() # ★この行を移動・追加
            total_deal = result['商談実施数'].sum() # ★この行を移動・追加
            total_plan = result['商談予約数'].sum() # ★この行を移動・追加

            result['CPA'] = result.apply(
                lambda x: int(x[spend_col] / x['リード数']) if x['リード数'] > 0 else 0, axis=1
            )
            result['接続率'] = result.apply(
                lambda x: (x['接続数'] / x['リード数'] * 100) if x['リード数'] > 0 else 0, axis=1
            )
            result['商談化率'] = result.apply(
                lambda x: ((x['商談実施数'] + x['商談予約数']) / x['リード数'] * 100) if x['リード数'] > 0 else 0, axis=1
            )

            # === 7. 全体サマリー === (中略)

            # ... (既存のコード: 全体サマリーのHTML表示ロジック) ...
            
            # --- ここから書き込みトリガー ---
            
            # === 10. KPIを辞書にまとめる ===
            avg_cpa = int(total_spend / total_leads) if total_leads > 0 else 0
            avg_meeting = ((total_deal + total_plan) / total_leads * 100) if total_leads > 0 else 0

            summary_data = {
                'ファイル名': meta_file.name + ' & ' + hs_file.name,
                '総リード数': total_leads,
                '平均CPA': avg_cpa,
                '総課金化金額': int(total_spend),
                '商談化率': avg_meeting
            }
            
            st.markdown("---")
            
            # === 11. 反映ボタンの設置 ===
            if st.button("✅ KPI分析結果をスプレッドシートに反映！", help="このボタンを押すと、上のサマリー結果がKPIシートに記録されます。"):
                write_analysis_to_sheet(summary_data, SPREADSHEET_URL, KPI_SHEET_INDEX)
            
            # ... (既存のコード: バナー別評価表、アクション提案、分布図ロジック) ...


        except Exception as e:
            st.error(f"処理エラー: {e}")
            import traceback
            st.code(traceback.format_exc())

else:
    st.info("Meta広告実績とHubSpotデータをアップロードしてください")
