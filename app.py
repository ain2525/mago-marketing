import streamlit as st
import pandas as pd
import altair as alt
import json
import gspread
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from PIL import Image

# =========================================================
# 設定
# =========================================================

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1dJwYYK-koOgU0V9z83hfz-Wjjjl_UNbl_N6eHQk5OmI/edit"
KPI_SHEET_INDEX = 1  # 2枚目のシート

# =========================================================
# Google Sheets 書き込み処理
# =========================================================

def write_analysis_to_sheet(analysis_data, spreadsheet_url, sheet_index):
    status = st.sidebar.empty()
    status.info("スプレッドシートへ接続中...")

    try:
        service_account_info = json.loads(st.secrets["google_sheets"])
        client = gspread.service_account_from_dict(service_account_info)

        workbook = client.open_by_url(spreadsheet_url)
        sheet = workbook.get_worksheet(sheet_index)

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        row = [
            now,
            analysis_data.get("ファイル名", "N/A"),
            analysis_data.get("総リード数", 0),
            analysis_data.get("平均CPA", 0),
            analysis_data.get("総消化金額", 0),
            f'{analysis_data.get("商談化率", 0):.1f}%'
        ]

        sheet.append_row(row)
        status.success("✅ KPIをスプレッドシートに反映しました")

    except Exception as e:
        status.error(f"❌ 書き込み失敗: {e}")

# =========================================================
# アプリ設定
# =========================================================

st.set_page_config(page_title="Meta広告×セールスダッシュボード", layout="wide")

try:
    logo = Image.open("logo.png")
    st.image(logo, use_container_width=True)
except:
    pass

st.title("Meta広告 × セールスダッシュボード")
st.markdown("---")

# =========================================================
# ファイル読み込み
# =========================================================

def load_data(file):
    if file.name.endswith(".csv"):
        try:
            return pd.read_csv(file)
        except:
            return pd.read_csv(file, encoding="shift-jis")
    else:
        return pd.read_excel(file)

# =========================================================
# サイドバー
# =========================================================

st.sidebar.header("判定基準")
cpa_limit = st.sidebar.number_input("許容CPA", value=10000, step=1000)
connect_target = st.sidebar.slider("目標接続率 (%)", 0, 100, 50)
meeting_target = st.sidebar.slider("目標商談化率 (%)", 0, 50, 18)

# =========================================================
# ファイルアップロード
# =========================================================

col1, col2 = st.columns(2)
with col1:
    meta_file = st.file_uploader("Meta広告実績", type=["csv", "xlsx"])
with col2:
    hs_file = st.file_uploader("HubSpotデータ", type=["csv", "xlsx"])

st.markdown("---")

# =========================================================
# メイン処理
# =========================================================

if meta_file and hs_file:
    try:
        df_meta = load_data(meta_file)
        df_hs = load_data(hs_file)

        meta_cols = df_meta.columns.astype(str)
        hs_cols = df_hs.columns.astype(str)

        name_col = next(c for c in meta_cols if "名前" in c or "Name" in c)
        spend_col = next(c for c in meta_cols if "消化" in c or "Amount" in c)

        utm_col = next(c for c in hs_cols if "UTM" in c)
        connect_col = next((c for c in hs_cols if "接続" in c), None)
        deal_col = next((c for c in hs_cols if "商談" in c), None)

        df_meta["key"] = df_meta[name_col].astype(str).str.extract("(bn\\d+)")
        df_hs["key"] = df_hs[utm_col].astype(str)

        df_meta = df_meta.dropna(subset=["key"])
        df_hs = df_hs.dropna(subset=["key"])

        meta_spend = df_meta.groupby("key")[spend_col].sum().reset_index()
        meta_spend[spend_col] = pd.to_numeric(meta_spend[spend_col], errors="coerce").fillna(0)

        hs_count = df_hs.groupby("key").size().reset_index(name="リード数")

        if connect_col:
            connect_df = df_hs[df_hs[connect_col].astype(str).str.contains("あり|済|true|yes", case=False, na=False)]
            connect_count = connect_df.groupby("key").size().reset_index(name="接続数")
            hs_count = hs_count.merge(connect_count, on="key", how="left")
        else:
            hs_count["接続数"] = 0

        if deal_col:
            deal_df = df_hs[df_hs[deal_col].astype(str).str.contains("あり|済|完了", case=False, na=False)]
            deal_count = deal_df.groupby("key").size().reset_index(name="商談数")
            hs_count = hs_count.merge(deal_count, on="key", how="left")
        else:
            hs_count["商談数"] = 0

        hs_count = hs_count.fillna(0)

        result = hs_count.merge(meta_spend, on="key", how="left")
        result[spend_col] = result[spend_col].fillna(0)

        result["CPA"] = (result[spend_col] / result["リード数"]).fillna(0).astype(int)
        result["接続率"] = result["接続数"] / result["リード数"] * 100
        result["商談化率"] = result["商談数"] / result["リード数"] * 100

        total_spend = int(result[spend_col].sum())
        total_leads = int(result["リード数"].sum())
        avg_cpa = int(total_spend / total_leads) if total_leads > 0 else 0
        avg_meeting = (result["商談数"].sum() / total_leads * 100) if total_leads > 0 else 0

        st.subheader("全体サマリー")
        st.metric("総消化金額", f"¥{total_spend:,}")
        st.metric("総リード数", total_leads)
        st.metric("平均CPA", f"¥{avg_cpa:,}")
        st.metric("商談化率", f"{avg_meeting:.1f}%")

        summary = {
            "ファイル名": f"{meta_file.name} & {hs_file.name}",
            "総リード数": total_leads,
            "平均CPA": avg_cpa,
            "総消化金額": total_spend,
            "商談化率": avg_meeting
        }

        if st.button("✅ KPI分析結果をスプレッドシートに反映！"):
            write_analysis_to_sheet(summary, SPREADSHEET_URL, KPI_SHEET_INDEX)

        st.markdown("---")
        st.subheader("バナー別結果")
        st.dataframe(result, use_container_width=True)

    except Exception as e:
        st.error(e)
        import traceback
        st.code(traceback.format_exc())

else:
    st.info("Meta広告実績とHubSpotデータをアップロードしてください")
