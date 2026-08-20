import streamlit as st
import pandas as pd
import datetime
import re
import io
import requests
from streamlit_autorefresh import st_autorefresh

st.set_page_config(layout="wide", page_title="OnAir Monitor")

# Auto-refresh every 60 seconds
st_autorefresh(interval=60 * 1000, key="datarefresh")

# 日本時間 (JST: UTC+9) のタイムゾーン定義
JST = datetime.timezone(datetime.timedelta(hours=9))

def get_jst_now():
    """日本時間の現在日時を取得"""
    return datetime.datetime.now(JST)

# Constants for Excel Columns
COL_TRIGGER = 2
COL_TIME = 3
COL_MATERIAL = 4
COL_PROGRAM_G = 6
COL_PROGRAM_H = 7
COL_PROGRAM_J = 9

HEADER_ROW_INDEX = 4
MATERIAL_LIST = ["S1", "S2", "S3", "D1", "D2", "D3", "P2", "X5", "R3", "R4", "R5", "R6", "OS", "FL"]

def fetch_excel_from_gas(target_date):
    """GASからファイルIDを取得し、Google Driveから直接ダウンロードする"""
    date_str = target_date.strftime('%Y%m%d')
    try:
        gas_url = st.secrets.get("gas_api_url", "")
        if not gas_url or not gas_url.startswith("http"):
            st.error("【設定エラー】Secretsに 'gas_api_url' が正しく設定されていません。")
            return None
            
        # 1. GASにアクセスしてファイルIDを要求する
        response = requests.get(gas_url, params={"date": date_str}, timeout=20, allow_redirects=True)
        
        if response.status_code != 200:
            st.error(f"【通信エラー】GASへのアクセスに失敗しました (Status: {response.status_code})")
            return None
            
        text = response.text.strip()
        
        if text.startswith("GAS_INTERNAL_ERROR:"):
            st.error(f"【GAS内部エラー】{text}")
            return None
        elif text == "File Not Found":
            st.warning(f"【見つかりません】{date_str} の進行表がGoogle Driveにありません。")
            return None
        elif text == "Error: Date parameter missing":
            st.error("【GAS応答】日付パラメータが不足しています。")
            return None
            
        # 2. ファイルIDを正しく受け取れた場合、Google Driveから直接ダウンロード！
        if text.startswith("FILE_ID:"):
            file_id = text.split("FILE_ID:")[1]
            
            # ダイレクトダウンロード用URL
            download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
            dl_response = requests.get(download_url, timeout=20)
            
            if dl_response.status_code == 200:
                # 成功！バイナリデータを返す
                return io.BytesIO(dl_response.content)
            else:
                st.error(f"【Drive取得エラー】ファイルのダウンロードに失敗しました (Status: {dl_response.status_code})")
                return None
        else:
            # HTMLが返ってきてしまった場合（ボット弾き）
            st.error("【Googleブロック】アクセスが弾かれました。Driveフォルダの権限が「リンクを知っている全員」になっているか確認してください。")
            with st.expander("詳細エラー内容(HTML)を開く"):
                st.code(text[:1000])
            return None

    except Exception as e:
        st.error(f"【取得エラー】({date_str}): {e}")
        return None

@st.cache_data(ttl=60)
def load_data_for_date(target_date):
    file_buffer = fetch_excel_from_gas(target_date)
    if not file_buffer:
        return pd.DataFrame()
        
    try:
        xls = pd.ExcelFile(file_buffer)
        sheet_name = "放送進行表"
        if sheet_name not in xls.sheet_names:
            sheet_name = xls.sheet_names[0]
            
        df_raw = pd.read_excel(file_buffer, sheet_name=sheet_name, header=None)
        
        data_rows = []
        last_processed_program = None
        
        for i in range(HEADER_ROW_INDEX + 1, len(df_raw)):
            row = df_raw.iloc[i]
            if len(row) < 8: continue
            
            val_col4 = str(row[COL_MATERIAL]).strip() if pd.notna(row[COL_MATERIAL]) else ""
            if val_col4 == "": continue
            
            original_time = str(row[COL_TIME]).strip() if pd.notna(row[COL_TIME]) else ""
            material = val_col4
            program = str(row[COL_PROGRAM_H]).strip() if pd.notna(row[COL_PROGRAM_H]) else ""
            
            if original_time == "05:00:00" and material == "P2": continue
            if material not in MATERIAL_LIST: continue
            if material in ["OS", "FL"]: material = "放送休止"
            
            trigger_u = str(row[COL_TRIGGER]).strip() if pd.notna(row[COL_TRIGGER]) else ""
            if trigger_u == 'U':
                last_processed_program = None
                
            if program != "" and program == last_processed_program:
                continue

            if program != "":
                last_processed_program = program

            program_display_val = ""
            if len(row) > COL_PROGRAM_J:
                program_display_val = str(row[COL_PROGRAM_J]).strip() if pd.notna(row[COL_PROGRAM_J]) else ""
                
            if program_display_val == "":
                 if len(row) > COL_PROGRAM_G:
                     program_display_val = str(row[COL_PROGRAM_G]).strip() if pd.notna(row[COL_PROGRAM_G]) else ""
            
            if program_display_val == "":
                program_display_val = program
                
            match = re.search(r'(?:<<|≪|＜＜)\s*D[1-3]\s+(.*?)\s+スタート\s*(?:>>|≫|＞＞)', program_display_val)
            if match:
                program_display_val = match.group(1)
            
            program_display_val = program_display_val.replace('<', '').replace('>', '').replace('≪', '').replace('≫', '')
            
            try:
                if isinstance(row[COL_TIME], datetime.time):
                    t = row[COL_TIME]
                elif isinstance(row[COL_TIME], datetime.datetime):
                    t = row[COL_TIME].time()
                else:
                    t_str = str(original_time).split('.')[0]
                    parts = t_str.split(':')
                    
                    if len(parts) >= 2:
                        h = int(parts[0])
                        m = int(parts[1])
                        s = int(parts[2]) if len(parts) > 2 else 0
                        
                        if h >= 24:
                            h = h % 24
                        t = datetime.time(h, m, s)
                    else:
                        t = datetime.datetime.strptime(t_str, "%H:%M:%S").time()
            except (ValueError, AttributeError, TypeError, IndexError):
                t = datetime.time(0, 0)

            full_dt = datetime.datetime.combine(target_date, t)
            
            if t < datetime.time(5, 0):
                full_dt += datetime.timedelta(days=1)
            
            display_hour = t.hour
            if t.hour < 5:
                display_hour += 24
            time_str_val = f"{display_hour:02d}:{t.minute:02d}"
            
            data_rows.append({
                'datetime': full_dt,
                'time_str': time_str_val, 
                'material': material,
                'program': program,
                'program_display': program_display_val
            })
            
        return pd.DataFrame(data_rows)

    except Exception as e:
        st.error(f"【Excel解析エラー】({target_date}): {e}")
        return pd.DataFrame()

def main():
    st.markdown("""
        <style>
        .block-container {
            padding-top: 0rem !important;
            padding-bottom: 0rem !important;
            padding-left: 1rem !important;
            padding-right: 1rem !important;
            max-width: 100%;
        }
        #MainMenu {display: none !important;}
        footer {display: none !important;}
        header {display: none !important;}
        
        @keyframes pulse-glow {
            0% { text-shadow: 0 0 10px #00FFFF, 0 0 20px #00FFFF; transform: translate(-50%, -50%) scale(1); }
            50% { text-shadow: 0 0 20px #00FFFF, 0 0 40px #00FFFF, 0 0 60px #00FFFF; transform: translate(-50%, -50%) scale(1.02); }
            100% { text-shadow: 0 0 10px #00FFFF, 0 0 20px #00FFFF; transform: translate(-50%, -50%) scale(1); }
        }
        
        .pulsing-header {
            animation: pulse-glow 2s infinite ease-in-out;
        }
        
        @keyframes pulse-frame {
            0% { box-shadow: inset 0 0 10px #FF0000, 0 0 10px #FF0000; border: 2px solid #FF0000; }
            50% { box-shadow: inset 0 0 30px #FF0000, 0 0 20px #FF0000; border: 2px solid #FF0000; }
            100% { box-shadow: inset 0 0 10px #FF0000, 0 0 10px #FF0000; border: 2px solid #FF0000; }
        }
        
        .pulsing-frame {
            animation: pulse-frame 2s infinite ease-in-out;
            z-index: 10;
            position: relative;
        }
        
        @keyframes scroll-vertical {
            0% { transform: translateY(0); }
            100% { transform: translateY(-50%); }
        }
        
        .scrolling-text {
            animation: scroll-vertical 15s linear infinite;
            white-space: nowrap;
            display: inline-block;
        }
        
        .scrolling-content-wrapper {
            display: flex; 
            flex-direction: row;
            align-items: center;
        }

        .header-time-container {
            position: absolute;
            top: 50%;
            left: 20px;
            transform: translateY(-50%);
            font-family: 'Meiryo', sans-serif;
            font-weight: bold;
            color: #E0E0E0;
            text-shadow: 0 0 5px rgba(255,255,255,0.5);
            font-size: 16px;
        }

        .header-title-container {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            font-family: 'Meiryo', sans-serif;
            font-weight: bold;
            color: #FFFFFF;
            white-space: nowrap;
            font-size: 18px;
            width: 80%;
            text-align: center;
        }

        .header-logo-text {
            color: #FF8C00;
            margin-right: 10px;
        }

        @media (min-width: 1000px) {
            .header-time-container { font-size: 28px; }
            .header-title-container { font-size: 32px; width: auto; }
            .header-logo-text { margin-right: 30px; }
        }
        </style>
    """, unsafe_allow_html=True)

    now_dt = get_jst_now()
    
    if now_dt.hour < 5:
        display_date = now_dt.date() - datetime.timedelta(days=1)
        display_hour = now_dt.hour + 24
    else:
        display_date = now_dt.date()
        display_hour = now_dt.hour
        
    weekdays = ["月", "火", "水", "木", "金", "土", "日"]
    weekday_str = weekdays[display_date.weekday()]
    
    now_display = f"{display_date.strftime('%Y.%m.%d')} ({weekday_str}) {display_hour:02}:{now_dt.minute:02}"
    st.markdown(f"""
    <div style="position: fixed; top: 0; left: 0; width: 100%; height: 60px; z-index: 999999; background: linear-gradient(90deg, #0f2027, #203a43, #2c5364); border-bottom: 2px solid #00FFFF; display: block; box-sizing: border-box; box-shadow: 0 4px 15px rgba(0,255,255, 0.3);">
        <div class="header-time-container">
            {now_display}
        </div>
        <div class="pulsing-header header-title-container">
            <span class="header-logo-text">FM NORTH WAVE</span>On Air Monitor
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    today = now_dt.date()
    tomorrow = today + datetime.timedelta(days=1)
    
    df1 = load_data_for_date(today)
    df2 = load_data_for_date(tomorrow)
    
    if df1.empty and df2.empty:
        st.warning("進行表ファイルが見つかりません。")
        return

    df = pd.concat([df1, df2], ignore_index=True)
    if df.empty:
        st.warning("データが空です。")
        return
        
    df = df.sort_values('datetime')
    
    now_naive = now_dt.replace(tzinfo=None)
    candidates = df[df['datetime'] <= now_naive]
    start_index = 0
    if not candidates.empty:
        start_index = candidates.index[-1]
    
    df_display = df.iloc[start_index:].copy()
    
    color_map = {
        "D1": "#C0C0C0", "D2": "#00FFFF", "D3": "#FFFF00", "P2": "#B8860B", "S1": "#00b31e",
        "S2": "#FF8000", "S3": "#FF00FF", "X5": "#4B0082", "放送休止": "#000033"
    }
    text_color_map = {
        "X5": "white", "放送休止": "white", "P2": "white"
    }
    
    html_content = '<div style="display: flex; overflow-x: auto; white-space: nowrap; height: 95vh; width: 100%; gap: 0px; padding-top: 55px; margin-margin-top: -40px; box-sizing: border-box;">'
    
    last_date_display = None
    
    for i, (_, row) in enumerate(df_display.iterrows()):
        mat = row['material']
        bg = color_map.get(mat, "#FFFFFF")
        txt = text_color_map.get(mat, "black")
        
        width = "80px" 
        if mat == "放送休止":
            width = "160px"
            
        time_str = row['time_str']
        program_text = row['program_display']
        
        current_dt = row['datetime']
        check_time = current_dt.time()
        broadcast_date = current_dt.date()
        if check_time < datetime.time(5, 0):
             broadcast_date -= datetime.timedelta(days=1)
        
        mat_justify = "flex-end"
        if mat == "P2":
            mat_justify = "flex-start"
        if mat in ["S1", "S2", "S3"]:
            mat_justify = "center"
        
        pulse_class = "pulsing-frame" if i == 0 else ""
        border_style = "" if i == 0 else "border-right: 1px solid #404040;"
        
        is_new_day = False
        if last_date_display is not None and broadcast_date != last_date_display:
            is_new_day = True
            
        if is_new_day:
             border_style += " border-left: 6px double #FFFF00;"
        
        last_date_display = broadcast_date
        
        if i == 0:
             program_html = f"""
        <div style="flex-grow: 1; overflow: hidden; position: relative; writing-mode: vertical-rl; text-orientation: mixed; font-family: 'Meiryo', sans-serif; font-weight: bold; font-size: 28px; box-sizing: border-box; display: flex; align-items: center; justify-content: center;">
             <div class="scrolling-text scrolling-content-wrapper">
                <div style="padding-left: 20px; padding-right: 20px; padding-top: 50px; padding-bottom: 50px;">
                    {program_text}
                </div>
                <div style="padding-left: 20px; padding-right: 20px; padding-top: 50px; padding-bottom: 50px;">
                    {program_text}
                </div>
             </div>
        </div>"""
        else:
             program_html = f"""
        <div style="flex-grow: 1; writing-mode: vertical-rl; text-orientation: mixed; display: flex; align-items: center; justify-content: center; padding-top: 10px; padding-bottom: 10px; font-weight: bold; font-size: 28px; white-space: nowrap; overflow: hidden; font-family: 'Meiryo', sans-serif; box-sizing: border-box;">{program_text}</div>"""
            
        card_html = f"""
<div style="flex: 0 0 {width}; min-width: {width}; max-width: {width}; position: relative; z-index: 0; background-color: {bg}; color: {txt}; display: flex; flex-direction: column; align-items: stretch; height: 100%; box-sizing: border-box;">
    <div class="{pulse_class}" style="flex-grow: 1; display: flex; flex-direction: column; {border_style} box-sizing: border-box;">
        <div style="height: 45px; width: 100%; border-bottom: 1px solid rgba(0,0,0,0.1); display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 24px; background-color: rgba(255,255,255,0.3); color: black; box-sizing: border-box;">{time_str}</div>
        <div style="height: 120px; width: 100%; display: flex; flex-direction: column; justify-content: {mat_justify}; align-items: center; font-weight: bold; font-size: 28px; padding-top: 5px; padding-bottom: 5px; box-sizing: border-box;">{mat}</div>
        {program_html}
    </div>
</div>"""
        html_content += card_html.replace('\n', '')
        
    html_content += "</div>"
    
    st.markdown(html_content, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
