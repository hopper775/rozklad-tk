import pandas as pd
from datetime import date
import time
import json
import requests
from io import BytesIO
import os, sys

def parse_dataframe(df):
    days = ["mon",
            "tue",
            "wed",
            "thu",
            "fri"]
    result = {}
    combined_result = []

    for day_idx, day in enumerate(days):
        day_schedule = []

        num_col = df.columns[1 + day_idx * 3]
        subj_col = df.columns[2 + day_idx * 3]
        room_col = df.columns[3 + day_idx * 3]

        row = 0
        while row < len(df):
            num = df.iloc[row][num_col] 
            subject = df.iloc[row][subj_col]
            room = df.iloc[row][room_col]
            if pd.notna(subject):
                teacher = df.iloc[row + 1][subj_col] if row + 1 < len(df) else ""
                teacher = str(teacher).replace("(", "").replace(")", "")
                weeks = "all"
                #визначення тижня пар
                if "(1" in str(subject):
                    old_num = num           #зберігаємо номер пари якщо вона по чисельнику
                    weeks = "odd"           #чисельник (1.3)
                elif "(2" in str(subject):
                    weeks = "even"          #знаменник (2.4)
                    num = old_num           #присвоюємо номер пари по знаменнику бо його нема в таблиці

                lessons = ({
                    "num": str(int(num)),
                    "time": getTime(int(num)),
                    "subject": str(subject),
                    "room": "" if pd.isna(room) else str(int(room)) if isinstance(room, (int,float)) else str(room),
                    "teacher": teacher,
                    "weeks": weeks
                })
                
                #ігнорує дивні дублікати останньої пари 
                #if day_schedule and day_schedule[-1]["num"] == lessons["num"]:
                #    pass
                #else:
                day_schedule.append(lessons)
                
            row += 2
            
        result[day] =  day_schedule
        result["metadata"] = {
            "Time": time.time() #час в форматі UNIX таймстемпа
        }
    return result

def getTime(num):
    match num:
        case 1:
            time = "8:00-9:20"
        case 2:
            time = "9:30-10:50"
        case 3:
            time = "11:10-12:30"
        case 4:
            time = "12:40-14:00"
        case 5:
            time = "14:10-15:30"
        case 6:
            time = "15:40-17:00"
        case 7:
            time = "17:10-18:30"
        case 8:
            time = "18:40-20:00"
    return time

def get_sheet(url):
    sheet = requests.get(url)
    sheet.raise_for_status()
    return sheet
    
if __name__=="__main__":
    G_sheets_url = "https://docs.google.com/spreadsheets/d/1SU2Y5o8zzJsYwzvu07GkFWv0zHRNxytFFmRLUXfyYYk/export?gid=743307244&format=xlsx"
    
    try:
        sheet = get_sheet(G_sheets_url)
    except requests.exceptions.RequestException:
        time.sleep(60)
        sheet = get_sheet(G_sheets_url)


    csv_data = pd.read_excel(BytesIO(sheet.content),
                              skiprows=6,
                              #skipfooter=1,
                              sheet_name=0
                            )
    
    
    #видаляє пусті строки
    csv_data = csv_data.dropna(axis=1, how="all")
    csv_data = csv_data.dropna(how="all")
    
    start_EL = csv_data[csv_data["Група"].astype(str).str.startswith("ЕЛ-")].index[0]
    start_TK = csv_data[csv_data["Група"].astype(str).str.startswith("ТК-")].index[0]

    EL_data = csv_data.loc[start_EL:start_TK-1].reset_index(drop=True)
    TK_data = csv_data.loc[start_TK:].reset_index(drop=True)

    schedule = parse_dataframe(TK_data)
    
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(sys.argv[0])))
    output_path = os.path.join(base_path, "schedule.json")


    with open (output_path, "w", encoding="utf-8") as f:
        json.dump(schedule, f, ensure_ascii=False, indent=2)



