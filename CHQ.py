import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os
from flask import Flask, render_template, request, abort, redirect, url_for, flash
import sqlite3
from PIL import Image, ImageDraw, ImageFont
import webbrowser
from dotenv import load_dotenv

#데이터베이스 초기화
def initialize_database(filepaths):
        
        try:
            # 데이터베이스 연결
            conn = sqlite3.connect('data.db') 
            cursor = conn.cursor()
            
            # 기존 테이블 삭제 (초기화 과정)
            cursor.execute("DROP TABLE IF EXISTS data")
            
            
            # 데이터베이스 테이블 생성
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS data (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    no INTEGER,
                    sample_id TEXT,
                    call_rate REAL,
                    dqc REAL,
                    qc_call_rate REAL,
                    chip_barcode TEXT,
                    chip_position TEXT,
                    order_no TEXT,
                    tube_id TEXT,
                    lims_sex TEXT,
                    inferred_sex TEXT,
                    sex_match BOOLEAN,
                    wkst TEXT
                )
            ''')

            # 엑셀 파일 데이터 삽입
            
            for filepath in filepaths:
                filename = os.path.basename(filepath)  # 파일 이름 추출
                wkst = filename.split('_')[1]  # '_' 구분자의 두 번째 부분 추출
                
                
                df = pd.read_excel(filepath)
                df = df[['No.', 'Sample ID', 'Call Rate', 'DQC', 'QC Call Rate', 'Chip Barcode', 
                         'Chip Position', 'Order No.', 'Tube ID', 'Lims Sex', 'Inferred Sex', 'Sex Match']]
                

                # 열 이름을 데이터베이스 테이블의 열 이름에 맞게 변경
                df.columns = ['no', 'sample_id', 'call_rate', 'dqc', 'qc_call_rate', 'chip_barcode', 
                              'chip_position', 'order_no', 'tube_id', 'lims_sex', 'inferred_sex', 'sex_match']
                
                # 'file_part' 열을 추가
                df['wkst'] = wkst
                
                # 소수점 자릿수 조정
                df['call_rate'] = df['call_rate'].round(3)
                df['dqc'] = df['dqc'].round(3)
                df['qc_call_rate'] = df['qc_call_rate'].round(3)
                #print(df.head())
                
                df.to_sql('data', conn, if_exists='append', index=False)

            conn.commit()
            
        except sqlite3.Error as e:
            print(f"SQLite error: {e}")
        finally:
            if conn:
                conn.close()
                
                
def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2 - 200
    window.geometry(f"{width}x{height}+{x}+{y}")


def create_table_layout(df, output_dir, table_width=300, table_height=200, cell_size=50): 
    # 이미지 크기 설정
    image_width = table_width * cell_size
    image_height = table_height * cell_size
    
    # 이미지 생성
    img = Image.new('RGBA', (image_width, image_height), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    
    # 폰트 설정 (글씨 크기 조절)
    font_size = 6  # 원하는 글씨 크기로 설정
    try:
        font = ImageFont.truetype("arial.ttf", font_size)  # 시스템 폰트 파일 경로 필요
    except IOError:
        font = ImageFont.load_default()  # 기본 폰트 사용
    
    # 칼럼 선택 (예시: 'Call Rate')
    column = 'call_rate'
    threshold_value = 97  # 예시 기준값
    
    # 데이터 값 표시
    for index, row in df.iterrows():
        chip_position = row['chip_position']
        value = row[column]
        
        # 좌표 계산 (chip_position 형식 확인 및 예외 처리)
        try:
            col = int(chip_position[1:]) - 1
            row = ord(chip_position[0].upper()) - ord('A')
            
            # 기준값 이하의 값은 빨간색으로 표시
            if value < threshold_value:
                color = (255, 0, 0, 178)  # 투명도 70의 빨간색
            elif value < 90:
                color = (0, 0, 0, 178)  # 투명도 70의 검은색
            else:
                color = (255, 255, 0, 178)  # 노랑
            
            # 셀 그리기
            x0 = col * cell_size
            y0 = row * cell_size
            x1 = x0 + cell_size
            y1 = y0 + cell_size
            draw.rectangle([x0, y0, x1, y1], fill=color)
            # draw.text((x0 + 10, y0 + 10), str(value), fill=(0, 0, 0, 255))

            # 텍스트 크기 조절 및 그리기
   # 텍스트 크기 조절 및 그리기
            text = str(value)
            bbox = draw.textbbox((x0, y0), text, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            text_x = x0 + (cell_size - text_width) / 2
            text_y = y0 + (cell_size - text_height) / 2
            draw.text((text_x, text_y), text, font=font, fill=(0, 0, 0, 255))
            
            
        except (ValueError, IndexError) as e:
            print(f"Error processing chip position {chip_position}: {e}")
    
    # 이미지 저장
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    img.save(os.path.join(output_dir, 'table_layout.png'))


# 3. 동적 HTML 페이지 생성

from flask import Flask, render_template, request, abort
import sqlite3
import pandas as pd

app = Flask(__name__)

"""로딩애니메이션"""

@app.route('/loading')
def loading():
    return render_template('loading.html')

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        files = request.files.getlist('files')
        if not files or files[0].filename == '':
            flash('파일을 선택하세요.')
            return redirect(url_for('index'))

        # 업로드된 파일 저장
        filepaths = []
        for file in files:
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            filepaths.append(filepath)

        # DB 생성
        initialize_database(filepaths)
        flash(f'{len(filepaths)}개의 파일이 업로드되어 처리되었습니다.')
        
        # DB에 저장 후 파일 삭제
        for filepath in filepaths:
            os.remove(filepath)


        return redirect(url_for('index'))
    
      
    
    sex_mismatch, wkst = fetch_summary_data()
    return render_template('index.html', sex_mismatch=sex_mismatch, wkst = wkst)

def fetch_summary_data():
    conn = sqlite3.connect('data.db')
    cursor = conn.cursor()
    # 1. 'Sex Match'가 FALSE인 행 추출
    cursor.execute("SELECT * FROM data WHERE sex_match = 'FALSE'")
    sex_mismatch = cursor.fetchall()
    
    # 3. 고유한 WKST ID 가져오기. 
    cursor.execute("SELECT DISTINCT wkst FROM data")
    wkst = [row[0] for row in cursor.fetchall()]

    conn.close()
    return sex_mismatch, wkst


@app.route('/sex_mismatch')
def view_sex_mismatch():
    conn = sqlite3.connect('data.db')
    query = "SELECT * FROM data WHERE sex_match = FALSE AND inferred_sex != 'unknown'"
    
    query6 = """
    SELECT wkst FROM data
    """
    
    df = pd.read_sql_query(query, conn)
    df6 = pd.read_sql_query(query6, conn)
    conn.close()

    # if df.empty:    
        # abort(404)
        
    # 원하는 열 순서 설정
    
    #df = df.drop[''no', 'sample_id', 'call_rate', 'dqc', 'qc_call_rate', 'chip_barcode', 'chip_position', 'order_no', 'tube_id', 'lims_sex', 'inferred_sex', 'sex_match']']

    desired_order = ['qc_call_rate', 'sample_id', 'call_rate', 'chip_barcode', 'chip_position','order_no',  'tube_id','lims_sex', 'inferred_sex', 'wkst']  # 원하는 열 순서로 변경(수정하기)
    df = df[desired_order]
    

    #로그 작성
    if df.empty:
        log = '성별불일치 샘플 없음.'   
    else:
        log = '안녕하세요 수행부입니다. 림스성별과 실험성별이 일치하지 않아 동의서 및 의뢰서 확인 부탁드립니다.<br>'
        
    comments = []
    for _, row in df.iterrows():
        comments.append(f"{row['sample_id']}<br>")
    
    log = f"{log} {''.join(comments)} 감사합니다."
    columns = df.columns.tolist()
    
    content = "mismatch"
    
    return render_template('view.html', title="성별불일치", columns=columns, data=df.to_dict(orient='records'), comments=log, content = content, data4 = df6.to_dict(orient = 'records'))


@app.route('/view/<wkst>')
def view_chip(wkst):
    conn = sqlite3.connect('data.db')
    query = """
    SELECT * FROM data 
    WHERE wkst = ? AND 
    (call_rate < 0.9 OR dqc < 0.82 OR qc_call_rate < 0.9)
    """
    
    query2 = """
    SELECT * FROM data
    WHERE wkst = ? AND
    (call_rate < 0.97)"""
    
    query3 = """
    SELECT * FROM data
    WHERE wkst = ?
    """
    
    query4 = """
    SELECT * FROM data WHERE wkst =? AND sex_match = FALSE AND inferred_sex != 'unknown'
    """
    query4_1 = """
    SELECT * FROM data WHERE wkst =? AND sex_match = FALSE AND inferred_sex == 'unknown'
    """
    
    query5 = """
    SELECT * FROM data
    WHERE wkst = ? AND dqc >= (SELECT AVG(dqc) * 0.98 FROM data) AND qc_call_rate < 0.90;
    """
    
    query6 = """
    SELECT wkst FROM data
    """
    
    
    
    # query3 = """
    # SELECT * FROM data 
    # WHERE wkst = ? AND 
    # (call_rate >= 0.9 AND dqc > 0.93 AND qc_call_rate < 0.9)
    # """
    
    df = pd.read_sql_query(query, conn, params=(wkst,))
    df2 = pd.read_sql_query(query2, conn, params=(wkst,))
    df3 = pd.read_sql_query(query3, conn, params=(wkst,))
    df4 = pd.read_sql_query(query4, conn, params=(wkst,))
    df4_1 = pd.read_sql_query(query4_1, conn, params=(wkst,))
    df5 = pd.read_sql_query(query5, conn, params=(wkst,))
    df6 = pd.read_sql_query(query6, conn)
    conn.close()

    # if df.empty:
    #     abort(404)

    # 3. 실패한 샘플에 대한 코멘트 생성

    comments = []
    for _, row in df.iterrows():
        comments.append(f"{row['order_no']}({row['tube_id']})")
    
    mismatch = []
    unknown  = []
    contam = []
    
    for _, row in df4.iterrows():
        mismatch.append(f"{row['order_no']}({row['sample_id']})")
        
    for _, row in df4_1.iterrows():
        unknown.append(f"{row['order_no']}({row['sample_id']})")
        
    for _, row in df5.iterrows():
        contam.append(f"{row['order_no']}({row['tube_id']})")
    
        
    log = 'Affymetrix Axiom CHQ 결과. chip control probe로 실험성공 유무를 확인합니다. \n'
    if df.empty and df4.empty and df4_1.empty and df5.empty and df2.empty:
        log = log + f"모든 control결과는 정상입니다."
    else:
        if not df4.empty:
            log = log  + f"성별불일치 샘플 {len(df4)}개, {', '.join(mismatch)}는 의뢰서 및 동의서와 림스 성별정보가 일치하고, 수행과정에 특이사항이 없어 재실험없이 분석 진행하는 건입니다.<br>"
        if not df4_1.empty:
            log = log  + f"성별 unknown샘플 {len(df4_1)}개, {', '.join(unknown)} 입니다.<br>"
        if not df5.empty:
            log = log + f"QC CR FAIL로 재실험 처리한 샘플은 {len(df5)}개로, {', '.join(contam)} 입니다. "
        log = log  + f"CR기준 미만 샘플은 {len(df2)}개이며 재샘플링 필요한 FAIL샘플, {len(comments)}개 {', '.join(comments)} 입니다."
    
    df['chip_barcode'] = df['chip_barcode'].str[-3:]
    code3 = df3['chip_barcode'].iloc[0][-3:]
    desired_order = ['order_no', 'sample_id', 'tube_id', 'chip_barcode', 'chip_position','call_rate','dqc',
                    'qc_call_rate','lims_sex', 'inferred_sex']  # 원하는 열 순서로 변경(수정하기)
    df = df[desired_order]

    columns = df.columns.tolist()
    
    content = "wkst"
    return render_template('view.html', wkst=wkst, columns=columns, data=df.to_dict(orient='records'), comments=log,\
        data2 = df3.to_dict(orient='records'), data3 = df4.to_dict(orient = 'records'), content = content, data4 = df6.to_dict(orient = 'records'), code3 = code3)
    
if __name__ == '__main__':
    # ✅ 업로드 폴더 설정
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)

    app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
    
    load_dotenv()  # 현재 디렉토리의 .env 파일 읽음
    app.secret_key = os.getenv('SECRET_KEY')


    # 웹 서버 시작 전 브라우저 창 열기
    webbrowser.open_new('http://127.0.0.1:5000/')
    app.run(debug=True, use_reloader=False)