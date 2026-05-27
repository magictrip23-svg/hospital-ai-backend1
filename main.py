import base64
import json
import io
import os
import re
import zipfile
import olefile
import sqlite3
import time
from datetime import datetime
from PyPDF2 import PdfReader
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from openai import OpenAI
import uvicorn
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True, 
    allow_methods=["*"],
    allow_headers=["*"],
)

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=OPENAI_API_KEY)

def init_db():
    conn = sqlite3.connect("hospital_ai.db")
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, password TEXT)")
    cursor.execute("INSERT OR IGNORE INTO users (username, password) VALUES ('admin', '1234')")
    cursor.execute("CREATE TABLE IF NOT EXISTS projects (project_id TEXT PRIMARY KEY, project_name TEXT, filename TEXT, plan_text TEXT, budget INTEGER)")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS minutes (
            id INTEGER PRIMARY KEY AUTOINCREMENT, project_id TEXT, meeting_type TEXT, store TEXT, 
            date TEXT, amount INTEGER, plan_task TEXT, time TEXT, location TEXT, attendees TEXT, 
            content TEXT, status TEXT, violation_reason TEXT, input_guide TEXT
        )
    """)
    cursor.execute("CREATE TABLE IF NOT EXISTS contracts (id INTEGER PRIMARY KEY AUTOINCREMENT, project_id TEXT, researcher_name TEXT, salary INTEGER, start_date TEXT, end_date TEXT)")
    conn.commit()
    conn.close()

init_db()

def extract_text(file_bytes, filename):
    text = ""
    ext = filename.lower().split('.')[-1]
    try:
        if ext == "pdf":
            reader = PdfReader(io.BytesIO(file_bytes))
            for page in reader.pages: text += page.extract_text() + "\n"
        elif ext == "hwp":
            f = io.BytesIO(file_bytes)
            if olefile.isOleFile(f):
                ole = olefile.OleFileIO(f)
                if ole.exists('PrvText'): text = ole.openstream('PrvText').read().decode('utf-16le', errors='ignore')
        elif ext == "hwpx":
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
                for item in z.namelist():
                    if item.endswith('.xml'): text += z.read(item).decode('utf-8', errors='ignore')
        elif ext == "docx":
            doc_obj = Document(io.BytesIO(file_bytes))
            for p in doc_obj.paragraphs: text += p.text + "\n"
            for table in doc_obj.tables:
                for row in table.rows:
                    for cell in row.cells: text += cell.text + " "
                    text += "\n"
    except Exception as e:
        print(f"추출 에러: {e}")
    return text

@app.post("/signup")
async def signup(username: str = Form(...), password: str = Form(...)):
    conn = sqlite3.connect("hospital_ai.db")
    cursor = conn.cursor()
    cursor.execute("SELECT username FROM users WHERE username = ?", (username,))
    if cursor.fetchone():
        conn.close()
        return {"status": "error", "message": "이미 존재하는 아이디입니다."}
    cursor.execute("INSERT INTO users (username, password) VALUES (?, ?)", (username, password))
    conn.commit()
    conn.close()
    return {"status": "success", "message": "계정 생성이 완료되었습니다."}

@app.post("/login")
async def login(username: str = Form(...), password: str = Form(...)):
    conn = sqlite3.connect("hospital_ai.db")
    cursor = conn.cursor()
    cursor.execute("SELECT password FROM users WHERE username = ?", (username,))
    row = cursor.fetchone()
    conn.close()
    if not row or row[0] != password:
        return {"status": "error", "message": "비밀번호 또는 아이디가 일치하지 않습니다."}
    return {"status": "success"}

@app.post("/setup-demo-data")
async def setup_demo_data():
    project_id = "proj_seoul_st_mary_demo"
    project_name = "(예시)서울성모병원 연구과제"
    budget = 100000000
    conn = sqlite3.connect("hospital_ai.db")
    cursor = conn.cursor()
    
    cursor.execute("DELETE FROM projects WHERE project_id = ?", (project_id,))
    cursor.execute("INSERT INTO projects (project_id, project_name, filename, plan_text, budget) VALUES (?, ?, '서울성모병원_국책과제_사업계획서.hwp', '서울성모병원 연구과제 데이터셋', ?)", (project_id, project_name, budget))
    
    cursor.execute("DELETE FROM contracts WHERE project_id = ?", (project_id,))
    cursor.execute("INSERT INTO contracts (project_id, researcher_name, salary, start_date, end_date) VALUES (?, '김기훈', 2800000, '2022-06-01', '2027-12-31')", (project_id,))
    cursor.execute("INSERT INTO contracts (project_id, researcher_name, salary, start_date, end_date) VALUES (?, '홍준기', 3200000, '2022-01-01', '2022-12-31')", (project_id,))
    
    cursor.execute("DELETE FROM minutes WHERE project_id = ?", (project_id,))
    
    demo_content = """1. 임상시험 방법에 대한 논의
   가. 5년치의 데이터를 다 쓰는 방안과 1회성의 데이터로 재발률을 예측하는 방안에 대한 비교 논의
   나. 가장 최근 검사일 기준의 1회성 데이터로 임상시험 진행하기로 결정
2. 임상시험용 SW 수정에 관한 논의
   가. 초음파 팝업 전체 삭제
   나. 재발률 퍼센트로 표기하는 UI 수정

[제출용 자동 기안 공문]
수신: 서울성모병원 연구부
제목: 국책과제 연구활동비(회의비) 지출 결의 상신

1. 귀 부서의 노고에 감사드립니다.
2. 관련근거: 서울성모병원 연구비 관리 지침
3. 위 관련하여, 위 연구개발과제의 성공적 수행을 위하여 첨부와 같이 기 집행된 연구비 법인카드 지출 내역(회의비)을 결의하오니, 검토 후 승인하여 주시기 바랍니다.

AI센터장"""
    
    cursor.execute("""
        INSERT INTO minutes (project_id, meeting_type, store, date, amount, plan_task, time, location, attendees, content, status, violation_reason, input_guide)
        VALUES (?, 'conference', '청주경복궁', '2024-04-08', 283000, 
        '임상시험계획서 검토',
        '2024년 4월 8일 19:09 ~ 20:24', '서울성모병원 서관 9층 회의실', '내부: 김기훈, 이유진 외 4명 / 외부: 이강빈(아이도트)',
        ?, 'normal', '정상', '비목: 직접비 > 연구활동비 > 회의비\\n결제수단: 연구비 법인카드\\n금액: 283,000원')
    """, (project_id, demo_content))
    
    conn.commit()
    conn.close()
    return {"status": "success", "project_id": project_id, "project_name": project_name, "budget": budget}

@app.post("/upload-plan")
async def upload_plan(project_id: str = Form(...), project_name: str = Form(...), budget: int = Form(...), plan: UploadFile = File(...)):
    file_bytes = await plan.read()
    full_text = extract_text(file_bytes, plan.filename)
    plan_text_sliced = full_text[:15000] if full_text.strip() else "텍스트를 추출할 수 없거나 비어있는 파일입니다."
    conn = sqlite3.connect("hospital_ai.db")
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO projects (project_id, project_name, filename, plan_text, budget) VALUES (?, ?, ?, ?, ?)", (project_id, project_name, plan.filename, plan_text_sliced, budget))
    conn.commit()
    conn.close()
    return {"status": "success"}

@app.post("/upload-contract")
async def upload_contract(project_id: str = Form(...), contract_file: UploadFile = File(...)):
    file_bytes = await contract_file.read()
    filename = contract_file.filename
    ext = filename.lower().split('.')[-1]
    
    salary_instruction = (
        "주의: 'salary' 키의 값은 절대로 콤마(,), '원', '₩' 등의 텍스트가 포함되지 않은 순수 정수 숫자(Integer) 형식이어야 한다.\n"
        "중요 지침: 계약서 내부의 '전체 계약 기간'과 '총 임금 액수'가 수학적으로 상충하거나 모순되더라도, 절대 총액을 기간으로 나누지 마라."
    )
    
    try:
        if ext in ["jpg", "jpeg", "png", "gif", "bmp", "webp"]:
            base64_image = base64.b64encode(file_bytes).decode('utf-8')
            messages = [
                {"role": "system", "content": f"너는 인사노무 전문가야. JSON: {{\"name\": \"이름\", \"salary\": 월임금숫자, \"start_date\": \"YYYY-MM-DD\", \"end_date\": \"YYYY-MM-DD\"}}\n{salary_instruction}"},
                {"role": "user", "content": [{"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}]}
            ]
        elif ext in ["pdf", "hwp", "hwpx", "docx"]:
            extracted_text = extract_text(file_bytes, filename)
            if not extracted_text.strip(): return {"status": "error", "message": "가독 텍스트가 없습니다."}
            messages = [
                {"role": "system", "content": f"JSON: {{\"name\": \"이름\", \"salary\": 월임금숫자, \"start_date\": \"YYYY-MM-DD\", \"end_date\": \"YYYY-MM-DD\"}}\n{salary_instruction}"},
                {"role": "user", "content": f"[고용계약서 본문]\n{extracted_text}"}
            ]
        else:
            return {"status": "error", "message": "허용되지 않는 확장자 포맷입니다."}

        response = client.chat.completions.create(model="gpt-4o", response_format={ "type": "json_object" }, messages=messages)
        data = json.loads(response.choices[0].message.content)
        
        raw_salary = str(data.get("salary", "0"))
        cleaned_salary = "".join(filter(str.isdigit, raw_salary))
        salary_val = int(cleaned_salary) if cleaned_salary else 0
        
        if 'extracted_text' in locals() and extracted_text:
            for line in extracted_text.split('\n'):
                match = re.search(r'(?:매월|월\s*급여|월\s*보수)[^0-9\n]{0,30}([0-9,]{6,9})', line)
                if match:
                    regex_salary = int(match.group(1).replace(',', ''))
                    if regex_salary > 0:
                        salary_val = regex_salary
                        break
        
        conn = sqlite3.connect("hospital_ai.db")
        cursor = conn.cursor()
        cursor.execute("DELETE FROM contracts WHERE project_id = ? AND researcher_name = ?", (project_id, data.get("name")))
        cursor.execute("""
            INSERT INTO contracts (project_id, researcher_name, salary, start_date, end_date)
            VALUES (?, ?, ?, ?, ?)
        """, (project_id, data.get("name"), salary_val, data.get("start_date"), data.get("end_date")))
        conn.commit()
        conn.close()
        
        return {"status": "success", "data": {**data, "salary": salary_val}}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@app.get("/contracts-list/{project_id}")
async def get_contracts(project_id: str):
    conn = sqlite3.connect("hospital_ai.db")
    cursor = conn.cursor()
    cursor.execute("SELECT researcher_name, salary, start_date, end_date FROM contracts WHERE project_id = ?", (project_id,))
    rows = cursor.fetchall()
    conn.close()
    return [{"name": r[0], "salary": r[1], "start_date": r[2], "end_date": r[3]} for r in rows]

@app.post("/upload-receipt")
async def process_receipt(project_id: str = Form(...), receipts: list[UploadFile] = File(...), category: str = Form("conference")):
    conn = sqlite3.connect("hospital_ai.db")
    cursor = conn.cursor()
    cursor.execute("SELECT plan_text FROM projects WHERE project_id = ?", (project_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return {"error": "사업계획서가 먼저 등록되어야 합니다."}
    plan_text = row[0]
    output_results = []
    
    for receipt in receipts:
        receipt_contents = await receipt.read()
        base64_image = base64.b64encode(receipt_contents).decode('utf-8')
        
        common_instruction = "너는 서울성모병원 연구부의 수석 행정관이야. 제공된 [사업계획서]를 기반으로 과업 연계성을 입증해라."

        if category == "equipment":
            system_prompt = f"""{common_instruction}
            - 시설장비비 규정을 심사하고 대금 지급 공문을 작성해라.
            - 공문 작성 시 반드시 1.인사말, 2.관련근거, 3.요청사항(가,나,다 항목 포함) 형태의 표준 양식을 준수해라.
            {{
                "store": "공급사명", "date": "취득일자", "amount": 10000, "plan_task": "도입 타당성",
                "settlement_status": "normal", "violation_reason": "없음", "system_input_guide": "비목: 직접비 > 연구시설장비비",
                "official_memo": {{
                    "receiver": "서울성모병원 연구부",
                    "title": "국책과제 연구시설장비 대금 지급 요청",
                    "body": "1. 귀 부서의 노고에 감사드립니다.\\n2. 관련근거: 서울성모병원 연구비 지침\\n3. 위 관련하여, 다음과 같이 장비 대금 지급을 요청하오니 검토 후 승인하여 주시기 바랍니다.\\n 가. 안건: ...\\n 나. 금액: ...\\n 다. 사유: ...",
                    "sender": "AI센터장"
                }},
                "minutes": {{ "time": "검수일자", "location": "장소", "attendees": "검수원", "content": "검수조서 본문" }}
            }}"""
        elif category == "material":
            system_prompt = f"""{common_instruction}
            - 재료비 규정을 심사하고 대금 지급 공문을 작성해라.
            - 공문 작성 시 반드시 1.인사말, 2.관련근거, 3.요청사항(가,나,다 항목 포함) 형태의 표준 양식을 준수해라.
            {{
                "store": "공급사명", "date": "입고일자", "amount": 10000, "plan_task": "연계 타당성",
                "settlement_status": "normal", "violation_reason": "없음", "system_input_guide": "비목: 직접비 > 연구재료비",
                "official_memo": {{
                    "receiver": "서울성모병원 연구부",
                    "title": "국책과제 연구재료비 대금 지급 요청",
                    "body": "1. 귀 부서의 노고에 감사드립니다.\\n2. 관련근거: ...\\n3. 위 관련하여, 다음과 같이 재료비 지급을 요청하오니 승인하여 주시기 바랍니다.\\n 가. 안건: ...\\n 나. 금액: ...\\n 다. 사유: ...",
                    "sender": "AI센터장"
                }},
                "minutes": {{ "time": "검수일자", "location": "장소", "attendees": "검수원", "content": "검수 본문" }}
            }}"""
        else:
            system_prompt = f"""{common_instruction}
            - 회의비 규정을 심사해라. 회의비는 이미 '법인카드'로 결제된 건이므로 절대 '업체에 대금 지급을 요청한다'는 말을 쓰지 마라.
            - 반드시 JSON 내 `minutes.content`에 제공된 사업계획서를 바탕으로 회의록 양식에 들어갈 '1. 임상시험 논의 가. ... 2. SW 수정 논의 가. ...' 와 같이 아주 구체적이고 전문적인 [회의록 상세 본문]을 길게 창작해라.
            - 공문 작성 시 반드시 1.인사말, 2.관련근거, 3.요청사항(가,나,다 항목 포함) 형태의 표준 양식을 준수해라.
            {{
                "store": "가맹점명", "date": "결제일시", "amount": 10000, "plan_task": "회의 목적 (예: 임상시험계획서 검토 등)",
                "settlement_status": "normal / caution / invalid", "violation_reason": "판정 사유",
                "system_input_guide": "비목: 연구활동비 > 회의비",
                "official_memo": {{
                    "receiver": "서울성모병원 연구부",
                    "title": "국책과제 회의비 지출 결의(법인카드) 상신",
                    "body": "1. 귀 부서의 노고에 감사드립니다.\\n2. 관련근거: 서울성모병원 연구비 지침\\n3. 위 관련하여, 다음과 같이 기 집행된 카드 지출 내역을 결의하오니 승인하여 주시기 바랍니다.\\n 가. 회의목적: ...\\n 나. 집행금액: ...",
                    "sender": "AI센터장"
                }},
                "minutes": {{ 
                    "time": "202X년 X월 X일 19:00 ~ 20:30", 
                    "location": "서울성모병원 회의실", 
                    "attendees": "내부: OOO / 외부: OOO", 
                    "content": "1. 임상시험 방법에 대한 논의\\n 가. 5년치 데이터 사용 방안...\\n2. SW 수정 논의\\n 가. 초음파 팝업 삭제..." 
                }}
            }}"""

        try:
            # 🛠️ 버그 수정: 중복된 "content" 키 구조를 단일 배열 객체로 합쳐 데이터 손실 방지
            response = client.chat.completions.create(
                model="gpt-4o",
                response_format={ "type": "json_object" },
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": [
                        {"type": "text", "text": f"[사업계획서 발췌]\n{plan_text}\n\n[영수증/계산서 이미지 분석 요청]"},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
                    ]}
                ]
            )
            result = json.loads(response.choices[0].message.content)
            m = result.get('minutes', {})
            memo = result.get('official_memo', {})
            
            # 🛠️ undefined 원천 차단: 텍스트 섞임 및 자료형 누락 에러 방어 처리
            raw_amount = str(result.get("amount", "0"))
            cleaned_amount = "".join(filter(str.isdigit, raw_amount))
            amount_val = int(cleaned_amount) if cleaned_amount else 0
            
            result['amount'] = amount_val
            result['store'] = result.get('store', '알 수 없음')
            result['plan_task'] = result.get('plan_task', '과업 연계성 검토 수행')
            result['settlement_status'] = result.get('settlement_status', 'normal')
            result['violation_reason'] = result.get('violation_reason', '정상')
            result['system_input_guide'] = result.get('system_input_guide', '비목: 연구활동비 > 회의비')
            
            memo_text = f"수신: {memo.get('receiver', '서울성모병원 연구부')}\n제목: {memo.get('title', '')}\n\n{memo.get('body', '')}\n\n{memo.get('sender', '')}"
            final_content = f"{m.get('content', '')}\n\n[제출용 자동 기안 공문]\n{memo_text}"
            
            cursor.execute("""
                INSERT INTO minutes (project_id, meeting_type, store, date, amount, plan_task, time, location, attendees, content, status, violation_reason, input_guide)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                project_id, category, result.get('store'), result.get('date'), result.get('amount'),
                result.get('plan_task'), m.get('time'), m.get('location'), m.get('attendees'), final_content,
                result.get('settlement_status'), result.get('violation_reason'), result.get('system_input_guide')
            ))
            
            result['minute_id'] = cursor.lastrowid
            result['final_content'] = final_content
            output_results.append(result)
        except Exception as e:
            output_results.append({"error": str(e)})
            
    conn.commit()
    conn.close()
    return output_results

@app.post("/upload-audio")
async def process_audio(project_id: str = Form(...), audio: UploadFile = File(...)):
    conn = sqlite3.connect("hospital_ai.db")
    cursor = conn.cursor()
    cursor.execute("SELECT plan_text FROM projects WHERE project_id = ?", (project_id,))
    row = cursor.fetchone()
    plan_text = row[0] if row else "등록된 사업계획서 없음"
    try:
        audio_bytes = await audio.read()
        audio_buffer = io.BytesIO(audio_bytes)
        audio_buffer.name = audio.filename 
        transcript = client.audio.transcriptions.create(model="whisper-1", file=audio_buffer, language="ko")
        raw_text = transcript.text
        
        response = client.chat.completions.create(
            model="gpt-4o", response_format={ "type": "json_object" },
            messages=[
                {"role": "system", "content": "너는 서울성모병원 연구부 행정 매니저야. 받아쓰기 원문을 국가과제 서식으로 구조화 요약해줘."},
                {"role": "user", "content": f"[계획서]\n{plan_text}\n\n[받아쓰기]\n{raw_text}"}
            ]
        )
        result = json.loads(response.choices[0].message.content)
        cursor.execute("""
            INSERT INTO minutes (project_id, meeting_type, store, date, amount, plan_task, time, location, attendees, content, status, violation_reason, input_guide)
            VALUES (?, 'real', '원내 회의', '-', 0, ?, ?, ?, ?, ?, 'normal', '실제 회의 요약', '원내 회의로 재정 시스템 입력 대상 아님')
        """, (project_id, result.get('plan_task'), result.get('time'), result.get('location'), result.get('attendees'), result.get('content')))
        conn.commit()
        conn.close()
        return {"status": "success", "data": result, "raw_text": raw_text}
    except Exception as e:
        conn.close()
        return {"error": str(e)}

@app.post("/execute-expense-rpa")
async def execute_expense_rpa(minute_id: int = Form(...)):
    conn = sqlite3.connect("hospital_ai.db")
    cursor = conn.cursor()
    cursor.execute("SELECT store, amount, meeting_type FROM minutes WHERE id = ?", (minute_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return {"status": "error", "message": "결재 데이터를 찾을 수 없습니다."}
        
    store_name, amount, meeting_type = row
    cursor.execute("UPDATE minutes SET status = 'settled' WHERE id = ?", (minute_id,))
    conn.commit()
    conn.close()
    
    target_dept = "서울성모병원 연구부"
    if meeting_type == 'conference':
        message = f"🎉 [지출 결의 상신 완료]\n{target_dept}로 지출 결의서 및 협조 공문이 전송되었습니다.\n(회의비는 기 결제된 건이므로 별도의 업체 송금은 진행되지 않습니다.)"
    else:
        message = f"🎉 [대금 지급 상신 완료]\n{target_dept}로 지출 결의서 및 협조 공문이 전송되었습니다.\n승인 완료 시, ezbaro 연동을 통해 '{store_name}'(으)로 대금({amount:,}원)이 자동 송금됩니다."
    return {"status": "success", "message": message}

@app.post("/sync-ezbaro")
async def sync_ezbaro(project_id: str = Form(...), billing_month: str = Form(...)):
    conn = sqlite3.connect("hospital_ai.db")
    cursor = conn.cursor()
    cursor.execute("SELECT researcher_name, end_date, salary FROM contracts WHERE project_id = ?", (project_id,))
    contracts = cursor.fetchall()
    
    if not contracts:
        conn.close()
        return {"status": "error", "message": "등록된 연구원 데이터가 없습니다."}
        
    expired_people = []
    active_payroll_amount = 0
    
    for name, end_date_str, salary in contracts:
        try:
            contract_end = datetime.strptime(end_date_str, "%Y-%m-%d")
            billing_date = datetime.strptime(f"{billing_month}-01", "%Y-%m-%d")
            if contract_end < billing_date:
                expired_people.append(f"{name}(만료일: {end_date_str})")
            else:
                active_payroll_amount += salary
        except: pass
            
    if expired_people:
        conn.close()
        return {"status": "intercepted", "message": "계약기간 경과 인원이 있어 자동 결재를 차단했습니다.", "details": expired_people}
        
    try:
        time.sleep(1.5) 
        cursor.execute("""
            INSERT INTO minutes (project_id, meeting_type, store, date, amount, plan_task, time, location, attendees, content, status, violation_reason, input_guide)
            VALUES (?, 'labor', '인사행정시스템', ?, ?, '연구인건비 자동 계상', ?, '원내 인사 전산망', '계약 연구원 전체', '자동 집행 완료', 'settled', '정상', 'RPA 상신 완료')
        """, (project_id, billing_month, active_payroll_amount, billing_month))
        conn.commit()
        conn.close()
        return {"status": "success", "message": f"🎉 [{billing_month}] 총 {active_payroll_amount:,}원의 인건비 지급 기안이 서울성모병원 연구부로 상신되었습니다."}
    except Exception as e:
        if conn: conn.close()
        return {"status": "error", "message": str(e)}

@app.get("/project-stats/{project_id}")
async def get_project_stats(project_id: str):
    conn = sqlite3.connect("hospital_ai.db")
    cursor = conn.cursor()
    cursor.execute("SELECT project_name, budget FROM projects WHERE project_id = ?", (project_id,))
    proj_row = cursor.fetchone()
    if not proj_row:
        conn.close()
        return {"budget": 0, "total_spent": 0, "remaining": 0, "normal": 0, "caution": 0, "invalid": 0}
    p_name, budget = proj_row
    cursor.execute("SELECT amount, status FROM minutes WHERE project_id = ?", (project_id,))
    minutes_rows = cursor.fetchall()
    conn.close()
    total_spent = sum(row[0] for row in minutes_rows if row[0])
    return {
        "project_name": p_name, "budget": budget, "total_spent": total_spent, "remaining": budget - total_spent,
        "normal": sum(1 for r in minutes_rows if r[1] in ['normal', 'settled']),
        "caution": sum(1 for r in minutes_rows if r[1] == 'caution'),
        "invalid": sum(1 for r in minutes_rows if r[1] == 'invalid'),
        "total_count": len(minutes_rows)
    }

@app.get("/export-excel")
async def export_excel(project_id: str):
    conn = sqlite3.connect("hospital_ai.db")
    cursor = conn.cursor()
    cursor.execute("SELECT meeting_type, store, date, amount, plan_task, status, violation_reason FROM minutes WHERE project_id = ? AND meeting_type != 'real' ORDER BY id ASC", (project_id,))
    rows = cursor.fetchall()
    conn.close()
    
    wb = Workbook()
    ws = wb.active
    ws.title = "ezbaro_Bulk_Template"
    headers = ["비목", "세부비목", "증빙일자(납품일)", "가맹점/공급처", "총지출액", "공급가액", "부가세", "적요(집행타당성 사유)", "감사결과"]
    ws.append(headers)
    
    for row in rows:
        m_type, store, date_str, amount, plan_task, status, violation = row
        amount = amount if amount else 0
        supply_value = int(amount / 1.1) if m_type != "labor" else amount
        vat = amount - supply_value if m_type != "labor" else 0
        ws.append([m_type, "세부비목", date_str, store, amount, supply_value, vat, plan_task, status])
        
    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return StreamingResponse(file_stream, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename={project_id}_Upload.xlsx"})

@app.get("/export-word")
async def export_word(project_id: str):
    conn = sqlite3.connect("hospital_ai.db")
    cursor = conn.cursor()
    cursor.execute("SELECT project_name FROM projects WHERE project_id = ?", (project_id,))
    proj_row = cursor.fetchone()
    project_name = proj_row[0] if proj_row else "서울성모병원 원내과제"
    
    cursor.execute("SELECT meeting_type, store, amount, plan_task, time, location, attendees, content FROM minutes WHERE project_id = ? ORDER BY id ASC", (project_id,))
    rows = cursor.fetchall()
    conn.close()
    
    doc = Document()
    
    for idx, row in enumerate(rows):
        m_type, store, amount, plan_task, meeting_time, location, attendees, content = row
        
        if m_type == "conference":
            title = doc.add_paragraph()
            r = title.add_run("회   의   록")
            r.font.size = Pt(18)
            r.bold = True
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            doc.add_paragraph("1. 과제 개요")
            t1 = doc.add_table(rows=5, cols=4)
            t1.style = 'Table Grid'
            t1.cell(0,0).text = "사 업 명"
            t1.cell(0,1).text = "국가연구개발사업(원내과제)"
            t1.cell(0,2).text = "주관기관"
            t1.cell(0,3).text = "서울성모병원"
            t1.cell(1,0).text = "연구개발과제명"
            t1.cell(1,1).merge(t1.cell(1,3)).text = project_name
            t1.cell(2,0).text = "연구책임자"
            t1.cell(2,1).text = "소속: 임상과"
            t1.cell(2,2).text = "성명:"
            t1.cell(2,3).text = "김기훈 (서명)"
            t1.cell(3,0).text = "전문기관명"
            t1.cell(3,1).text = "연구부"
            t1.cell(3,2).text = "과제번호"
            t1.cell(3,3).text = project_id
            t1.cell(4,0).text = "전체 연구기간"
            t1.cell(4,1).merge(t1.cell(4,3)).text = "2024. 1. 1. ~ 2027. 12. 31."
            
            doc.add_paragraph("\n2. 회의 내용")
            t2 = doc.add_table(rows=7, cols=2)
            t2.style = 'Table Grid'
            t2.columns[0].width = Inches(1.5)
            t2.columns[1].width = Inches(5.0)
            
            t2.cell(0,0).text = "회 의 일 시"
            t2.cell(0,1).text = meeting_time if meeting_time else ""
            t2.cell(1,0).text = "회 의 장 소"
            t2.cell(1,1).text = location if location else ""
            t2.cell(2,0).text = "참 석 자"
            t2.cell(2,1).text = attendees if attendees else ""
            t2.cell(3,0).text = "카드 사용처"
            t2.cell(3,1).text = store if store else ""
            t2.cell(4,0).text = "집 행 금 액"
            t2.cell(4,1).text = f"금{amount:,}원" if amount else ""
            t2.cell(5,0).text = "회 의 목 적"
            t2.cell(5,1).text = plan_task if plan_task else ""
            t2.cell(6,0).text = "회 의 내 용"
            
            clean_content = content.split("[제출용 자동 기안 공문]")[0].strip() if content else ""
            t2.cell(6,1).text = clean_content
            
            if idx < len(rows) - 1:
                doc.add_page_break()
        else:
            doc.add_paragraph(f"[{m_type.upper()}] 정산 증빙서 - {store}")
            doc.add_paragraph(f"금액: {amount:,}원")
            clean_content = content.split("[제출용 자동 기안 공문]")[0].strip() if content else ""
            doc.add_paragraph(f"내용:\n{clean_content}")
            if idx < len(rows) - 1:
                doc.add_page_break()
                
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return StreamingResponse(file_stream, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", headers={"Content-Disposition": f"attachment; filename={project_id}_Meeting_Minutes.docx"})

@app.get("/export-memo/{minute_id}")
async def export_memo(minute_id: int):
    conn = sqlite3.connect("hospital_ai.db")
    cursor = conn.cursor()
    cursor.execute("SELECT content FROM minutes WHERE id = ?", (minute_id,))
    row = cursor.fetchone()
    conn.close()

    if not row:
        return {"error": "데이터를 찾을 수 없습니다."}

    full_content = row[0]
    memo_parts = full_content.split("[제출용 자동 기안 공문]")
    memo_text = memo_parts[1].strip() if len(memo_parts) > 1 else "공문 내용이 없습니다."

    doc = Document()
    
    header = doc.add_paragraph()
    r = header.add_run("서 울 성 모 병 원  기 안 문")
    r.font.size = Pt(20)
    r.bold = True
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("\n")
    
    for line in memo_text.split('\n'):
        if line.startswith("수신:") or line.startswith("제목:"):
            p = doc.add_paragraph()
            p.add_run(line).bold = True
        elif line.startswith("AI센터장") or line.startswith("연구책임자"):
            p = doc.add_paragraph("\n\n" + line)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.runs[0].font.size = Pt(16)
            p.runs[0].bold = True
        else:
            doc.add_paragraph(line)
            
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return StreamingResponse(file_stream, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", headers={"Content-Disposition": f"attachment; filename=Official_Memo_{minute_id}.docx"})

if __name__ == "__main__":
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
