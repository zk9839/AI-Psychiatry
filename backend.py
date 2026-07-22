from fastapi import FastAPI, Form, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from openai import OpenAI
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from datetime import datetime, timedelta
from email.message import EmailMessage
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import base64
import os
import json
import psycopg2
from dotenv import load_dotenv


load_dotenv()


OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")
GOOGLE_TOKEN_JSON = os.getenv("GOOGLE_TOKEN_JSON")
DATABASE_URL = os.getenv("DATABASE_URL")

MAX_QUESTIONS = 4
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
CALENDAR_SCOPES = ["https://www.googleapis.com/auth/calendar.events"]

def get_db_connection():
    return psycopg2.connect(DATABASE_URL)

def create_patients_table():
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS patients (
            id SERIAL PRIMARY KEY,
            timestamp TIMESTAMP,
            age TEXT,
            gender TEXT,
            duration TEXT,
            sleep TEXT,
            energy TEXT,
            stress TEXT,
            mood TEXT,
            concern TEXT,
            summary TEXT,
            doctor_email TEXT
        );
    """)

    cur.execute("""
        ALTER TABLE patients ADD COLUMN IF NOT EXISTS conversation TEXT;
    """)

    cur.execute("""
        ALTER TABLE patients ADD COLUMN IF NOT EXISTS appointment_datetime TEXT;
    """)

    conn.commit()
    cur.close()
    conn.close()

def write_google_auth_files():
    if GOOGLE_CREDENTIALS_JSON and not os.path.exists("credentials.json"):
        with open("credentials.json", "w") as f:
            f.write(GOOGLE_CREDENTIALS_JSON)

    if GOOGLE_TOKEN_JSON and not os.path.exists("token.json"):
        with open("token.json", "w") as f:
            f.write(GOOGLE_TOKEN_JSON)


write_google_auth_files()
create_patients_table()
app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

client = OpenAI(api_key=OPENAI_API_KEY)

EXCEL_FILE = "patients.xlsx"


def send_summary_email(doctor_email, summary):
    if not doctor_email:
        return

    creds = Credentials.from_authorized_user_file(
        "token.json",
        GMAIL_SCOPES
    )

    service = build("gmail", "v1", credentials=creds)

    message = EmailMessage()
    message["To"] = doctor_email
    message["From"] = SENDER_EMAIL
    message["Subject"] = "Patient Visit Preparation Summary"

    message.set_content(f"""
Patient Visit Preparation Summary

{summary}

Note: This is an AI-generated visit preparation summary. It is not a diagnosis or treatment recommendation.
""")

    encoded_message = base64.urlsafe_b64encode(
        message.as_bytes()
    ).decode()

    create_message = {
        "raw": encoded_message
    }

    service.users().messages().send(
        userId="me",
        body=create_message
    ).execute()


def create_calendar_event(appointment_datetime, doctor_email, summary):
    if not appointment_datetime:
        return None

    creds = Credentials.from_authorized_user_file(
        "token.json",
        CALENDAR_SCOPES
    )

    service = build("calendar", "v3", credentials=creds)

    start_dt = datetime.fromisoformat(appointment_datetime)
    end_dt = start_dt + timedelta(minutes=30)

    event = {
        "summary": "Patient Appointment",
        "description": summary,
        "start": {"dateTime": start_dt.isoformat()},
        "end": {"dateTime": end_dt.isoformat()},
    }

    if doctor_email:
        event["attendees"] = [{"email": doctor_email}]

    created_event = service.events().insert(
        calendarId="primary",
        body=event
    ).execute()

    return created_event.get("htmlLink")


@app.get("/")
def home():
    return FileResponse("welcome.html")

@app.get("/intake")
def intake():
    return FileResponse("User_Input.html")



def save_patient_to_db(age, gender, duration, sleep, energy, stress, mood, concern, summary, doctor_email, conversation, appointment_datetime):
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO patients (
            timestamp,
            age,
            gender,
            duration,
            sleep,
            energy,
            stress,
            mood,
            concern,
            summary,
            doctor_email,
            conversation,
            appointment_datetime
        )
        VALUES (
            NOW(), %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
        );
    """, (
        age,
        gender,
        duration,
        sleep,
        energy,
        stress,
        mood,
        concern,
        summary,
        doctor_email,
        conversation,
        appointment_datetime
    ))

    conn.commit()
    cur.close()
    conn.close()


@app.post("/start_interview")
async def start_interview(request: Request):
    body = await request.json()
    intake = body.get("intake", {})

    prompt = f"""You are Sage, a warm, calm assistant conducting a brief pre-appointment intake interview.

You have the patient's initial intake form below. Based on ALL of these answers together, decide if there is ONE valuable follow-up question to ask that would help the doctor prepare — something that fills a real gap or clarifies something vague.

Ask it in a warm, conversational tone, as Sage would.

Do not diagnose. Do not assess severity or risk. Do not recommend treatment. Only ask for clarifying detail a doctor would find useful.

If a follow-up question is warranted, return ONLY that question (under 25 words).
If the intake is already sufficiently detailed, return exactly: NONE

Patient intake:
Age: {intake.get('age', '')}
Gender: {intake.get('gender', '')}
Duration of concern: {intake.get('duration', '')}
Sleep quality: {intake.get('sleep', '')}
Energy level: {intake.get('energy', '')}
Stress level: {intake.get('stress', '')}
Mood: {intake.get('mood', '')}
Main concern: {intake.get('concern', '')}
"""

    response = client.responses.create(
        model="gpt-4o-mini",
        input=prompt
    )

    result = response.output_text.strip()

    blocked_terms = ["concerning", "severe", "risk", "diagnos", "urgent"]
    if any(term in result.lower() for term in blocked_terms):
        result = "NONE"

    question = None if result == "NONE" else result

    return {"question": question}


@app.post("/continue_interview")
async def continue_interview(request: Request):
    body = await request.json()
    intake = body.get("intake", {})
    conversation = body.get("conversation", [])

    questions_so_far = sum(1 for turn in conversation if turn.get("role") == "ai")

    if questions_so_far >= MAX_QUESTIONS:
        return {"question": None}

    conversation_text = ""
    for turn in conversation:
        role = "Sage" if turn.get("role") == "ai" else "Patient"
        conversation_text += f"{role}: {turn.get('content', '')}\n"

    prompt = f"""You are Sage, a warm, calm assistant conducting a brief pre-appointment intake interview.

Patient intake:
Age: {intake.get('age', '')}
Gender: {intake.get('gender', '')}
Duration of concern: {intake.get('duration', '')}
Sleep quality: {intake.get('sleep', '')}
Energy level: {intake.get('energy', '')}
Stress level: {intake.get('stress', '')}
Mood: {intake.get('mood', '')}
Main concern: {intake.get('concern', '')}

Conversation so far:
{conversation_text}

Based on everything above, decide if there is ONE more valuable follow-up question to ask. Ask it in a warm, conversational tone, as Sage would. Do not diagnose. Do not assess severity or risk. Do not recommend treatment. Only ask for clarifying detail a doctor would find useful. Avoid repeating anything already asked.

If another question is warranted, return ONLY that question (under 25 words).
If enough has been gathered, return exactly: NONE
"""

    response = client.responses.create(
        model="gpt-4o-mini",
        input=prompt
    )

    result = response.output_text.strip()

    blocked_terms = ["concerning", "severe", "risk", "diagnos", "urgent"]
    if any(term in result.lower() for term in blocked_terms):
        result = "NONE"

    question = None if result == "NONE" else result

    return {"question": question}


@app.post("/summarize")
def summarize(
    age: str = Form(""),
    gender: str = Form(""),
    duration: str = Form(""),
    sleep: str = Form(""),
    energy: str = Form(""),
    stress: str = Form(""),
    mood: str = Form(""),
    concern: str = Form(...),
    doctor_email: str = Form(""),
    appointment_datetime: str = Form(""),
    conversation: str = Form("[]")
):
    try:
        conversation_list = json.loads(conversation)
    except json.JSONDecodeError:
        conversation_list = []

    conversation_text = ""
    if conversation_list:
        conversation_text = "Follow-up interview:\n"
        for turn in conversation_list:
            role = "Question" if turn.get("role") == "ai" else "Patient answer"
            conversation_text += f"{role}: {turn.get('content', '')}\n"

    response = client.responses.create(
        model="gpt-4o-mini",
        input=f"""
You are a mental health visit preparation assistant.

Create a neutral, clinician-ready visit summary based on the intake information.

Do not diagnose.
Do not recommend treatment.
Do not give medication advice.
Include all information including age, gender, etc.
Do not use asterisks.
Return plain text only.

Patient information:
Age: {age}
Gender: {gender}
Duration of concern: {duration}
Sleep quality: {sleep}
Energy level: {energy}
Stress level: {stress}
Mood: {mood}

Main concern:
{concern}

{conversation_text}
"""
    )

    summary = response.output_text

    save_patient_to_db(
        age,
        gender,
        duration,
        sleep,
        energy,
        stress,
        mood,
        concern,
        summary,
        doctor_email,
        conversation,
        appointment_datetime
    )

    print("Patient saved to PostgreSQL successfully.")

    if doctor_email:
        send_summary_email(doctor_email, summary)

    calendar_link = create_calendar_event(appointment_datetime, doctor_email, summary)

    return {
        "received": concern,
        "summary": summary,
        "email_sent": bool(doctor_email),
        "calendar_link": calendar_link
    }