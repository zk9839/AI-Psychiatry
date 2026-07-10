from fastapi import FastAPI, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from openai import OpenAI
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from datetime import datetime
from email.message import EmailMessage
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import base64
import os
import psycopg2
from dotenv import load_dotenv


load_dotenv()


OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")
GOOGLE_TOKEN_JSON = os.getenv("GOOGLE_TOKEN_JSON")
DATABASE_URL = os.getenv("DATABASE_URL")

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
            revision TEXT,
            summary TEXT,
            doctor_email TEXT
        );
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
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]


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


@app.get("/")
def home():
    return FileResponse("Welcome.html")

@app.get("/intake")
def intake():
    return FileResponse("User_Input.html")



def save_patient_to_db(age, gender, duration, sleep, energy, stress, mood, concern, revision, summary, doctor_email):
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
            revision,
            summary,
            doctor_email
        )
        VALUES (
            NOW(), %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
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
        revision,
        summary,
        doctor_email
    ))

    conn.commit()
    cur.close()
    conn.close()
    
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
    revision: str = Form(""),
    doctor_email: str = Form("")
):
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

If the user provides a revision request, follow it when generating the summary.
If no revision request is provided, generate the best summary you can.

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

User revision request:
{revision}
"""
    )

    summary = response.output_text
    print("Patient saved to PostgreSQL successfully.")
    save_patient_to_db(
    age,
    gender,
    duration,
    sleep,
    energy,
    stress,
    mood,
    concern,
    revision,
    summary,
    doctor_email
)

    print("Patient saved to PostgreSQL successfully.")

    if doctor_email:
        send_summary_email(doctor_email, summary)

    return {
        "received": concern,
        "summary": summary,
        "email_sent": bool(doctor_email)
    }

