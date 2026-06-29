from fastapi import FastAPI, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from openai import OpenAI
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from datetime import datetime
from email.message import EmailMessage
import smtplib
import os


OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

client = OpenAI()

EXCEL_FILE = "patients.xlsx"



def send_summary_email(doctor_email, summary):
    if not doctor_email:
        return

    app_password = os.getenv("GMAIL_APP_PASSWORD")

    if not app_password:
        raise Exception("GMAIL_APP_PASSWORD is not set")

    msg = EmailMessage()
    msg["Subject"] = "Patient Visit Preparation Summary"
    msg["From"] = SENDER_EMAIL
    msg["To"] = doctor_email

    msg.set_content(f"""
Patient Visit Preparation Summary

{summary}

Note: This is an AI-generated visit preparation summary. It is not a diagnosis or treatment recommendation.
""")

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(SENDER_EMAIL, app_password)
        smtp.send_message(msg)


@app.get("/")
def home():
    return FileResponse("User_Input.html")


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
Include all information including age, gender, etc but don't include the asterisks
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

    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active

        ws.append([
            "Timestamp",
            "Age",
            "Gender",
            "Duration",
            "Sleep",
            "Energy",
            "Stress",
            "Mood",
            "Concern",
            "Revision",
            "Summary",
            "Doctor Email"
        ])
    else:
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active

    ws.append([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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
    ])

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 15
    ws.column_dimensions["G"].width = 10
    ws.column_dimensions["H"].width = 15
    ws.column_dimensions["I"].width = 40
    ws.column_dimensions["J"].width = 40
    ws.column_dimensions["K"].width = 170
    ws.column_dimensions["L"].width = 30

    current_row = ws.max_row
    ws.row_dimensions[current_row].height = 120

    for cell in ws[current_row]:
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(EXCEL_FILE)

    if doctor_email:
        send_summary_email(doctor_email, summary)

    return {
        "received": concern,
        "summary": summary,
        "email_sent": bool(doctor_email)
    }