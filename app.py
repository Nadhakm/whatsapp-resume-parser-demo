from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse
import gspread
import pickle
import re
import requests
from io import BytesIO
from PyPDF2 import PdfReader
from docx import Document
from dotenv import load_dotenv
import os
import openai
import pytesseract
from PIL import Image

load_dotenv()

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

openai.api_key = os.environ.get("OPENAI_API_KEY")
TWILIO_ACCOUNT_SID = os.environ.get("TWILIO_ACCOUNT_SID")
TWILIO_AUTH_TOKEN = os.environ.get("TWILIO_AUTH_TOKEN")

app = Flask(__name__)

def get_gsheet_client():
    with open("token.pkl", "rb") as token:
        creds = pickle.load(token)
    client = gspread.authorize(creds)
    return client

def extract_email(text):
    match = re.search(r"[\w\.-]+@[\w\.-]+", text)
    return match.group(0) if match else ""

def extract_phone(text):
    match = re.search(r"\+?\d[\d\s-]{7,}\d", text)
    return match.group(0) if match else ""

def extract_name(text):
    return " ".join(text.strip().split()[:3])

def extract_pdf(url):
    try:
        response = requests.get(url, auth=(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN))
        response.raise_for_status()
        pdf_bytes = BytesIO(response.content)
        reader = PdfReader(pdf_bytes)
        text = ""
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
        return text.strip()
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return "[Could not extract PDF text]"

def extract_docx(media_url):
    try:
        response = requests.get(media_url, auth=(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN))
        response.raise_for_status()

        doc = Document(BytesIO(response.content))
        text = "\n".join([p.text for p in doc.paragraphs])
        return text or "(No text found in document.)"

    except Exception as e:
        print(f"Error reading DOCX: {e}")
        return "Sorry, I couldn’t read the DOCX file properly."
    
def extract_image(media_url):
    try:
        response = requests.get(media_url, auth=(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN))
        response.raise_for_status()

        img = Image.open(BytesIO(response.content))
        text = pytesseract.image_to_string(img)
        return text.strip() or "(No text detected in image.)"

    except Exception as e:
        print(f"Error reading image: {e}")
        return "Sorry, I couldn’t extract text from the image."

@app.route("/whatsapp", methods=["POST"])
def whatsapp_webhook():
    msg = request.form.get("Body", "")
    sender = request.form.get("From")

    print(f"Message from {sender}: {msg}")

    media_count = int(request.form.get("NumMedia", 0))
    if media_count > 0:
        media_url = request.form.get("MediaUrl0")
        content_type = request.form.get("MediaContentType0")
        print(f"Attachment received: {media_url} ({content_type})")

        if content_type == "application/pdf":
            msg = extract_pdf(media_url)
        elif content_type in ["application/vnd.openxmlformats-officedocument.wordprocessingml.document", "application/msword"]:
            msg = extract_docx(media_url)
        elif content_type.startswith("image/"):
            msg = extract_image(media_url)
        else:
            msg = "Unsupported file type. Please send a text, PDF, DOCX, or image."


    name = extract_name(msg)
    email = extract_email(msg)
    phone = extract_phone(msg)

    client = get_gsheet_client()
    sheet = client.open("Data Demo").sheet1
    sheet.append_row([sender, name, email, phone, msg])

    resp = MessagingResponse()
    resp.message("Message/file received and parsed successfully!")
    return str(resp)

if __name__ == "__main__":
    app.run(port=5000, debug=True)
