from openai import OpenAI
from fastapi import FastAPI, Form, Request, WebSocket, UploadFile
from typing import Annotated
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
import os
from dotenv import load_dotenv
from docx import Document
import pandas as pd
import fitz  # PyMuPDF to read PDF
import logging
 
# Load environment variables from .env file
load_dotenv()
 
# Initialize OpenAI with API key from environment variable
openai = OpenAI(
    api_key=os.getenv('OPENAI_API_SECRET_KEY')
)
 
# Set up logging for debugging purposes
logging.basicConfig(level=logging.INFO)
 
# Create FastAPI app
app = FastAPI()
 
# Mount the "templates1" directory to serve static files (like images)
app.mount("/static", StaticFiles(directory="templates1"), name="static")
 
# Specify the directory for Jinja2 templates
templates = Jinja2Templates(directory="templates1")
 
chat_responses = []
 
def load_chat_log_from_docx(file_path):
    """Load text and tables from a Word document with improved handling."""
    doc = Document(file_path)
    chat_log = []
 
    # Extract text from paragraphs, skipping headers/footers
    for para in doc.paragraphs:
        text = para.text.strip()
        if text and not para.style.name.startswith('Header') and not para.style.name.startswith('Footer'):
 
 chat_log.append({'role': 'user', 'content': text})
 
    # Extract text from tables, including nested tables
    for table in doc.tables:
        for row in table.rows:
            row_text = ' | '.join(cell.text.strip() for cell in row.cells if cell.text.strip())
            if row_text:
 chat_log.append({'role': 'user', 'content': row_text})
 
    return chat_log
 
def load_chat_log_from_excel(file_path):
    """Load chat log by reading all worksheets from an Excel file, handling NaN and merged cells."""
    chat_log = []
    try:
        excel_file = pd.ExcelFile(file_path)
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
 
            # Log sheet name
logging.info(f'Reading sheet: {sheet_name}')
            
            # Replace NaN values with empty strings
            df.fillna('', inplace=True)
            
            for row in df.itertuples(index=False):
                row_texts = []
                for cell in row:
                    if isinstance(cell, str):
                        # Handle multi-line text in Excel cells
                        cell = cell.replace('\n', ' ')
                    row_texts.append(str(cell))
                
                row_text = ' | '.join(row_texts)
chat_log.append({'role': 'user', 'content': row_text})
 
    except Exception as e:
        logging.error(f"Error reading Excel file {file_path}: {e}")
 
    return chat_log
 
def load_chat_log_from_pdf(file_path):
    """Load chat log from a PDF file."""
    chat_log = []
    try:
        pdf_document = fitz.open(file_path)
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text = page.get_text("text")
            if text.strip():
                chat_log.append({'role': 'user', 'content': text.strip()})
 
    except Exception as e:
        logging.error(f"Error reading PDF file {file_path}: {e}")
 
    return chat_log
 
def load_chat_logs_from_folder(folder_path):
    """Load chat logs by reading from all Word, Excel, and PDF files in a folder."""
    chat_log = []
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        extension = filename.split('.')[-1].lower()
 
        if extension == 'docx':
            chat_log.extend(load_chat_log_from_docx(file_path))
        elif extension == 'xlsx':
            chat_log.extend(load_chat_log_from_excel(file_path))
        elif extension == 'pdf':
            chat_log.extend(load_chat_log_from_pdf(file_path))
 
    return chat_log
 
# Load the initial chat log from all files in the "templates1" folder
chat_log = load_chat_logs_from_folder("templates1/")
 
@app.get("/", response_class=HTMLResponse)
async def chat_page(request: Request):
    """Serve the main chat page."""
    return templates.TemplateResponse("home.html", {"request": request, "chat_responses": chat_responses})
 
@app.websocket("/ws")
async def chat(websocket: WebSocket):
    """WebSocket connection for real-time chat."""
    await websocket.accept()
 
    while True:
        user_input = await websocket.receive_text()
        chat_log.append({'role': 'user', 'content': user_input})
        chat_responses.append(user_input)
 
        try:
            response = openai.chat.completions.create(
                model='gpt-3.5-turbo',
                messages=chat_log,
                temperature=0.6,
                stream=True
            )
 
            ai_response = ''
 
            for chunk in response:
                if chunk.choices[0].delta.content is not None:
                    ai_response += chunk.choices[0].delta.content
                    await websocket.send_text(chunk.choices[0].delta.content)
 
            # Append AI response to the log in sentence format
            chat_log.append({'role': 'assistant', 'content': ai_response})
            chat_responses.append(ai_response)
 
        except Exception as e:
            await websocket.send_text(f'Error: {str(e)}')
            break
 
@app.post("/", response_class=HTMLResponse)
async def chat(request: Request, user_input: Annotated[str, Form()]):
    """Handle chat interactions from the form."""
    chat_log.append({'role': 'user', 'content': user_input})
    chat_responses.append(user_input)
 
    response = openai.chat.completions.create(
        model='gpt-4',
        messages=chat_log,
        temperature=0.6
    )
 
    bot_response = response.choices[0].message.content
    chat_log.append({'role': 'assistant', 'content': bot_response})
    chat_responses.append(bot_response)
 
    return templates.TemplateResponse("home.html", {"request": request, "chat_responses": chat_responses})
 
@app.post("/upload", response_class=HTMLResponse)
async def upload_file(request: Request, file: UploadFile):
    """Handle file uploads (Word, Excel, PDF) and load chat logs."""
    file_location = f"templates1/{file.filename}"
    with open(file_location, "wb+") as file_object:
        file_object.write(file.file.read())
 
    # Reload chat logs from all files in the templates1 folder after upload
    chat_log.extend(load_chat_logs_from_folder("templates1/"))
 
    return templates.TemplateResponse("home.html", {"request": request, "chat_responses": chat_responses})
