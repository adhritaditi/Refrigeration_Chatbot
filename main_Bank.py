from openai import OpenAI
from fastapi import FastAPI, Form, Request, WebSocket
from typing import Annotated
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
import os
from dotenv import load_dotenv
from docx import Document
 
# Load environment variables from .env file
load_dotenv()
 
# Initialize OpenAI with API key from environment variable
openai = OpenAI(
    api_key=os.getenv('OPENAI_API_SECRET_KEY')
)
 
# Create FastAPI app
app = FastAPI()
 
# Mount the "templates1" directory to serve static files (like images)
app.mount("/static", StaticFiles(directory="templates1"), name="static")
 
# Specify the directory for Jinja2 templates
templates = Jinja2Templates(directory="templates1")
 
chat_responses = []
 
def load_chat_log_from_docx(file_path):
    """Load chat log from a Word document, handling various formats."""
    doc = Document(file_path)
    chat_log = []
 
    for para in doc.paragraphs:
        if para.text.strip():
            chat_log.append({'role': 'user', 'content': para.text.strip()})
 
    for table in doc.tables:
        for row in table.rows:
            row_text = ' | '.join(cell.text.strip() for cell in row.cells if cell.text.strip())
            if row_text:
                chat_log.append({'role': 'user', 'content': row_text})
 
    return chat_log
 
# Load the initial chat log from the Word document in the templates1 folder
chat_log = load_chat_log_from_docx("templates1/Bank.docx")
 
def trim_chat_log(log, max_tokens=2000):
    """Trim the chat log to stay within a certain token limit."""
    trimmed_log = []
    token_count = 0
 
    # Reverse the log to start from the most recent messages
    for entry in reversed(log):
        entry_tokens = len(entry['content'].split())
        if token_count + entry_tokens > max_tokens:
            break
        trimmed_log.insert(0, entry)  # Insert at the beginning
        token_count += entry_tokens
 
    return trimmed_log
 
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
 
        # Trim the chat log to avoid exceeding the context length
        trimmed_chat_log = trim_chat_log(chat_log)
 
        try:
            response = openai.chat.completions.create(
                model='gpt-3.5-turbo',
                messages=trimmed_chat_log,
                temperature=0.6,
                stream=True
            )
 
            ai_response = ''
 
            for chunk in response:
                if chunk.choices[0].delta.content is not None:
                    ai_response += chunk.choices[0].delta.content
                    await websocket.send_text(chunk.choices[0].delta.content)
 
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
 
    # Trim the chat log to avoid exceeding the context length
    trimmed_chat_log = trim_chat_log(chat_log)
 
    response = openai.chat.completions.create(
        model='gpt-4',
        messages=trimmed_chat_log,
        temperature=0.6
    )
 
    bot_response = response.choices[0].message.content
    chat_log.append({'role': 'assistant', 'content': bot_response})
    chat_responses.append(bot_response)
 
    # Ensure the response is a coherent sentence
    return templates.TemplateResponse("home.html", {"request": request, "chat_responses": chat_responses})
 
@app.get("/image", response_class=HTMLResponse)
async def image_page(request: Request):
    """Serve the image generation page."""
    return templates.TemplateResponse("image.html", {"request": request})
 
@app.post("/image", response_class=HTMLResponse)
async def create_image(request: Request, user_input: Annotated[str, Form()]):
    """Handle image generation requests."""
    response = openai.images.generate(
        prompt=user_input,
        n=1,
        size="256x256"
    )
 
    image_url = response.data[0].url
    return templates.TemplateResponse("image.html", {"request": request, "image_url": image_url})
