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
    doc = Document(file_path)
    chat_log = []
    for para in doc.paragraphs:
        if para.text.startswith("role:"):
            role = para.text.split(":")[1].strip()
            content = para.text.split(":")[2].strip()
            chat_log.append({'role': role, 'content': content})
    return chat_log
 
# Load the initial chat log from the Word document in the templates1 folder
chat_log = load_chat_log_from_docx("templates1/Bank.docx")
 
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
