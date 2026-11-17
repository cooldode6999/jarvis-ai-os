import os
import io
import re
import sys
import json
import time
import sqlite3
import ctypes
import asyncio
import threading
import webbrowser
import subprocess
from urllib.parse import quote
from datetime import datetime, date
import requests
import psutil
import keyboard
import pyautogui
import pyperclip
import numpy as np
import sounddevice as sd
from scipy.io.wavfile import write
from PIL import ImageGrab
import speech_recognition as sr
import edge_tts
import pywhatkit
from pypdf import PdfReader
from pptx import Presentation
from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
import uvicorn

from google import genai
from google.genai import types
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ==========================================================
# CONFIGURATION
# ==========================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
from dotenv import load_dotenv
load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY_HERE")
VOICEPRINT_FILE = os.path.join(BASE_DIR, "nayan_voiceprint.npy")
DB_FILE = os.path.join(BASE_DIR, "jarvis_brain.db")
NOTES_DIR = os.path.join(BASE_DIR, "PDEU_Notes")
os.makedirs(NOTES_DIR, exist_ok=True)

# Combined Scopes for BOTH Google Drive & Gmail (Ek hi login se dono chalenge!)
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/gmail.readonly",
]

SAMPLE_RATE = 16000
MAX_LISTEN_SECONDS = 18
JARVIS_VOICE = "en-IN-PrabhatNeural"
OLLAMA_MODELS = ["qwen2.5-coder:7b", "llama3.2:3b", "qwen2.5:7b", "deepseek-r1:7b"]
EMBED_MODEL = "nomic-embed-text"

pyautogui.FAILSAFE = True
app = FastAPI()

STATE = {
    "status": "STANDBY",
    "logs": ["[BOOT] J.A.R.V.I.S. MK-107 // Full Automation (WhatsApp + Gmail + Drive + RAG) Ready."],
    "abort": False,
    "voice": JARVIS_VOICE,
    "wake_active": True,
    "shutdown": False,
    "active_mode": "FULL-AUTO-OS",
    "active_ollama_model": "llama3.2:3b",
}

def add_log(msg: str):
    ts = time.strftime("%H:%M:%S")
    STATE["logs"].append(f"[{ts}] {msg}")
    if len(STATE["logs"]) > 50:
        STATE["logs"].pop(0)
    print(f"[{ts}] {msg}")

# ==========================================================
# 1. LOCAL SQLITE DATABASE (MEMORIES + CHATS + RAG VECTORS)
# ==========================================================
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS memories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            topic TEXT NOT NULL,
            detail TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS chat_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            role TEXT NOT NULL,
            message TEXT NOT NULL,
            timestamp TEXT NOT NULL
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS study_chunks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT NOT NULL,
            chunk_text TEXT NOT NULL,
            embedding_json TEXT NOT NULL
        )
    """)
    conn.commit()
    conn.close()

init_db()

# ==========================================================
# 2. OLLAMA AUTO-BOOT & BACKGROUND PRE-WARM
# ==========================================================
def _prewarm_model_worker(model_name: str):
    try:
        add_log(f"🔥 [PRE-WARM] Loading '{model_name}' into RAM...")
        res = requests.post(
            "http://localhost:11434/api/generate",
            json={"model": model_name, "prompt": "", "keep_alive": -1},
            timeout=90,
        )
        if res.status_code == 200:
            add_log(f"✅ [LOCAL AI READY] '{model_name}' locked in RAM.")
    except Exception:
        pass

def ensure_ollama_running() -> str:
    chosen_model = STATE["active_ollama_model"]
    server_up = False
    try:
        r = requests.get("http://localhost:11434/api/tags", timeout=2)
        if r.status_code == 200:
            server_up = True
            installed = [m["name"] for m in r.json().get("models", [])]
            for preferred in OLLAMA_MODELS:
                match = next((inst for inst in installed if preferred in inst), None)
                if match:
                    chosen_model = match
                    break
            STATE["active_ollama_model"] = chosen_model
    except Exception:
        pass

    if not server_up:
        add_log("⚡ [AUTO-BOOT] Starting Local Ollama Service...")
        subprocess.Popen(
            "ollama serve",
            shell=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
        )
        time.sleep(3)

    threading.Thread(target=_prewarm_model_worker, args=(chosen_model,), daemon=True).start()
    return chosen_model

# ==========================================================
# 3. UNIFIED GOOGLE OAUTH (DRIVE + GMAIL)
# ==========================================================
def get_google_creds():
    creds_path = os.path.join(BASE_DIR, "credentials.json")
    token_path = os.path.join(BASE_DIR, "google_token.json")
    if not os.path.exists(creds_path):
        return None
    creds = Credentials.from_authorized_user_file(token_path, GOOGLE_SCOPES) if os.path.exists(token_path) else None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, GOOGLE_SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w") as token:
            token.write(creds.to_json())
    return creds

def search_and_sync_google_drive(query_topic: str) -> str:
    """Searches Nayan's Google Drive for PDF/PPTX/Docs matching `query_topic`, downloads & indexes them."""
    creds = get_google_creds()
    if not creds:
        return "credentials.json jarvis folder mein rakhein aur Google Cloud Console mein apna email Test Users mein add karein."

    try:
        service = build("drive", "v3", credentials=creds)
        add_log(f"☁️ [GOOGLE DRIVE] Searching: '{query_topic}'...")
        safe_q = query_topic.replace("'", "\\'")
        results = service.files().list(q=f"name contains '{safe_q}' and trashed = false", pageSize=5, fields="files(id, name, mimeType)").execute()
        files = results.get("files", [])
        if not files:
            return f"Google Drive mein '{query_topic}' naam se koi file nahi mili."

        downloaded = []
        for f in files:
            fid, fname, mime = f["id"], f["name"], f["mimeType"]
            if "google-apps.document" in mime or "google-apps.presentation" in mime:
                out_name = f"{fname}.pdf"
                request = service.files().export_media(fileId=fid, mimeType="application/pdf")
            elif fname.lower().endswith((".pdf", ".pptx")):
                out_name = fname
                request = service.files().get_media(fileId=fid)
            else:
                continue

            out_path = os.path.join(NOTES_DIR, out_name)
            fh = io.FileIO(out_path, "wb")
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            fh.close()
            downloaded.append(out_name)

        if downloaded:
            index_pdpu_notes()
            return f"Google Drive se {len(downloaded)} files ({', '.join(downloaded)}) download aur index ho gayi hain!"
        return f"Drive par files milin ({', '.join([x['name'] for x in files])}), lekin PDF/PPT nahi thin."
    except Exception as e:
        return f"Google Drive Error (Test User check karein): {e}"

def read_gmail_digest(max_results: int = 5) -> str:
    """Reads latest unread emails from Nayan's Gmail inbox."""
    creds = get_google_creds()
    if not creds:
        return "Gmail access ke liye credentials.json rakhein aur Google Cloud mein apna email Test Users mein add karein."
    try:
        service = build("gmail", "v1", credentials=creds)
        results = service.users().messages().list(userId="me", labelIds=["UNREAD"], maxResults=max_results).execute()
        messages = results.get("messages", [])
        if not messages:
            return "Inbox clear hai, koi unread email nahi hai."
        subjects = []
        for m in messages:
            txt = service.users().messages().get(userId="me", id=m["id"], format="metadata").execute()
            headers = txt["payload"]["headers"]
            subj = next((h["value"] for h in headers if h["name"] == "Subject"), "No Subject")
            subjects.append(subj)
        return "Latest Unread Emails: " + " | ".join(subjects)
    except Exception as e:
        return f"Gmail Error (Test User check karein): {e}"

# ==========================================================
# 4. RELIABLE WHATSAPP AUTOMATION (DIRECT WEB URL + MEMORY)
# ==========================================================
def send_whatsapp_message(phone_or_name: str, message: str) -> str:
    """
    Sends a WhatsApp message via WhatsApp Web.
    Supports phone numbers (+9198XXXXXXXX) OR contact names saved in SQLite memory!
    """
    target = phone_or_name.strip()
    if not any(ch.isdigit() for ch in target):
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT detail FROM memories WHERE topic LIKE ? OR detail LIKE ? ORDER BY id DESC LIMIT 1", (f"%{target}%", f"%{target}%"))
        row = c.fetchone()
        conn.close()
        if row:
            match = re.search(r"(\+?\d[\d\s-]{8,14}\d)", row[0])
            if match:
                target = match.group(1).replace(" ", "").replace("-", "")

    digits = re.sub(r"\D", "", target)
    if len(digits) < 10:
        return f"'{phone_or_name}' ka valid 10-digit number nahi mila. Pehle bolein: Yaad rakhna {phone_or_name} ka number 98XXXXXXXX hai."

    full_phone = digits if len(digits) > 10 else "91" + digits
    encoded_msg = quote(message)
    wa_url = f"https://web.whatsapp.com/send?phone={full_phone}&text={encoded_msg}"

    add_log(f"💬 [WHATSAPP] Opening chat for +{full_phone}...")
    webbrowser.open(wa_url)

    def _auto_press_enter():
        time.sleep(14)  # WhatsApp Web chat load hone ka wait karein
        if not STATE["abort"]:
            pyautogui.press("enter")
            add_log(f"✅ [WHATSAPP] Message sent to +{full_phone}!")

    threading.Thread(target=_auto_press_enter, daemon=True).start()
    return f"WhatsApp Web khol diya hai, 14 second mein {phone_or_name} ko message send ho jayega."

# ==========================================================
# 5. LOCAL PDF/PPTX RAG, MEMORY & PDEU SCHEDULE
# ==========================================================
def get_ollama_embedding(text: str) -> list:
    try:
        res = requests.post("http://localhost:11434/api/embeddings", json={"model": EMBED_MODEL, "prompt": text}, timeout=20)
        if res.status_code == 200:
            return res.json().get("embedding", [])
    except Exception:
        pass
    return []

def _extract_text_from_file(filepath: str) -> str:
    ext = os.path.splitext(filepath)[1].lower()
    text_parts = []
    try:
        if ext == ".pdf":
            reader = PdfReader(filepath)
            for i, page in enumerate(reader.pages):
                txt = page.extract_text() or ""
                if txt.strip():
                    text_parts.append(f"[Page {i+1}] {txt.strip()}")
        elif ext in [".pptx", ".ppt"]:
            prs = Presentation(filepath)
            for i, slide in enumerate(prs.slides):
                slide_txt = [s.text.strip() for s in slide.shapes if hasattr(s, "text") and s.text.strip()]
                if slide_txt:
                    text_parts.append(f"[Slide {i+1}] " + " ".join(slide_txt))
    except Exception:
        pass
    return "\n".join(text_parts)

def index_pdpu_notes() -> str:
    """Indexes all PDF/PPTX files inside PDEU_Notes into jarvis_brain.db."""
    files = [f for f in os.listdir(NOTES_DIR) if f.lower().endswith((".pdf", ".pptx"))]
    if not files:
        os.startfile(NOTES_DIR)
        return "PDEU_Notes folder khaali hai! Maine folder khol diya hai, usme PDF/PPT files daalein."

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT DISTINCT file_name FROM study_chunks")
    already_indexed = {r[0] for r in c.fetchall()}

    new_files, total_chunks = 0, 0
    for fname in files:
        if fname in already_indexed:
            continue
        full_text = _extract_text_from_file(os.path.join(NOTES_DIR, fname))
        if not full_text:
            continue
        words = full_text.split()
        for idx in range(0, len(words), 105):
            chunk = " ".join(words[idx: idx + 130])
            if len(chunk) < 40:
                continue
            emb = get_ollama_embedding(chunk)
            if emb:
                c.execute("INSERT INTO study_chunks (file_name, chunk_text, embedding_json) VALUES (?, ?, ?)", (fname, chunk, json.dumps(emb)))
                total_chunks += 1
        new_files += 1

    conn.commit()
    conn.close()
    return f"Indexed {new_files} new files ({total_chunks} chunks) into jarvis_brain.db!" if new_files else f"All {len(files)} files already indexed!"

def search_pdpu_notes(question: str, top_k: int = 4) -> str:
    """Searches indexed PDEU lecture notes in jarvis_brain.db."""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT file_name, chunk_text, embedding_json FROM study_chunks")
    rows = c.fetchall()
    conn.close()
    if not rows:
        return "Koi indexed notes nahi mile."

    q_emb = get_ollama_embedding(question)
    scored = []
    if q_emb:
        q_vec = np.array(q_emb, dtype=np.float32)
        q_norm = np.linalg.norm(q_vec) + 1e-8
        for fname, chunk_txt, emb_json in rows:
            d_vec = np.array(json.loads(emb_json), dtype=np.float32)
            sim = float(np.dot(q_vec, d_vec) / (q_norm * (np.linalg.norm(d_vec) + 1e-8)))
            scored.append((sim, fname, chunk_txt))
        scored.sort(key=lambda x: x[0], reverse=True)
    else:
        q_words = set(question.lower().split())
        for fname, chunk_txt, _ in rows:
            scored.append((len(q_words.intersection(set(chunk_txt.lower().split()))), fname, chunk_txt))
        scored.sort(key=lambda x: x[0], reverse=True)

    return "\n\n".join([f"Source ({h[1]}): {h[2]}" for h in scored[:top_k]])

def log_chat_turn(role: str, message: str):
    if not message or not message.strip(): return
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO chat_history (role, message, timestamp) VALUES (?, ?, ?)", (role, message.strip(), datetime.now().strftime("%d %b %Y %I:%M %p")))
    conn.commit()
    conn.close()

def get_recent_chat_context(limit: int = 10) -> list:
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT role, message FROM chat_history ORDER BY id DESC LIMIT ?", (limit,))
    rows = c.fetchall()
    conn.close()
    rows.reverse()
    return [{"role": r[0], "content": r[1]} for r in rows]

def save_memory(topic: str, detail: str) -> str:
    """Saves facts, phone numbers, passwords, or notes permanently in jarvis_brain.db."""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO memories (topic, detail, created_at) VALUES (?, ?, ?)", (topic.strip(), detail.strip(), datetime.now().strftime("%d %b %Y, %I:%M %p")))
    conn.commit()
    conn.close()
    add_log(f"🧠 [SAVED TO DB] {topic}: {detail}")
    return f"Locally save kar liya hai Nayan: {topic} — {detail}."

def recall_memory(query: str = "") -> str:
    """Retrieves saved facts from jarvis_brain.db."""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    if query.strip():
        like_q = f"%{query.strip()}%"
        c.execute("SELECT topic, detail FROM memories WHERE topic LIKE ? OR detail LIKE ? ORDER BY id DESC LIMIT 5", (like_q, like_q))
    else:
        c.execute("SELECT topic, detail FROM memories ORDER BY id DESC LIMIT 6")
    rows = c.fetchall()
    conn.close()
    return " | ".join([f"{r[0]}: {r[1]}" for r in rows]) if rows else "Koi saved entry nahi mili."

def get_pdpu_academic_status() -> dict:
    today = date.today()
    weekday = datetime.now().strftime("%A")
    if weekday in ["Monday", "Tuesday", "Wednesday", "Thursday"]:
        sched = "07:00 AM – 08:00 PM (PDEU CSE Regular)"
    elif weekday == "Friday":
        sched = "09:00 AM – 05:00 PM (PDEU CSE Friday)"
    else:
        sched = "Weekend / Self-Study & Exam Prep"
    return {
        "today": today.strftime("%d %b %Y (%a)"),
        "days_to_endsem": max((date(2026, 11, 26) - today).days, 0),
        "days_to_neat": max((date(2026, 11, 29) - today).days, 0),
        "schedule_today": sched,
        "student_id": "26BCP143",
    }

def check_pdpu_exams_and_schedule(query: str = "") -> str:
    """Checks PDEU End-Sem (26 Nov), Anandi-NEAT (29 Nov), and today's class schedule."""
    info = get_pdpu_academic_status()
    return f"PDEU End-Sem (26 Nov) mein {info['days_to_endsem']} din aur Anandi-NEAT (29 Nov) mein {info['days_to_neat']} din bache hain. Aaj: {info['schedule_today']}."

def activate_workflow_mode(mode_name: str) -> str:
    """Activates multi-step workspace workflows: 'study', 'dev', or 'chill'."""
    m = mode_name.lower().strip()
    if "study" in m or "exam" in m:
        STATE["active_mode"] = "EXAM-FOCUS"
        subprocess.run('taskkill /F /IM "chrome.exe" /T', shell=True, capture_output=True)
        keyboard.send("volume mute")
        os.startfile(NOTES_DIR)
        return "Study Mode active! Distractions band karke PDEU_Notes folder khol diya hai."
    elif "dev" in m or "coding" in m:
        STATE["active_mode"] = "DEV-MATRIX"
        subprocess.Popen("code", shell=True)
        webbrowser.open("https://github.com")
        return "Dev Mode active! VS Code aur GitHub launch kar diya hai."
    elif "chill" in m:
        STATE["active_mode"] = "CHILL-WAVE"
        subprocess.run('taskkill /F /IM "Code.exe" /T', shell=True, capture_output=True)
        pywhatkit.playonyt("lofi hip hop beats to relax study")
        return "Chill Mode active! Lo-Fi beats play kar diye hain."
    return f"Unknown workflow: {mode_name}"

# ==========================================================
# 6. WINDOWS OS, APPS, KEYBOARD & SCREEN VISION
# ==========================================================
def play_youtube(song_or_video: str) -> str:
    pywhatkit.playonyt(song_or_video)
    return f"YouTube par {song_or_video} play kar diya hai."

def open_or_close_app(app_name: str, action: str = "open") -> str:
    name = app_name.lower().strip()
    if action == "close":
        exe_map = {
            "vscode": "Code.exe", "vs code": "Code.exe", "code": "Code.exe",
            "chrome": "chrome.exe", "youtube": "chrome.exe", "browser": "chrome.exe",
            "notepad": "notepad.exe", "calculator": "CalculatorApp.exe",
            "cmd": "cmd.exe", "spotify": "Spotify.exe", "vlc": "vlc.exe",
        }
        exe = exe_map.get(name, f"{name.replace(' ', '')}.exe")
        res = subprocess.run(f"taskkill /F /IM {exe} /T", shell=True, capture_output=True)
        return f"{app_name} band kar diya hai." if res.returncode == 0 else f"{app_name} abhi chal nahi raha."

    websites = {
        "youtube": "https://www.youtube.com", "github": "https://github.com",
        "gmail": "https://mail.google.com", "linkedin": "https://www.linkedin.com",
        "whatsapp": "https://web.whatsapp.com", "chatgpt": "https://chatgpt.com",
    }
    for k, url in websites.items():
        if k in name:
            webbrowser.open(url)
            return f"{k.capitalize()} open kar diya hai."

    pyautogui.press("win")
    time.sleep(0.35)
    pyautogui.write(app_name, interval=0.04)
    time.sleep(0.45)
    pyautogui.press("enter")
    return f"{app_name} launch kar diya hai."

def keyboard_and_mouse_action(action_type: str, value: str = "") -> str:
    """Controls typing, hotkeys (ctrl+c, alt+tab, win+d), or scrolling."""
    if action_type == "type":
        pyautogui.write(value, interval=0.03)
        return f"Typed: {value}"
    elif action_type == "hotkey":
        keys = [k.strip() for k in value.lower().split("+")]
        pyautogui.hotkey(*keys)
        return f"Shortcut {value} executed."
    elif action_type == "scroll":
        pyautogui.scroll(-600 if "down" in value.lower() else 600)
        return f"Scrolled {value}."
    return "GUI action done."

def create_file_or_folder(name: str, is_folder: bool = False, content: str = "") -> str:
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    target_path = os.path.join(desktop, name)
    if is_folder:
        os.makedirs(target_path, exist_ok=True)
        return f"Desktop par '{name}' folder bana diya hai."
    with open(target_path, "w", encoding="utf-8") as f:
        f.write(content)
    return f"Desktop par '{name}' file create kar di hai."

def run_windows_command(ps_command: str) -> str:
    blocked = ["format ", "remove-item -recurse", "rd /s", "del /s", "shutdown /s"]
    if any(b in ps_command.lower() for b in blocked):
        return "Safety Block: Destructive command allowed nahi hai."
    add_log(f"[SHELL] {ps_command}")
    res = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True, timeout=15)
    output = (res.stdout or res.stderr).strip()
    return output[:350] if output else "Command successfully executed."

def analyze_or_fix_screen(task_prompt: str = "Explain what is on my screen.", copy_fix_to_clipboard: bool = False) -> str:
    snap_path = os.path.join(BASE_DIR, "temp_vision.png")
    ImageGrab.grab().save(snap_path)
    try:
        with open(snap_path, "rb") as f:
            img_bytes = f.read()
        prompt = (
            "Provide ONLY the corrected code without backticks so it can be pasted directly."
            if copy_fix_to_clipboard else task_prompt + " (Reply concisely in Hinglish.)"
        )
        resp = client.models.generate_content(
            model="gemini-3.8-flash",
            contents=[types.Part.from_bytes(data=img_bytes, mime_type="image/png"), prompt],
        )
        if copy_fix_to_clipboard:
            clean = resp.text.replace("```python", "").replace("```javascript", "").replace("```", "").strip()
            pyperclip.copy(clean)
            return "Fixed code clipboard par copy kar diya hai, aap seedha Ctrl+V kar sakte hain."
        return resp.text
    except Exception as e:
        return f"Screen vision error: {e}"
    finally:
        if os.path.exists(snap_path):
            os.remove(snap_path)

# ==========================================================
# 7. TTS & SHUTDOWN
# ==========================================================
tts_lock = threading.Lock()
winmm = ctypes.windll.winmm

def _mci(cmd: str) -> str:
    buf = ctypes.create_unicode_buffer(255)
    winmm.mciSendStringW(cmd, buf, 254, 0)
    return buf.value

def trigger_abort():
    STATE["abort"] = True
    try:
        _mci("stop jarvis_tts")
        _mci("close jarvis_tts")
    except Exception:
        pass
    if not STATE["shutdown"]:
        STATE["status"] = "STANDBY"
    add_log("🛑 [ABORTED] Task stopped!")

keyboard.add_hotkey("esc", trigger_abort)

async def _generate_tts_file(text: str, out_file: str, voice_id: str):
    communicate = edge_tts.Communicate(text, voice_id, rate="+8%", pitch="-4Hz")
    await communicate.save(out_file)

def speak(text: str, next_status: str = "STANDBY"):
    if (STATE["abort"] and not STATE["shutdown"]) or not text:
        if not STATE["shutdown"]:
            STATE["status"] = next_status
        return
    STATE["status"] = "SPEAKING"
    with tts_lock:
        out_file = os.path.join(BASE_DIR, f"tts_{int(time.time() * 1000)}.mp3")
        try:
            clean_text = text.replace("*", "").replace("#", "").replace("`", "").strip()
            spoken_text = clean_text[:360]
            asyncio.run(_generate_tts_file(spoken_text, out_file, STATE["voice"]))
            if os.path.exists(out_file):
                _mci("close jarvis_tts")
                _mci(f'open "{out_file}" type mpegvideo alias jarvis_tts')
                _mci("play jarvis_tts")
                while True:
                    if (STATE["abort"] and not STATE["shutdown"]) or _mci("status jarvis_tts mode") != "playing":
                        _mci("stop jarvis_tts")
                        break
                    time.sleep(0.08)
                _mci("close jarvis_tts")
        except Exception as e:
            print("TTS Error:", e)
        finally:
            if os.path.exists(out_file):
                try: os.remove(out_file)
                except Exception: pass
    if not STATE["shutdown"]:
        STATE["status"] = next_status

def shutdown_jarvis_and_llama():
    if STATE["shutdown"]: return
    STATE["shutdown"] = True
    STATE["wake_active"] = False
    STATE["status"] = "SHUTDOWN"
    add_log("🔻 [SHUTDOWN] Unloading models & saving local DB...")

    for model in OLLAMA_MODELS + [EMBED_MODEL]:
        subprocess.run(f"ollama stop {model}", shell=True, capture_output=True)
    for proc_name in ["llama-server.exe", "ollama_llama_server.exe"]:
        subprocess.run(f'taskkill /F /IM "{proc_name}" /T', shell=True, capture_output=True)

    speak("All memories saved locally. Goodbye Nayan!", next_status="SHUTDOWN")
    time.sleep(0.5)
    os._exit(0)

# ==========================================================
# 8. HYBRID AI BRAIN (ALL TOOLS CONNECTED IN BOTH CLOUD & OLLAMA)
# ==========================================================
client = genai.Client(
    api_key=GEMINI_API_KEY,
    http_options=types.HttpOptions(timeout=14000, retry_options=types.HttpRetryOptions(attempts=1)),
)

my_tools = [
    send_whatsapp_message,
    read_gmail_digest,
    search_and_sync_google_drive,
    index_pdpu_notes,
    search_pdpu_notes,
    save_memory,
    recall_memory,
    check_pdpu_exams_and_schedule,
    activate_workflow_mode,
    play_youtube,
    open_or_close_app,
    keyboard_and_mouse_action,
    create_file_or_folder,
    run_windows_command,
    analyze_or_fix_screen,
]

def run_with_powerful_ollama(user_input: str, history_msgs: list, saved_facts: str) -> str:
    active_model = STATE["active_ollama_model"]
    add_log(f"🧠 [LOCAL AI: {active_model}] Processing...")

    rag_context = ""
    if any(k in user_input.lower() for k in ["note", "notes", "pdf", "ppt", "slide", "unit", "syllabus", "chapter", "lecture", "formula", "exam"]):
        rag_context = search_pdpu_notes(user_input, top_k=3)

    sys_prompt = f"""You are JARVIS, Nayan's autonomous Windows OS controller & AI assistant.
Saved Facts & Contacts: {saved_facts}
{rag_context}

Respond ONLY in valid JSON with one of these exact actions:
- General question/study/chat: {{"action": "chat", "args": {{}}, "reply": "clear Hinglish answer"}}
- Send WhatsApp message: {{"action": "whatsapp", "args": {{"phone_or_name": "...", "message": "..."}}, "reply": ""}}
- Read unread Gmail: {{"action": "gmail", "args": {{}}, "reply": ""}}
- Sync Google Drive notes: {{"action": "drive_sync", "args": {{"query_topic": "..."}}, "reply": ""}}
- Index local PDF/PPT notes: {{"action": "index_notes", "args": {{}}, "reply": ""}}
- Save memory/contact: {{"action": "save_memory", "args": {{"topic": "...", "detail": "..."}}, "reply": ""}}
- PDEU schedule/exams: {{"action": "pdpu_status", "args": {{}}, "reply": ""}}
- Workflow mode (study/dev/chill): {{"action": "workflow", "args": {{"mode_name": "study|dev|chill"}}, "reply": ""}}
- YouTube play: {{"action": "play_youtube", "args": {{"song_or_video": "..."}}, "reply": ""}}
- Open/Close app: {{"action": "app_control", "args": {{"app_name": "...", "mode": "open|close"}}, "reply": ""}}
- GUI type/hotkey/scroll: {{"action": "gui", "args": {{"action_type": "type|hotkey|scroll", "value": "..."}}, "reply": ""}}
- Create file/folder: {{"action": "create_item", "args": {{"name": "...", "is_folder": false, "content": "..."}}, "reply": ""}}
- PowerShell: {{"action": "shell", "args": {{"ps_command": "..."}}, "reply": ""}}"""

    messages = [{"role": "system", "content": sys_prompt}] + history_msgs[-6:] + [{"role": "user", "content": user_input}]

    try:
        res = requests.post(
            "http://localhost:11434/api/chat",
            json={
                "model": active_model,
                "messages": messages,
                "format": "json",
                "stream": False,
                "keep_alive": -1,
                "options": {"temperature": 0.25, "num_ctx": 2048},
            },
            timeout=90,
        )
        if res.status_code == 200:
            raw_content = res.json()["message"]["content"]
            if "<think>" in raw_content and "</think>" in raw_content:
                raw_content = raw_content.split("</think>")[-1].strip()
            data = json.loads(raw_content)
            act = data.get("action", "chat")
            args = data.get("args", {})

            if act == "whatsapp": return send_whatsapp_message(args.get("phone_or_name", ""), args.get("message", ""))
            elif act == "gmail": return read_gmail_digest()
            elif act == "drive_sync": return search_and_sync_google_drive(args.get("query_topic", user_input))
            elif act == "index_notes": return index_pdpu_notes()
            elif act == "save_memory": return save_memory(args.get("topic", "Note"), args.get("detail", user_input))
            elif act == "pdpu_status": return check_pdpu_exams_and_schedule()
            elif act == "workflow": return activate_workflow_mode(args.get("mode_name", "study"))
            elif act == "play_youtube": return play_youtube(args.get("song_or_video", user_input))
            elif act == "app_control": return open_or_close_app(args.get("app_name", ""), args.get("mode", "open"))
            elif act == "gui": return keyboard_and_mouse_action(args.get("action_type", "hotkey"), args.get("value", ""))
            elif act == "create_item": return create_file_or_folder(args.get("name", "note.txt"), args.get("is_folder", False), args.get("content", ""))
            elif act == "shell": return run_windows_command(args.get("ps_command", ""))
            return data.get("reply", "Main ready hoon Nayan, bataiye!")
    except Exception as e:
        add_log(f"⚠️ Local Ollama Error: {e}")
    return "Local model load ho raha hai, 5 second baad ek baar aur try karein."

def execute_ai_pipeline(cmd_text: str, use_voice_reply: bool = True):
    low = cmd_text.lower().strip()
    if any(p in low for p in ["exit jarvis", "jarvis exit", "quit jarvis", "jarvis band karo", "shutdown jarvis"]):
        add_log(f"NAYAN > {cmd_text}")
        log_chat_turn("user", cmd_text)
        shutdown_jarvis_and_llama()
        return

    STATE["abort"] = False
    STATE["status"] = "THINKING"
    add_log(f"NAYAN > {cmd_text}")
    log_chat_turn("user", cmd_text)

    if any(p in low for p in ["index notes", "index my notes", "scan notes", "notes scan karo"]):
        reply = index_pdpu_notes()
        log_chat_turn("assistant", reply)
        add_log(f"JARVIS > {reply}")
        if use_voice_reply: speak(reply, next_status="STANDBY")
        else: STATE["status"] = "STANDBY"
        return

    history_msgs = get_recent_chat_context(limit=10)
    saved_facts = recall_memory("")
    reply = ""

    if GEMINI_API_KEY and "AapkiAsliKey" not in GEMINI_API_KEY:
        convo_text = "\n".join([f"{m['role'].upper()}: {m['content']}" for m in history_msgs[:-1]])
        sys_msg = (
            "You are JARVIS MK-107, Nayan's autonomous Windows AI OS. "
            "Use `send_whatsapp_message` for WhatsApp, `read_gmail_digest` for Gmail, `search_and_sync_google_drive` for Google Drive notes, "
            "`search_pdpu_notes` for study questions, `activate_workflow_mode` for study/dev/chill modes, and OS tools for apps/GUI. "
            "Reply in crisp Hinglish.\n"
            f"Saved Facts: {saved_facts}\nRecent Chat:\n{convo_text}"
        )
        for model_name in ["gemini-3.8-flash-lite", "gemini-3.8-flash"]:
            if STATE["abort"]: break
            try:
                chat = client.chats.create(
                    model=model_name,
                    config=types.GenerateContentConfig(
                        system_instruction=sys_msg,
                        tools=my_tools,
                        temperature=0.25,
                    ),
                )
                response = chat.send_message(cmd_text)
                reply = response.text
                break
            except Exception:
                continue

    if not reply and not STATE["abort"]:
        reply = run_with_powerful_ollama(cmd_text, history_msgs[:-1], saved_facts)

    if not STATE["abort"] and reply:
        log_chat_turn("assistant", reply)
        add_log(f"JARVIS > {reply}")
        if use_voice_reply:
            speak(reply, next_status="STANDBY")
        else:
            STATE["status"] = "STANDBY"
    else:
        STATE["status"] = "STANDBY"

# ==========================================================
# 9. VOICE BIOMETRICS & WAKE-WORD ENGINE
# ==========================================================
def extract_voice_features(audio_data: np.ndarray):
    audio = audio_data.flatten().astype(np.float32)
    non_silent = np.abs(audio) > 0.015
    if np.any(non_silent): audio = audio[non_silent]
    if len(audio) < 400: return None
    audio = audio / (np.max(np.abs(audio)) + 1e-8)
    fft_vals = np.abs(np.fft.rfft(audio, n=4096))
    bands = np.array_split(fft_vals[:1000], 40)
    features = np.array([np.mean(np.log1p(b)) for b in bands])
    return features / (np.linalg.norm(features) + 1e-8)

def verify_speaker(audio_data: np.ndarray, threshold: float = 0.83) -> bool:
    if not os.path.exists(VOICEPRINT_FILE): return True
    saved_fp = np.load(VOICEPRINT_FILE)
    current_fp = extract_voice_features(audio_data)
    if current_fp is None: return False
    sim = float(np.dot(saved_fp, current_fp))
    add_log(f"[BIOMETRIC] Voice Match: {sim * 100:.0f}%")
    return sim >= threshold

def capture_and_run_command():
    if STATE["status"] == "LISTENING" or STATE["shutdown"]: return
    try:
        STATE["abort"] = False
        STATE["status"] = "LISTENING"
        add_log("🎙️ Active! Command ya sawal bolein (Max 18s)...")
        chunk_duration = 0.4
        chunk_samples = int(chunk_duration * SAMPLE_RATE)
        recorded_chunks, speech_started, silent_chunks = [], False, 0

        with sd.InputStream(samplerate=SAMPLE_RATE, channels=1, dtype="float32") as stream:
            for _ in range(int(MAX_LISTEN_SECONDS / chunk_duration)):
                if STATE["abort"] or STATE["shutdown"]:
                    STATE["status"] = "STANDBY"
                    return
                chunk, _ = stream.read(chunk_samples)
                recorded_chunks.append(chunk.copy())
                if np.max(np.abs(chunk)) > 0.025:
                    speech_started, silent_chunks = True, 0
                elif speech_started:
                    silent_chunks += 1
                    if silent_chunks >= 6: break

        if STATE["abort"] or not speech_started:
            STATE["status"] = "STANDBY"
            return

        rec = np.concatenate(recorded_chunks, axis=0)
        if not verify_speaker(rec):
            msg = "Access Denied. Yeh Nayan ki aawaaz nahi hai."
            add_log(f"🚨 {msg}")
            speak(msg, next_status="STANDBY")
            return

        wav_path = os.path.join(BASE_DIR, "temp_cmd.wav")
        write(wav_path, SAMPLE_RATE, np.int16(rec * 32767))
        recognizer = sr.Recognizer()
        with sr.AudioFile(wav_path) as source:
            audio = recognizer.record(source)
        text = recognizer.recognize_google(audio, language="en-IN")
        execute_ai_pipeline(text, use_voice_reply=True)
    except Exception:
        STATE["status"] = "STANDBY"

def background_wake_word_loop():
    recognizer = sr.Recognizer()
    chunk_samples = int(2.4 * SAMPLE_RATE)
    wake_wav = os.path.join(BASE_DIR, "temp_wake.wav")
    while not STATE["shutdown"]:
        if not STATE["wake_active"] or STATE["status"] != "STANDBY":
            time.sleep(0.35)
            continue
        try:
            rec = sd.rec(chunk_samples, samplerate=SAMPLE_RATE, channels=1, dtype="float32")
            sd.wait()
            if STATE["status"] != "STANDBY" or STATE["shutdown"] or np.max(np.abs(rec)) < 0.03:
                continue
            write(wake_wav, SAMPLE_RATE, np.int16(rec * 32767))
            with sr.AudioFile(wake_wav) as source:
                audio = recognizer.record(source)
            heard = recognizer.recognize_google(audio, language="en-IN").lower()
            if any(w in heard for w in ["hey", "jarvis", "hi jarvis", "ok jarvis"]):
                add_log(f"⚡ [WAKE-WORD] Heard: '{heard}'")
                cleaned = heard.replace("hey jarvis", "").replace("jarvis", "").replace("hey", "").strip()
                if len(cleaned) > 4:
                    execute_ai_pipeline(cleaned, use_voice_reply=True)
                else:
                    speak("Yes Nayan, boliye.", next_status="LISTENING")
                    capture_and_run_command()
        except Exception:
            pass

# ==========================================================
# 10. FASTAPI ENDPOINTS
# ==========================================================
class CommandReq(BaseModel):
    command: str
    voice_reply: bool = True

@app.get("/", response_class=HTMLResponse)
def serve_hud():
    with open(os.path.join(BASE_DIR, "hud.html"), "r", encoding="utf-8") as f:
        return f.read()

@app.get("/api/telemetry")
def get_telemetry():
    bat = psutil.sensors_battery()
    disk = psutil.disk_usage("C:\\")
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT topic, detail FROM memories ORDER BY id DESC LIMIT 4")
    mem_list = [{"topic": r[0], "detail": r[1]} for r in c.fetchall()]
    c.execute("SELECT COUNT(*) FROM study_chunks")
    rag_chunks = c.fetchone()[0]
    conn.close()

    return {
        "cpu": psutil.cpu_percent(interval=None),
        "ram": psutil.virtual_memory().percent,
        "disk": disk.percent,
        "battery": bat.percent if bat else 100,
        "status": STATE["status"],
        "active_mode": f"{STATE['active_ollama_model'].upper()} ({rag_chunks} RAG)",
        "logs": STATE["logs"],
        "voice_enrolled": os.path.exists(VOICEPRINT_FILE),
        "voice_name": "INDIAN NEURAL" if "IN" in STATE["voice"] else "BRITISH JARVIS",
        "pdpu": get_pdpu_academic_status(),
        "memories": mem_list,
    }

@app.post("/api/command")
def handle_text_command(req: CommandReq):
    threading.Thread(target=execute_ai_pipeline, args=(req.command, req.voice_reply), daemon=True).start()
    return {"ok": True}

@app.post("/api/toggle-voice")
def toggle_voice_accent():
    STATE["voice"] = "en-GB-RyanNeural" if STATE["voice"] == "en-IN-PrabhatNeural" else "en-IN-PrabhatNeural"
    msg = "Voice profile switched."
    add_log(f"🔊 {msg}")
    threading.Thread(target=speak, args=(msg, "STANDBY"), daemon=True).start()
    return {"ok": True}

@app.post("/api/voice-listen")
def trigger_voice_listen():
    threading.Thread(target=capture_and_run_command, daemon=True).start()
    return {"ok": True}

@app.post("/api/abort")
def abort_api():
    trigger_abort()
    return {"ok": True}

if __name__ == "__main__":
    threading.Thread(target=ensure_ollama_running, daemon=True).start()
    threading.Thread(target=background_wake_word_loop, daemon=True).start()
    threading.Timer(1.5, lambda: webbrowser.open("http://localhost:8000")).start()
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="warning")