# .\activar_jarvis.ps1
"""
✅ Comandos recargados.
✔️ Usando micrófono: [0] Asignador de sonido Microsoft - Input
🎤 Voice Activity Detector running...
✅ Reconocido: hey jarvis
⚠️ Muy poco audio grabado. Ignorando...
❌ No entendí lo que dijiste.
❌ No entendí lo que dijiste.
❌ No entendí lo que dijiste.
"""
# --- Librerías estándar ---
import os
import time
import math
import random
import struct
import datetime
import collections
import unicodedata
import threading

# --- Librerías de terceros ---
import numpy as np
import pandas as pd
import pyttsx3
import pyaudio
import webrtcvad
import speech_recognition as sr
from pydub import AudioSegment
from pydub.playback import play

# --- Librerías de GUI (Tkinter) ---
import tkinter as tk
from tkinter import ttk
import itertools

# Crear ventana
root = tk.Tk()
root.title("Jarvis Assistant")
root.geometry("700x700")
root.configure(bg="#1e1e1e")  # Fondo oscuro estilo futurista

# --- EMOJI estado de comandos ---
command_status_label = tk.Label(root, text="❌", font=("Segoe UI", 10), bg="#1e1e1e", fg="white")
command_status_label.place(x=10, y=10)

# --- CÍRCULO estado de Jarvis ---
canvas = tk.Canvas(root, width=20, height=20, highlightthickness=0, bg="#1e1e1e")
canvas.place(relx=0.5, y=10, anchor="n")
indicator = canvas.create_oval(1, 1, 19, 19, fill="gray", outline="gray")

# --- BARRAS animadas estilo futurista ---
bar_canvas = tk.Canvas(root, width=500, height=80, bg="#1e1e1e", highlightthickness=0)
bar_canvas.place(relx=0.5, rely=0.5, anchor="center")

bar_count = 78
bar_ids = []
bar_width = 2
spacing = 4
bar_height = 40

# Calcular posición inicial para centrar las barras horizontalmente
total_width = (bar_width + spacing) * bar_count
start_x = (500 - total_width) // 2  # centrar dentro del canvas de 500px

# Función para generar color degradado azul -> cyan
def get_gradient_color(i, total):
    start_color = (0, 180, 255)  # Azul claro
    end_color = (0, 255, 180)    # Verde-cyan
    r = int(start_color[0] + (end_color[0] - start_color[0]) * i / total)
    g = int(start_color[1] + (end_color[1] - start_color[1]) * i / total)
    b = int(start_color[2] + (end_color[2] - start_color[2]) * i / total)
    return f'#{r:02x}{g:02x}{b:02x}'

# Crear barras con colores degradados
for i in range(bar_count):
    x = start_x + i * (bar_width + spacing)
    color = get_gradient_color(i, bar_count - 1)
    bar = bar_canvas.create_rectangle(x, bar_height, x + bar_width, bar_height + 20, fill=color, outline=color)
    bar_ids.append(bar)



def set_indicator_color(color):
    canvas.itemconfig(indicator, fill=color, outline=color)
"""
set_indicator_color("lightblue")   # Por ejemplo, cuando se activa
set_indicator_color("gray")        # Cuando está dormido
set_indicator_color("yellow")      # Cuando habla
set_indicator_color("white")       # Estado neutro
set_indicator_color("lightgreen")  # Voz detectada
"""

# Funcion para animar las barras# --- ANIMACIÓN de barras ---
bar_animation_running = False

def animate_bars(active=True):
    global bar_animation_running

    if active and not bar_animation_running:
        bar_animation_running = True
        def run():
            cycle = itertools.cycle(range(10, 30, 5))
            while bar_animation_running:
                for i, bar_id in enumerate(bar_ids):
                    height = random.randint(10, 30)
                    x1, _, x2, _ = bar_canvas.coords(bar_id)
                    bar_canvas.coords(bar_id, x1, 50 - height, x2, 50)
                time.sleep(0.1)
        threading.Thread(target=run, daemon=True).start()
    
    elif not active:
        bar_animation_running = False
        # Volver a su posición original
        for i, bar_id in enumerate(bar_ids):
            x1, _, x2, _ = bar_canvas.coords(bar_id)
            bar_canvas.coords(bar_id, x1, 40, x2, 50)

# --- FUNCIÓN para actualizar emoji del estado de comandos ---
def update_command_status(success):
    if success:
        command_status_label.config(text="✅", fg="lime")
    else:
        command_status_label.config(text="❌", fg="red")

# --- BOTÓN para recargar comandos ---
def reload_commands():
    global user_input, reply_jarvis
    try:
        user_input = pd.read_excel("C:/Jarvis/input.xlsx", sheet_name="User")
        reply_jarvis = pd.read_excel("C:/Jarvis/input.xlsx", sheet_name="Replies")
        update_command_status(True)
        #print("✅ Comandos recargados")
    except Exception as e:
        update_command_status(False)
        print(f"❌ Error al recargar comandos: {e}")

reload_button = tk.Button(root, text="🔄 Reload Commands", command=reload_commands,
                          font=("Segoe UI", 10), bg="#2c2c2c", fg="white")
reload_button.place(relx=0.5, rely=0.9, anchor="center")  # abajo centrado

#Funcion para música
def play_from(file_path, start_time_sec):
    sound = AudioSegment.from_wav(file_path)
    # Cortar el audio desde el tiempo especificado (en milisegundos)
    cropped = sound[start_time_sec * 1000:]
    play(cropped)
# Reproducir el audio en un hilo separado
def play_from_async(file_path, start_time_sec):
    threading.Thread(target=play_from, args=(file_path, start_time_sec), daemon=True).start()


# CODIGO MAIN  ------------------------------------------------------------------------------------
# --- Estado global ---
is_sleeping = True
custom_commands = {}
engine_lock = threading.Lock()
engine = pyttsx3.init()
engine.setProperty('rate', 170)

# --- Configuración de voz global ---
voices = engine.getProperty('voices')
for voice in voices:
    if "Zira" in voice.name or "David" in voice.name:
        engine.setProperty('voice', voice.id)
        break

# --- Función para cargar comandos desde Excel ---
def load_commands_from_excel(path):
    try:
        df_user = pd.read_excel(path, sheet_name="User", header=None)
        df_replies = pd.read_excel(path, sheet_name="Replies", header=None)

        global user_input
        user_input = df_user

        global reply_jarvis
        reply_jarvis = df_replies

        #print("✅ Comandos recargados.")

    except Exception as e:
        print(f"❌ Error loading Excel commands: {e}")
        return {}

custom_commands = load_commands_from_excel("C:/Jarvis/input.xlsx")

def normalize_text(text):
    if not isinstance(text, str):
        text = str(text)
    text = unicodedata.normalize("NFKD", text).encode("ASCII", "ignore").decode("utf-8")
    return text.strip().lower()

def find_word_in_user_input(word):
    set_indicator_color("yellow")      # Cuando habla
    word = normalize_text(word)
    for row_idx, row in user_input.iterrows():
        for col_idx, cell in enumerate(row):
            if pd.notna(cell):
                cell_text = normalize_text(cell)
                if cell_text in word:
                    return row_idx
    return None

# --- Sonidos personalizados ---
def play_effect_wav(path):
    try:
        import winsound
        winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)
    except Exception as e:
        print(f"Error al reproducir sonido: {e}")

# --- Frases y colores ---
jarvis_responses = [
    "At your service, sir.",
    "All systems functioning within normal parameters.",
    "Engaging now.",
    "Standing by.",
    "Online and ready.",
    "How may I assist you today?"
]
jarvis_responses_stanby = [
    "I'm here, sir.",
    "Awaiting your command.",
    "Ready when you are.",
    "Just say the word.",
    "Listening for your instructions."
]

def Speak(text=None):
    def run_and_animate():
        with engine_lock:
            set_indicator_color("yellow")      # Cuando habla
            animate_bars()
            response = text if text else random.choice(jarvis_responses)
            animate_bars(True)
            engine.say(response)
            engine.runAndWait()
            set_indicator_color("white")       # Estado neutro
            animate_bars(False)
    threading.Thread(target=run_and_animate).start()

# --- Detectar micrófono ---
def get_available_microphone():
    mic_list = sr.Microphone.list_microphone_names()
    for index, name in enumerate(mic_list):
        try:
            with sr.Microphone(device_index=index) as source:
                return index, name
        except Exception:
            continue
    raise RuntimeError("No se encontró micrófono disponible.")

def process_text_command(data):
    global is_sleeping

    row_idx = find_word_in_user_input(data)
    if row_idx is not None:
        rep1 = reply_jarvis.loc[row_idx]
        rep2 = [str(x).strip() for x in rep1 if pd.notna(x)]
        Speak(random.choice(rep2))
        return

    if "shutdown" in data or "shut down" in data or "bye-bye" in data or "turn off" in data or "power off" in data or "bye bye" in data or "goodbye" in data or "byebye" in data:
        animate_bars(False)
        play_effect_wav("C:/Jarvis/Effects/turn_off.wav")
        Speak("System powering down. See you next time.")
        set_indicator_color("black")      # Cuando se apaga
        time.sleep(4)
        animate_bars(True)
        set_indicator_color("black")      # Cuando se apaga
        play_from_async("C:/Jarvis/NGVUP.wav", 85)
        set_indicator_color("black")      # Cuando se apaga
        time.sleep(9)
        set_indicator_color("black")      # Cuando se apaga
        os._exit(0)
    elif "sleep mode" in data:
        Speak("Entering sleep mode. Say 'Hey Jarvis' to wake me.")
        set_indicator_color("gray")        # Cuando está dormido
        is_sleeping = True
    elif "cancel" in data or "never mind" in data:
        Speak("Cancelling voice command.")
    elif "what time is it" in data:
        current_time = datetime.datetime.now().strftime("%H:%M")
        Speak(f"The time is {current_time}.")
    elif "what day is it" in data:
        current_day = datetime.datetime.now().strftime("%A")
        Speak(f"Today is {current_day}.")
    elif "what is the date" in data:
        current_date = datetime.datetime.now().strftime("%Y-%m-%d")
        Speak(f"Today's date is {current_date}.")
    elif "what is your name" in data or "who are you" in data:
        Speak("I am Jarvis, your personal assistant.")
    elif any(k in data for k in ["jarvis", "hey jarvis", "harvis", "jarv"]):
        Speak(random.choice(jarvis_responses_stanby))
    else:
        Speak("I'm sorry, I didn't understand that command.")
    set_indicator_color("lightblue")   # Por ejemplo, cuando se activa
try:
    mic_index, mic_name = get_available_microphone()
    #print(f"✔️ Usando micrófono: [{mic_index}] {mic_name}")
except RuntimeError as e:
    print(str(e))
    exit()

r = sr.Recognizer()
source = sr.Microphone(device_index=mic_index)

class VoiceActivityDetector(threading.Thread):
    def __init__(self, aggressiveness=3):
        super().__init__(daemon=True)
        self.vad = webrtcvad.Vad(aggressiveness)
        self.running = True
        self.pa = pyaudio.PyAudio()
        self.stream = self.pa.open(
            format=pyaudio.paInt16,
            channels=1,
            rate=16000,
            input=True,
            frames_per_buffer=320,
        )

    def run(self):
        #print("🎤 Voice Activity Detector running...")

        while self.running:
            frame = self.stream.read(320, exception_on_overflow=False)
            if self.vad.is_speech(frame, 16000):
                self.capture_and_recognize()
            elif is_sleeping:
                set_indicator_color("gray")
                animate_bars(False)  # Detener animación
                self.capture_and_recognize()

    def capture_and_recognize(self):
        global is_sleeping
        ring_buffer = collections.deque(maxlen=20)
        triggered = False
        voiced_frames = []
        silence_count = 0
        set_indicator_color("lightgreen")  # Voz detectada
        animate_bars(True)  # Cuando detecta voz del usuario  

        for _ in range(100):  # hasta ~2 segundos de espera
            frame = self.stream.read(320, exception_on_overflow=False)
            is_speech = self.vad.is_speech(frame, 16000)

            if not triggered:
                ring_buffer.append((frame, is_speech))
                num_voiced = len([f for f, speech in ring_buffer if speech])
                if num_voiced > 0.6 * ring_buffer.maxlen:
                    triggered = True
                    voiced_frames.extend([f for f, s in ring_buffer])
                    ring_buffer.clear()
            else:
                voiced_frames.append(frame)
                if not is_speech:
                    silence_count += 1
                else:
                    silence_count = 0
                if silence_count > 10:  # espera más silencio antes de cortar
                    break          
        if len(voiced_frames) < 20:  # menos de ~0.4s de voz
            #print("⚠️ Muy poco audio grabado. Ignorando...")
            set_indicator_color("white")  # Cuando está dormido
            return

        audio_data = b"".join(voiced_frames)
        
        audio = sr.AudioData(audio_data, 16000, 2)
        animate_bars(False)
        try:
            set_indicator_color("yellow")  # Voz detectada
            text = r.recognize_google(audio).strip().lower()
            print("✅ Reconocido:", text)

            if is_sleeping:
                set_indicator_color("gray")
                animate_bars(False)
                if any(k in text for k in ["jarvis", "hey jarvis", "harvis", "jarv"]):
                    is_sleeping = False
                    hour = datetime.datetime.now().hour
                    if hour < 12:
                        Speak("Good morning, sir.")
                        Speak()
                    elif hour < 18:
                        Speak("Good afternoon, sir.")
                        Speak()
                    else:
                        Speak("Good evening, sir.")
                        Speak()
                return
            process_text_command(text)
            time.sleep(0.5)
            set_indicator_color("white")  # Estado neutro

        except sr.UnknownValueError:
            #print("❌ No entendí lo que dijiste.")
            pass
        except sr.RequestError as e:
            print(f"❌ Error en reconocimiento: {e}")
            pass


    def stop(self):
        self.running = False
        self.stream.stop_stream()
        self.stream.close()
        self.pa.terminate()

# Iniciar el detector de voz
def start_jarvis():
    play_effect_wav("C:/Jarvis/Effects/Startup_effect.wav")
    vad_thread = VoiceActivityDetector()
    vad_thread.start()

start_jarvis()

# Iniciar GUI
root.mainloop()
