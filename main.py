import keyboard
import threading
import time
import pyautogui
import pygetwindow as gw
import pyperclip
import requests
import difflib
from collections import deque
import customtkinter as ctk
import os
import sys

# --- Config ---
GITHUB_PHRASES_URL = "https://raw.githubusercontent.com/akaBrooklyn/grammar-tool-config/refs/heads/main/allowed_phrases.txt"
VERSION = "1.0.0"
VERSION_URL = "https://raw.githubusercontent.com/akaBrooklyn/grammar-tool-config/refs/heads/main/main.py"

# --- Globals ---
typed_chars = []
phrase_buffer = deque(maxlen=10)
recent_phrases = deque(maxlen=50)
allowed_phrases = []
popup = None

def download_allowed_phrases():
    try:
        response = requests.get(GITHUB_PHRASES_URL)
        if response.ok:
            return [line.strip() for line in response.text.splitlines() if line.strip()]
    except Exception as e:
        print(f"Error loading phrases: {e}")
    return []

def is_typing_window():
    win = gw.getActiveWindow()
    return win and win.title and "pycharm" not in win.title.lower()

def show_suggestions(suggestions, original_phrase):
    global popup
    if popup:
        popup.destroy()

    popup = ctk.CTk()
    popup.geometry("300x200+100+100")
    popup.title("Suggestions")

    def replace_text(suggestion):
        for _ in range(len(original_phrase)):
            pyautogui.press("backspace")
        pyautogui.typewrite(suggestion)
        popup.destroy()

    for suggestion in suggestions:
        btn = ctk.CTkButton(popup, text=suggestion, command=lambda s=suggestion: replace_text(s))
        btn.pack(pady=2)

    popup.mainloop()

def listen_keys():
    global allowed_phrases

    while True:
        event = keyboard.read_event()
        if event.event_type == keyboard.KEY_DOWN:
            char = event.name
            if len(char) == 1:
                typed_chars.append(char)
            elif char == "space" or char == "enter":
                phrase = ''.join(typed_chars).strip()
                typed_chars.clear()
                if not phrase:
                    continue
                phrase_buffer.append(phrase)
                full_phrase = ' '.join(phrase_buffer)
                if full_phrase in recent_phrases:
                    continue
                matches = [p for p in allowed_phrases if SequenceMatcher(None, full_phrase.lower(), p.lower()).ratio() > 0.7]
                if matches:
                    recent_phrases.append(full_phrase)
                    show_suggestions(matches, full_phrase)

def SequenceMatcher(a, b):
    return difflib.SequenceMatcher(None, a, b)

def main():
    print("INFO:GrammarTool:Listening for keys...")
    global allowed_phrases
    allowed_phrases = download_allowed_phrases()
    threading.Thread(target=listen_keys, daemon=True).start()
    while True:
        time.sleep(1)

if __name__ == "__main__":
    main()
