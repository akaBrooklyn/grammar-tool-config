import keyboard
import threading
import time
import customtkinter as ctk
import difflib
import pyautogui
import pygetwindow as gw
import pyperclip
import re
from collections import deque
from difflib import SequenceMatcher

# --- Globals ---
last_focused_window = None
typed_chars = []
phrase_buffer = deque(maxlen=10)
recent_phrases = deque(maxlen=50)
suggestion_active = False
listener_running = False
WORD_BOUNDARIES = {'space', 'enter', 'tab', '.', ',', '?', '!', ';', ':'}

# --- Normalize and Prepare Text ---
def normalize_text(text):
    text = text.lower()
    text = re.sub(r"[-_']", " ", text)
    text = re.sub(r"[^\w\s]", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

# --- Grammar Engine ---
class GrammarEngine:
    def __init__(self, keywords):
        self.keywords = [normalize_text(k) for k in keywords]
        self.original_map = {normalize_text(k): k for k in keywords}

    def check_phrase(self, phrase):
        norm = normalize_text(phrase)
        scored = sorted(
            [(k, SequenceMatcher(None, norm, k).ratio()) for k in self.keywords if k != norm],
            key=lambda x: x[1],
            reverse=True
        )
        return [self.original_map[k] for k, _ in scored[:10]]

# --- Suggestion Popup ---
class SuggestionPopup(ctk.CTkToplevel):
    def __init__(self, phrase, suggestions, callback):
        super().__init__()
        self.phrase = phrase
        self.callback = callback
        self.title("Suggestion")
        self.attributes("-topmost", True)
        self.geometry(self._center_geometry(300, 50 + 40 * len(suggestions)))
        self.resizable(False, False)

        label = ctk.CTkLabel(self, text=f"Correction for: '{phrase}'")
        label.pack(pady=10)

        for sug in suggestions:
            btn = ctk.CTkButton(self, text=sug, command=lambda s=sug: self.select(s))
            btn.pack(pady=3)

        self.after(8000, self.destroy)

    def _center_geometry(self, w, h):
        screen_width, screen_height = pyautogui.size()
        x = (screen_width - w) // 2
        y = (screen_height - h) // 2
        return f"{w}x{h}+{x}+{y}"

    def select(self, correction):
        self.callback(self.phrase, correction)
        self.destroy()

# --- Correction Handler ---
def apply_correction(original, correction):
    global last_focused_window, suggestion_active
    try:
        print(f"[Correction] Replacing '{original}' → '{correction}'")
        time.sleep(0.2)

        if last_focused_window:
            win = gw.getWindowsWithTitle(last_focused_window)
            if win:
                win[0].activate()
                time.sleep(0.4)

        pyperclip.copy(correction + ' ')

        # Select and delete entire phrase (ensures no characters are left behind)
        for _ in range(len(original)):
            pyautogui.press('backspace')
            time.sleep(0.001)
        pyautogui.press('backspace')  # Extra to ensure removal of one stray char

        # Paste the correct phrase
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.05)

        typed_chars.clear()
        phrase_buffer.clear()
        suggestion_active = False
        print("[Replacement done]")

    except Exception as e:
        print(f"[Error] Correction failed: {e}")

# --- Keyboard Listener ---
def on_key_press(event):
    global suggestion_active

    name = event.name
    if len(name) == 1 and name.isprintable():
        typed_chars.append(name)
    elif name in WORD_BOUNDARIES:
        if typed_chars:
            word = ''.join(typed_chars).strip()
            typed_chars.clear()
            if word:
                phrase_buffer.append(word)
                suggestion_active = False
                check_combinations()
    elif name == 'backspace' and typed_chars:
        typed_chars.pop()

def start_keyboard_listener():
    global listener_running
    if not listener_running:
        listener_running = True
        keyboard.on_press(on_key_press)

# --- Phrase Handling ---
def check_combinations():
    for n in range(4, 0, -1):
        if len(phrase_buffer) >= n:
            phrase = ' '.join(list(phrase_buffer)[-n:])
            norm_phrase = normalize_text(phrase)
            if norm_phrase not in recent_phrases:
                on_phrase_completed(phrase)

def on_phrase_completed(phrase):
    global suggestion_active
    if suggestion_active:
        return

    suggestions = ge.check_phrase(phrase)
    if suggestions:
        norm_phrase = normalize_text(phrase)
        recent_phrases.append(norm_phrase)
        suggestion_active = True
        print(f"[Match] '{phrase}' → {suggestions}")

        try:
            active = gw.getActiveWindow()
            if active:
                global last_focused_window
                last_focused_window = active.title
        except:
            last_focused_window = None

        threading.Thread(
            target=SuggestionPopup,
            args=(phrase, suggestions, apply_correction),
            daemon=True
        ).start()

# --- Keyword Loader ---
def load_keywords_from_file(filename):
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            return [line.strip() for line in f if line.strip()]
    except Exception as e:
        print(f"[Error] Loading keywords: {e}")
        return []

# --- Main ---
def main():
    global ge
    keywords = load_keywords_from_file("keywords.txt")
    if not keywords:
        print("[Error] No keywords loaded.")
        return

    ge = GrammarEngine(keywords)
    start_keyboard_listener()

    root = ctk.CTk()
    root.withdraw()
    root.mainloop()

if __name__ == "__main__":
    main()
