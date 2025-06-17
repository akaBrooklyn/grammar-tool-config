import keyboard
import threading
import time
import customtkinter as ctk
import difflib
import pyautogui
import pygetwindow as gw
import pyperclip
import re
import json
import logging
import os
import pystray
from collections import deque
from difflib import SequenceMatcher
from PIL import Image
from typing import List, Dict, Optional, Deque, Tuple

# --- Constants ---
APP_NAME = "GrammarPal"
VERSION = "1.0.1"
CONFIG_FILE = "config.json"
DEFAULT_KEYWORDS_FILE = "keywords.txt"
WORD_BOUNDARIES = {'space', 'enter', 'tab', '.', ',', '?', '!', ';', ':', '\n'}
MIN_SUGGESTIONS = 20
MAX_SUGGESTIONS = 20


# --- Setup Logging ---
def setup_logging():
    log_dir = "logs"
    os.makedirs(log_dir, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(os.path.join(log_dir, f"{APP_NAME.lower()}.log")),
            logging.StreamHandler()
        ]
    )


# --- Configuration Manager ---
class ConfigManager:
    def __init__(self):
        self.config = self._load_config()

    def _load_config(self) -> Dict:
        default_config = {
            "keywords_file": DEFAULT_KEYWORDS_FILE,
            "suggestion_timeout": 8,
            "max_phrase_length": 10,
            "recent_phrases_size": 50,
            "min_similarity": 0.5,
            "theme": "dark",
            "start_minimized": True,
            "enable_partial_matching": True
        }

        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    loaded = json.load(f)
                    return {**default_config, **loaded}
        except Exception as e:
            logging.error(f"Error loading config: {e}")

        return default_config

    def save_config(self):
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            logging.error(f"Error saving config: {e}")

    def get(self, key: str, default=None):
        return self.config.get(key, default)

    def set(self, key: str, value):
        self.config[key] = value
        self.save_config()


# --- Grammar Engine ---
class GrammarEngine:
    def __init__(self, keywords: List[str]):
        self.keywords = [self.normalize_text(k) for k in keywords]
        self.original_map = {self.normalize_text(k): k for k in keywords}
        self.word_index = self.build_word_index()

    @staticmethod
    def normalize_text(text: str) -> str:
        text = text.lower()
        text = re.sub(r"[-_']", " ", text)
        text = re.sub(r"[^\w\s]", "", text)
        text = re.sub(r"\s+", " ", text)
        return text.strip()

    def build_word_index(self) -> Dict[str, List[str]]:
        """Create an index of individual words to their full phrases"""
        index = {}
        for phrase in self.keywords:
            words = phrase.split()
            for word in words:
                if word not in index:
                    index[word] = []
                if phrase not in index[word]:
                    index[word].append(phrase)
        return index

    def check_phrase(self, phrase: str, min_similarity: float = 0.5, partial_matching: bool = True) -> List[str]:
        norm_phrase = self.normalize_text(phrase)

        # Get matches through multiple methods
        all_matches = []

        # 1. Exact prefix matches (highest priority)
        prefix_matches = [
            (k, 1.0) for k in self.keywords
            if k.startswith(norm_phrase) and k != norm_phrase
        ]
        all_matches.extend(prefix_matches)

        # 2. Partial word matches (if enabled)
        if partial_matching:
            input_words = norm_phrase.split()
            if input_words:
                candidate_phrases = set()
                for word in input_words:
                    if word in self.word_index:
                        candidate_phrases.update(self.word_index[word])

                for candidate in candidate_phrases:
                    common_words = set(input_words) & set(candidate.split())
                    score = len(common_words) / len(input_words)
                    if score >= 0.5:
                        all_matches.append((candidate, score * 0.9))  # Slightly lower than exact matches

        # 3. Similarity matches
        similarity_matches = [
            (k, SequenceMatcher(None, norm_phrase, k).ratio())
            for k in self.keywords
            if k != norm_phrase
        ]
        all_matches.extend(similarity_matches)

        # Combine and deduplicate matches
        unique_matches = {}
        for phrase, score in all_matches:
            if phrase not in unique_matches or score > unique_matches[phrase]:
                unique_matches[phrase] = score

        # Filter by minimum similarity and sort
        filtered = [
            (phrase, score)
            for phrase, score in unique_matches.items()
            if score >= min_similarity
        ]
        sorted_matches = sorted(filtered, key=lambda x: (-x[1], len(x[0])))

        # Get original phrases for top matches
        results = [self.original_map[phrase] for phrase, _ in sorted_matches[:MAX_SUGGESTIONS]]

        # Ensure minimum suggestions
        if len(results) < MIN_SUGGESTIONS:
            remaining = [
                            self.original_map[k] for k in self.keywords
                            if self.original_map[k] not in results
                        ][:MIN_SUGGESTIONS - len(results)]
            results.extend(remaining)

        return results


# --- Suggestion Popup ---
class SuggestionPopup(ctk.CTkToplevel):
    def __init__(self, phrase: str, suggestions: List[str], callback, timeout: int = 8):
        super().__init__()
        self.phrase = phrase
        self.callback = callback
        self.title(f"{APP_NAME} - Suggestion")
        self.attributes("-topmost", True)
        self.geometry(self._center_geometry(400, 50 + 40 * min(len(suggestions), 10)))
        self.resizable(False, False)

        # Limit displayed phrase length
        display_phrase = phrase[:50] + ('...' if len(phrase) > 50 else '')

        label = ctk.CTkLabel(self, text=f"Did you mean instead of: '{display_phrase}'")
        label.pack(pady=10)

        # Create a scrollable frame for suggestions
        scroll_frame = ctk.CTkScrollableFrame(self, height=min(300, 35 * len(suggestions)))
        scroll_frame.pack(pady=5, padx=10, fill="both", expand=True)

        for sug in suggestions[:10]:  # Show max 10 suggestions
            btn = ctk.CTkButton(
                scroll_frame,
                text=sug,
                command=lambda s=sug: self.select(s),
                anchor="w",
                width=380,
                height=30
            )
            btn.pack(pady=2, fill="x")

        self.after(timeout * 1000, self.destroy)

    def _center_geometry(self, w: int, h: int) -> str:
        screen_width, screen_height = pyautogui.size()
        x = (screen_width - w) // 2
        y = (screen_height - h) // 2
        return f"{w}x{h}+{x}+{y}"

    def select(self, correction: str):
        self.callback(self.phrase, correction)
        self.destroy()


# --- Main Application ---
class GrammarPalApp:
    def __init__(self):
        setup_logging()
        self.config = ConfigManager()
        self.keywords = self.load_keywords()
        self.ge = GrammarEngine(self.keywords)

        # State variables
        self.typed_chars: List[str] = []
        self.phrase_buffer: Deque[str] = deque(maxlen=self.config.get("max_phrase_length", 10))
        self.recent_phrases: Deque[str] = deque(maxlen=self.config.get("recent_phrases_size", 50))
        self.suggestion_active: bool = False
        self.listener_running: bool = False
        self.last_focused_window: Optional[str] = None

        # Initialize UI
        self.root = ctk.CTk()
        self.root.title(f"{APP_NAME} v{VERSION}")
        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)

        # System tray
        self.tray_icon = self.create_tray_icon()

        # Setup UI
        self.setup_ui()

        # Start keyboard listener
        self.start_keyboard_listener()

        # Start minimized if configured
        if self.config.get("start_minimized", True):
            self.minimize_to_tray()

    def setup_ui(self):
        ctk.set_appearance_mode(self.config.get("theme", "dark"))
        self.root.geometry("600x500")

        # Main frame
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Title frame
        title_frame = ctk.CTkFrame(main_frame)
        title_frame.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(title_frame, text=f"{APP_NAME}", font=("Arial", 18, "bold")).pack(side="left", padx=10)
        ctk.CTkLabel(title_frame, text=f"v{VERSION}", font=("Arial", 12)).pack(side="right", padx=10)

        # Status frame
        status_frame = ctk.CTkFrame(main_frame)
        status_frame.pack(fill="x", pady=5)

        self.status_label = ctk.CTkLabel(status_frame, text="Status: Running", text_color="green")
        self.status_label.pack(side="left", padx=10)

        # Stats frame
        stats_frame = ctk.CTkFrame(main_frame)
        stats_frame.pack(fill="x", pady=10)

        stats_grid = ctk.CTkFrame(stats_frame)
        stats_grid.pack(pady=5)

        ctk.CTkLabel(stats_grid, text="Keywords loaded:", font=("Arial", 12)).grid(row=0, column=0, sticky="w", padx=5,
                                                                                   pady=2)
        self.keywords_label = ctk.CTkLabel(stats_grid, text=str(len(self.keywords)), font=("Arial", 12, "bold"))
        self.keywords_label.grid(row=0, column=1, sticky="w", padx=5, pady=2)

        ctk.CTkLabel(stats_grid, text="Recent phrases:", font=("Arial", 12)).grid(row=1, column=0, sticky="w", padx=5,
                                                                                  pady=2)
        self.phrases_label = ctk.CTkLabel(stats_grid, text=str(len(self.recent_phrases)), font=("Arial", 12, "bold"))
        self.phrases_label.grid(row=1, column=1, sticky="w", padx=5, pady=2)

        # Settings frame
        settings_frame = ctk.CTkFrame(main_frame)
        settings_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(settings_frame, text="Settings", font=("Arial", 14, "bold")).pack(anchor="w", padx=10, pady=5)

        # Theme setting
        theme_frame = ctk.CTkFrame(settings_frame)
        theme_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(theme_frame, text="Theme:", font=("Arial", 12)).pack(side="left", padx=5)
        self.theme_var = ctk.StringVar(value=self.config.get("theme", "dark"))
        theme_menu = ctk.CTkOptionMenu(
            theme_frame,
            values=["dark", "light", "system"],
            variable=self.theme_var,
            command=self.change_theme,
            width=100
        )
        theme_menu.pack(side="left")

        # Partial matching toggle
        self.partial_match_var = ctk.BooleanVar(value=self.config.get("enable_partial_matching", True))
        ctk.CTkCheckBox(
            settings_frame,
            text="Enable partial word matching",
            variable=self.partial_match_var,
            command=self.toggle_partial_matching,
            font=("Arial", 12)
        ).pack(anchor="w", padx=10, pady=5)

        # Start minimized toggle
        self.start_minimized_var = ctk.BooleanVar(value=self.config.get("start_minimized", True))
        ctk.CTkCheckBox(
            settings_frame,
            text="Start minimized to tray",
            variable=self.start_minimized_var,
            command=self.toggle_start_minimized,
            font=("Arial", 12)
        ).pack(anchor="w", padx=10, pady=5)

        # Buttons frame
        buttons_frame = ctk.CTkFrame(main_frame)
        buttons_frame.pack(pady=10)

        ctk.CTkButton(
            buttons_frame,
            text="Reload Keywords",
            command=self.reload_keywords,
            width=120,
            height=30
        ).pack(side="left", padx=10)

        ctk.CTkButton(
            buttons_frame,
            text="Exit",
            command=self.quit_app,
            fg_color="#d9534f",
            hover_color="#c9302c",
            width=120,
            height=30
        ).pack(side="left", padx=10)

    def create_tray_icon(self):
        # Create a simple icon image
        image = Image.new('RGB', (64, 64), color='blue')

        menu = (
            pystray.MenuItem("Show", self.show_from_tray),
            pystray.MenuItem("Exit", self.quit_app)
        )

        icon = pystray.Icon(
            APP_NAME,
            image,
            f"{APP_NAME} {VERSION}",
            menu
        )

        threading.Thread(
            target=icon.run,
            daemon=True
        ).start()

        return icon

    def show_from_tray(self):
        self.root.after(0, self.root.deiconify)

    def minimize_to_tray(self):
        self.root.withdraw()

    def change_theme(self, choice):
        ctk.set_appearance_mode(choice)
        self.config.set("theme", choice)

    def toggle_partial_matching(self):
        self.config.set("enable_partial_matching", self.partial_match_var.get())

    def toggle_start_minimized(self):
        self.config.set("start_minimized", self.start_minimized_var.get())

    def load_keywords(self) -> List[str]:
        try:
            filename = self.config.get("keywords_file", DEFAULT_KEYWORDS_FILE)
            with open(filename, 'r', encoding='utf-8') as f:
                keywords = [line.strip() for line in f if line.strip()]
                logging.info(f"Loaded {len(keywords)} keywords from {filename}")
                return keywords
        except Exception as e:
            logging.error(f"Error loading keywords: {e}")
            return []

    def reload_keywords(self):
        self.keywords = self.load_keywords()
        self.ge = GrammarEngine(self.keywords)
        self.keywords_label.configure(text=str(len(self.keywords)))
        logging.info("Keywords reloaded")
        self.status_label.configure(text="Status: Keywords Reloaded", text_color="blue")
        self.root.after(3000, lambda: self.status_label.configure(text="Status: Running", text_color="green"))

    def start_keyboard_listener(self):
        if not self.listener_running:
            keyboard.unhook_all()
            keyboard.on_press(self.on_key_press)
            self.listener_running = True
            logging.info("Keyboard listener started")

    def on_key_press(self, event):
        name = event.name

        if len(name) == 1 and name.isprintable():
            self.typed_chars.append(name)
        elif name in WORD_BOUNDARIES:
            if self.typed_chars:
                word = ''.join(self.typed_chars).strip()
                self.typed_chars.clear()
                if word:
                    self.phrase_buffer.append(word)
                    self.suggestion_active = False
                    self.check_combinations()
        elif name == 'backspace' and self.typed_chars:
            self.typed_chars.pop()

    def check_combinations(self):
        for n in range(4, 0, -1):
            if len(self.phrase_buffer) >= n:
                phrase = ' '.join(list(self.phrase_buffer)[-n:])
                norm_phrase = GrammarEngine.normalize_text(phrase)
                if norm_phrase not in self.recent_phrases:
                    self.on_phrase_completed(phrase)

    def on_phrase_completed(self, phrase: str):
        if self.suggestion_active:
            return

        min_similarity = self.config.get("min_similarity", 0.5)
        partial_matching = self.config.get("enable_partial_matching", True)
        suggestions = self.ge.check_phrase(phrase, min_similarity, partial_matching)

        if suggestions:
            norm_phrase = GrammarEngine.normalize_text(phrase)
            self.recent_phrases.append(norm_phrase)
            self.suggestion_active = True
            logging.info(f"Match found: '{phrase}' → {suggestions[:3]}... (Total: {len(suggestions)})")

            try:
                active = gw.getActiveWindow()
                if active:
                    self.last_focused_window = active.title
            except Exception as e:
                logging.warning(f"Couldn't get active window: {e}")
                self.last_focused_window = None

            timeout = self.config.get("suggestion_timeout", 8)
            threading.Thread(
                target=lambda: SuggestionPopup(
                    phrase,
                    suggestions,
                    self.apply_correction,
                    timeout
                ),
                daemon=True
            ).start()

            self.root.after(0, lambda: self.phrases_label.configure(text=str(len(self.recent_phrases))))

    def apply_correction(self, original: str, correction: str):
        try:
            logging.info(f"Applying correction: '{original}' → '{correction}'")
            time.sleep(0.2)

            if self.last_focused_window:
                try:
                    win = gw.getWindowsWithTitle(self.last_focused_window)
                    if win:
                        win[0].activate()
                        time.sleep(0.4)
                except Exception as e:
                    logging.warning(f"Couldn't activate window: {e}")

            pyperclip.copy(correction + ' ')

            # Delete the original phrase
            for _ in range(len(original)):
                pyautogui.press('backspace')
                time.sleep(0.001)
            pyautogui.press('backspace')  # Extra to ensure removal

            # Paste the correction
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.05)

            self.typed_chars.clear()
            self.phrase_buffer.clear()

        except Exception as e:
            logging.error(f"Correction failed: {e}")
        finally:
            self.suggestion_active = False

    def quit_app(self):
        logging.info("Shutting down application")
        keyboard.unhook_all()
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.quit()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


# --- Main Entry Point ---
if __name__ == "__main__":
    app = GrammarPalApp()
    app.run()
