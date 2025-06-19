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
VERSION = "1.2.0"
CONFIG_FILE = "config.json"
DEFAULT_KEYWORDS_FILE = "keywords.txt"
DICTIONARY_FILE = "dictionary.json"
WORD_BOUNDARIES = {'space', 'enter', 'tab', '.', ',', '?', '!', ';', ':', '\n', ')', '(', '[', ']', '{', '}', '/', '\\', '|', '"', "'"}
MIN_SUGGESTIONS = 20
MAX_SUGGESTIONS = 30
MIN_PHRASE_LENGTH = 3

# --- Setup Logging ---
def setup_logging():
    log_dir = "logs"
    os.makedirs(log_dir, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(os.path.join(log_dir, f"{APP_NAME.lower()}.log"), encoding='utf-8'),
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
            "max_phrase_length": 25,
            "min_similarity": 0.5,
            "theme": "dark",
            "start_minimized": True,
            "enable_partial_matching": True,
            "enable_auto_correct": False,
            "hotkey_force_suggest": "ctrl+alt+g",
            "min_phrase_length": MIN_PHRASE_LENGTH,
            "show_definitions": True,
            "enable_sound": True
        }

        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
                    return {**default_config, **loaded}
        except Exception as e:
            logging.error(f"Error loading config: {e}")

        return default_config

    def save_config(self):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
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
        self.dictionary = self.load_dictionary()

    @staticmethod
    def normalize_text(text: str) -> str:
        text = text.lower()
        text = re.sub(r"[-_']", " ", text)
        text = re.sub(r"[^\w\s]", "", text)
        text = re.sub(r"\s+", " ", text)
        return text.strip()

    def build_word_index(self) -> Dict[str, List[str]]:
        index = {}
        for phrase in self.keywords:
            words = phrase.split()
            for word in words:
                if word not in index:
                    index[word] = []
                if phrase not in index[word]:
                    index[word].append(phrase)
        return index

    def load_dictionary(self) -> Dict[str, str]:
        try:
            if os.path.exists(DICTIONARY_FILE):
                with open(DICTIONARY_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except json.JSONDecodeError:
            logging.warning("Dictionary file corrupted, creating new one")
            return {}
        except Exception as e:
            logging.error(f"Dictionary load error: {e}")
            return {}

    def save_dictionary(self):
        try:
            with open(DICTIONARY_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.dictionary, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logging.error(f"Dictionary save error: {e}")

    def check_phrase(self, phrase: str, min_similarity: float = 0.5, partial_matching: bool = True) -> List[str]:
        norm_phrase = self.normalize_text(phrase)
        all_matches = []

        # 1. Exact prefix matches
        prefix_matches = [
            (k, 1.0) for k in self.keywords
            if k.startswith(norm_phrase) and k != norm_phrase
        ]
        all_matches.extend(prefix_matches)

        # 2. Partial word matches
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
                        all_matches.append((candidate, score * 0.9))

        # 3. Similarity matches
        similarity_matches = [
            (k, SequenceMatcher(None, norm_phrase, k).ratio())
            for k in self.keywords
            if k != norm_phrase
        ]
        all_matches.extend(similarity_matches)

        # Process matches
        unique_matches = {}
        for phrase, score in all_matches:
            if phrase not in unique_matches or score > unique_matches[phrase]:
                unique_matches[phrase] = score

        filtered = [
            (phrase, score)
            for phrase, score in unique_matches.items()
            if score >= min_similarity
        ]
        sorted_matches = sorted(filtered, key=lambda x: (-x[1], len(x[0])))

        results = [self.original_map[phrase] for phrase, _ in sorted_matches[:MAX_SUGGESTIONS]]
        
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
        self.title(f"{APP_NAME} - Suggestions")
        self.attributes("-topmost", True)
        self.geometry(self._center_geometry(450, min(600, 50 + 40 * min(len(suggestions), 10))))
        self.resizable(False, False)

        # Main container
        container = ctk.CTkFrame(self)
        container.pack(fill="both", expand=True, padx=10, pady=10)

        # Phrase display
        display_phrase = phrase[:50] + ('...' if len(phrase) > 50 else '')
        label = ctk.CTkLabel(container, text=f"Did you mean instead of: '{display_phrase}'")
        label.pack(pady=(0, 10))

        # Dictionary definition
        if hasattr(self.master, 'ge') and phrase.lower() in self.master.ge.dictionary and self.master.config.get("show_definitions", True):
            definition = self.master.ge.dictionary[phrase.lower()]
            def_frame = ctk.CTkFrame(container, fg_color=("gray90", "gray20"))
            def_frame.pack(fill="x", pady=(0, 10))
            ctk.CTkLabel(def_frame, text="Definition:", font=("Arial", 12, "bold")).pack(anchor="w")
            ctk.CTkLabel(
                def_frame,
                text=definition,
                wraplength=400,
                justify="left"
            ).pack(fill="x", padx=5, pady=5)

        # Suggestions scrollable area
        scroll_frame = ctk.CTkScrollableFrame(container, height=min(300, 35 * len(suggestions)))
        scroll_frame.pack(fill="both", expand=True)

        for sug in suggestions[:20]:
            btn = ctk.CTkButton(
                scroll_frame,
                text=sug,
                command=lambda s=sug: self.select(s),
                anchor="w",
                width=420,
                height=30
            )
            btn.pack(pady=2, fill="x")

        # Bottom buttons
        btn_frame = ctk.CTkFrame(container)
        btn_frame.pack(fill="x", pady=(10, 0))

        ctk.CTkButton(
            btn_frame,
            text="Ignore",
            command=self.destroy,
            fg_color="gray",
            hover_color="darkgray",
            width=100
        ).pack(side="right", padx=5)

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
        self.phrase_buffer: Deque[str] = deque(maxlen=self.config.get("max_phrase_length", 25))
        self.suggestion_active: bool = False
        self.listener_running: bool = False
        self.last_focused_window: Optional[str] = None
        self.force_suggest_mode: bool = False

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
        self.root.geometry("750x700")

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

        ctk.CTkLabel(stats_grid, text="Keywords loaded:", font=("Arial", 12)).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.keywords_label = ctk.CTkLabel(stats_grid, text=str(len(self.keywords)), font=("Arial", 12, "bold"))
        self.keywords_label.grid(row=0, column=1, sticky="w", padx=5, pady=2)

        ctk.CTkLabel(stats_grid, text="Dictionary words:", font=("Arial", 12)).grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.dict_label = ctk.CTkLabel(stats_grid, text=str(len(self.ge.dictionary)), font=("Arial", 12, "bold"))
        self.dict_label.grid(row=1, column=1, sticky="w", padx=5, pady=2)

        # Dictionary frame
        dict_frame = ctk.CTkFrame(main_frame)
        dict_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(dict_frame, text="Dictionary Tool", font=("Arial", 14, "bold")).pack(anchor="w", padx=10, pady=(0, 5))

        dict_input_frame = ctk.CTkFrame(dict_frame)
        dict_input_frame.pack(fill="x", padx=10, pady=5)

        self.dict_word = ctk.StringVar()
        self.dict_def = ctk.StringVar()

        ctk.CTkLabel(dict_input_frame, text="Word:", width=80).pack(side="left")
        ctk.CTkEntry(dict_input_frame, textvariable=self.dict_word).pack(side="left", fill="x", expand=True, padx=5)

        dict_btn_frame = ctk.CTkFrame(dict_frame)
        dict_btn_frame.pack(fill="x", padx=10, pady=(0, 5))

        ctk.CTkLabel(dict_btn_frame, text="Definition:", width=80).pack(side="left")
        ctk.CTkEntry(dict_btn_frame, textvariable=self.dict_def).pack(side="left", fill="x", expand=True, padx=5)

        dict_action_frame = ctk.CTkFrame(dict_frame)
        dict_action_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkButton(
            dict_action_frame,
            text="Add to Dictionary",
            command=self.add_to_dictionary,
            width=150
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            dict_action_frame,
            text="Lookup Word",
            command=self.lookup_word,
            width=150
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            dict_action_frame,
            text="Remove Word",
            command=self.remove_from_dictionary,
            width=150,
            fg_color="#d9534f",
            hover_color="#c9302c"
        ).pack(side="left", padx=5)

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

        # Hotkey frame
        hotkey_frame = ctk.CTkFrame(settings_frame)
        hotkey_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(hotkey_frame, text="Force suggestion hotkey:", font=("Arial", 12)).pack(side="left", padx=5)
        self.hotkey_var = ctk.StringVar(value=self.config.get("hotkey_force_suggest", "ctrl+alt+g"))
        hotkey_entry = ctk.CTkEntry(
            hotkey_frame,
            textvariable=self.hotkey_var,
            width=120
        )
        hotkey_entry.pack(side="left", padx=5)
        ctk.CTkButton(
            hotkey_frame,
            text="Set",
            command=self.update_hotkey,
            width=50
        ).pack(side="left", padx=5)

        # Checkboxes frame
        check_frame = ctk.CTkFrame(settings_frame)
        check_frame.pack(fill="x", padx=10, pady=5)

        self.partial_match_var = ctk.BooleanVar(value=self.config.get("enable_partial_matching", True))
        ctk.CTkCheckBox(
            check_frame,
            text="Partial word matching",
            variable=self.partial_match_var,
            command=self.toggle_partial_matching
        ).pack(side="left", padx=10)

        self.auto_correct_var = ctk.BooleanVar(value=self.config.get("enable_auto_correct", False))
        ctk.CTkCheckBox(
            check_frame,
            text="Auto-correct",
            variable=self.auto_correct_var,
            command=self.toggle_auto_correct
        ).pack(side="left", padx=10)

        self.show_def_var = ctk.BooleanVar(value=self.config.get("show_definitions", True))
        ctk.CTkCheckBox(
            check_frame,
            text="Show definitions",
            variable=self.show_def_var,
            command=self.toggle_show_definitions
        ).pack(side="left", padx=10)

        self.start_minimized_var = ctk.BooleanVar(value=self.config.get("start_minimized", True))
        ctk.CTkCheckBox(
            check_frame,
            text="Start minimized",
            variable=self.start_minimized_var,
            command=self.toggle_start_minimized
        ).pack(side="left", padx=10)

        self.enable_sound_var = ctk.BooleanVar(value=self.config.get("enable_sound", True))
        ctk.CTkCheckBox(
            check_frame,
            text="Enable sounds",
            variable=self.enable_sound_var,
            command=self.toggle_enable_sound
        ).pack(side="left", padx=10)

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
            text="Export Dictionary",
            command=self.export_dictionary,
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
        image = Image.new('RGB', (64, 64), color='blue')
        menu = (
            pystray.MenuItem("Show", self.show_from_tray),
            pystray.MenuItem("Force Suggestion", self.force_suggestion),
            pystray.MenuItem("Exit", self.quit_app)
        )
        icon = pystray.Icon(APP_NAME, image, f"{APP_NAME} {VERSION}", menu)
        threading.Thread(target=icon.run, daemon=True).start()
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

    def toggle_auto_correct(self):
        self.config.set("enable_auto_correct", self.auto_correct_var.get())

    def toggle_show_definitions(self):
        self.config.set("show_definitions", self.show_def_var.get())

    def toggle_start_minimized(self):
        self.config.set("start_minimized", self.start_minimized_var.get())

    def toggle_enable_sound(self):
        self.config.set("enable_sound", self.enable_sound_var.get())

    def update_hotkey(self):
        new_hotkey = self.hotkey_var.get().strip().lower()
        if new_hotkey:
            try:
                keyboard.remove_hotkey(self.config.get("hotkey_force_suggest", "ctrl+alt+g"))
                keyboard.add_hotkey(new_hotkey, self.force_suggestion)
                self.config.set("hotkey_force_suggest", new_hotkey)
                self.status_label.configure(text="Status: Hotkey Updated", text_color="blue")
                self.root.after(3000, lambda: self.status_label.configure(text="Status: Running", text_color="green"))
            except Exception as e:
                logging.error(f"Error setting hotkey: {e}")
                self.status_label.configure(text="Status: Hotkey Error", text_color="red")
                self.root.after(3000, lambda: self.status_label.configure(text="Status: Running", text_color="green"))

    def force_suggestion(self):
        if self.phrase_buffer:
            phrase = ' '.join(self.phrase_buffer)
            self.force_suggest_mode = True
            self.on_phrase_completed(phrase)

    def add_to_dictionary(self):
        word = self.dict_word.get().strip().lower()
        definition = self.dict_def.get().strip()
        
        if not word:
            self.status_label.configure(text="Error: Word cannot be empty", text_color="red")
            self.root.after(3000, lambda: self.status_label.configure(text="Status: Running", text_color="green"))
            return
            
        if not definition:
            self.status_label.configure(text="Error: Definition cannot be empty", text_color="red")
            self.root.after(3000, lambda: self.status_label.configure(text="Status: Running", text_color="green"))
            return

        try:
            self.ge.dictionary[word] = definition
            self.ge.save_dictionary()
            self.dict_label.configure(text=str(len(self.ge.dictionary)))
            self.status_label.configure(text=f"Added '{word}' to dictionary", text_color="blue")
            self.root.after(3000, lambda: self.status_label.configure(text="Status: Running", text_color="green"))
            self.dict_word.set("")
            self.dict_def.set("")
        except Exception as e:
            logging.error(f"Error adding to dictionary: {e}")
            self.status_label.configure(text="Error saving dictionary", text_color="red")
            self.root.after(3000, lambda: self.status_label.configure(text="Status: Running", text_color="green"))

    def lookup_word(self):
        word = self.dict_word.get().strip().lower()
        if word in self.ge.dictionary:
            self.dict_def.set(self.ge.dictionary[word])
            self.status_label.configure(text=f"Found definition for '{word}'", text_color="blue")
            self.root.after(3000, lambda: self.status_label.configure(text="Status: Running", text_color="green"))
        else:
            self.status_label.configure(text=f"'{word}' not in dictionary", text_color="orange")
            self.root.after(3000, lambda: self.status_label.configure(text="Status: Running", text_color="green"))

    def remove_from_dictionary(self):
        word = self.dict_word.get().strip().lower()
        if word in self.ge.dictionary:
            del self.ge.dictionary[word]
            self.ge.save_dictionary()
            self.dict_label.configure(text=str(len(self.ge.dictionary)))
            self.status_label.configure(text=f"Removed '{word}' from dictionary", text_color="blue")
            self.root.after(3000, lambda: self.status_label.configure(text="Status: Running", text_color="green"))
            self.dict_word.set("")
            self.dict_def.set("")
        else:
            self.status_label.configure(text=f"'{word}' not in dictionary", text_color="orange")
            self.root.after(3000, lambda: self.status_label.configure(text="Status: Running", text_color="green"))

    def export_dictionary(self):
        try:
            export_file = "grammarpal_dictionary_export.json"
            with open(export_file, 'w', encoding='utf-8') as f:
                json.dump(self.ge.dictionary, f, indent=2, ensure_ascii=False)
            self.status_label.configure(text=f"Dictionary exported to {export_file}", text_color="blue")
            self.root.after(3000, lambda: self.status_label.configure(text="Status: Running", text_color="green"))
        except Exception as e:
            logging.error(f"Error exporting dictionary: {e}")
            self.status_label.configure(text="Error exporting dictionary", text_color="red")
            self.root.after(3000, lambda: self.status_label.configure(text="Status: Running", text_color="green"))

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
            keyboard.add_hotkey(
                self.config.get("hotkey_force_suggest", "ctrl+alt+g"),
                self.force_suggestion
            )
            self.listener_running = True
            logging.info("Keyboard listener started")

    def on_key_press(self, event):
        name = event.name

        if len(name) == 1 and name.isprintable():
            if name.isdigit():
                self.typed_chars.clear()
                return
            self.typed_chars.append(name)
        elif name in WORD_BOUNDARIES:
            if self.typed_chars:
                word = ''.join(self.typed_chars).strip()
                self.typed_chars.clear()
                if word and not any(c.isdigit() for c in word):
                    self.phrase_buffer.append(word)
                    self.suggestion_active = False
                    self.check_combinations()
                    if len(self.phrase_buffer) == self.phrase_buffer.maxlen:
                        self.phrase_buffer.popleft()
        elif name == 'backspace' and self.typed_chars:
            self.typed_chars.pop()

    def check_combinations(self):
        for n in range(4, 0, -1):
            if len(self.phrase_buffer) >= n:
                phrase = ' '.join(list(self.phrase_buffer)[-n:])
                if len(phrase) >= self.config.get("min_phrase_length", MIN_PHRASE_LENGTH):
                    self.on_phrase_completed(phrase)

    def on_phrase_completed(self, phrase: str):
        if self.suggestion_active and not self.force_suggest_mode:
            return

        if len(phrase) < self.config.get("min_phrase_length", MIN_PHRASE_LENGTH):
            return

        min_similarity = self.config.get("min_similarity", 0.5)
        partial_matching = self.config.get("enable_partial_matching", True)
        suggestions = self.ge.check_phrase(phrase, min_similarity, partial_matching)

        if suggestions:
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
            self.force_suggest_mode = False

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

if __name__ == "__main__":
    app = GrammarPalApp()
    app.run()
