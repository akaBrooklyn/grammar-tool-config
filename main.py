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
import winsound
import webbrowser
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
HISTORY_FILE = "correction_history.json"
WORD_BOUNDARIES = {'space', 'enter', 'tab', '.', ',', '?', '!', ';', ':', '\n', ')', '(', '[', ']', '{', '}', '/', '\\',
                   '|', '"', "'"}
MIN_SUGGESTIONS = 20
MAX_SUGGESTIONS = 30
MIN_PHRASE_LENGTH = 3
MAX_HISTORY_ITEMS = 100
DEFAULT_HOTKEY = "ctrl+alt+g"
BEEP_FREQ = 1000  # Hz
BEEP_DUR = 200  # ms
GOOGLE_SEARCH_URL = "https://www.google.com/search?q="
SEARCH_ENGINES = {
    "Google": "https://www.google.com/search?q=",
    "Bing": "https://www.bing.com/search?q=",
    "DuckDuckGo": "https://duckduckgo.com/?q=",
    "Yahoo": "https://search.yahoo.com/search?p="
}


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
            "max_phrase_length": 20,
            "min_similarity": 0.5,
            "theme": "dark",
            "start_minimized": True,
            "enable_partial_matching": True,
            "enable_auto_correct": False,
            "hotkey_force_suggest": DEFAULT_HOTKEY,
            "min_phrase_length": MIN_PHRASE_LENGTH,
            "show_definitions": True,
            "enable_sound": True,
            "enable_history": True,
            "auto_save_history": True,
            "max_history_items": MAX_HISTORY_ITEMS,
            "enable_word_learning": True,
            "show_notifications": True,
            "enable_keyboard_nav": True,
            "enable_web_search": True,
            "search_engine": "Google",
            "web_search_keyword": "search",
            "clear_buffer_after_search": True
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
        self.history = self.load_history()
        self.learned_words = set()

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

    def load_history(self) -> List[Dict]:
        try:
            if os.path.exists(HISTORY_FILE):
                with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return []
        except Exception as e:
            logging.error(f"History load error: {e}")
            return []

    def save_history(self):
        try:
            with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.history[-MAX_HISTORY_ITEMS:], f, indent=2, ensure_ascii=False)
        except Exception as e:
            logging.error(f"History save error: {e}")

    def add_to_history(self, original: str, correction: str):
        entry = {
            "timestamp": time.time(),
            "original": original,
            "correction": correction,
            "learned": False
        }
        self.history.append(entry)
        if self.config.get("auto_save_history", True):
            self.save_history()

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

    def learn_word(self, word: str):
        """Add a word to the learned words set"""
        norm_word = self.normalize_text(word)
        self.learned_words.add(norm_word)
        # Update history if this word was corrected before
        for entry in self.history:
            if self.normalize_text(entry["original"]) == norm_word:
                entry["learned"] = True


# --- Suggestion Popup ---
class SuggestionPopup(ctk.CTkToplevel):
    def __init__(self, phrase: str, suggestions: List[str], callback, ignore_callback, config: Dict, timeout: int = 8):
        super().__init__()
        self.phrase = phrase
        self.callback = callback
        self.ignore_callback = ignore_callback
        self.config = config
        self.suggestions = suggestions
        self.selected_index = -1
        self.title(f"{APP_NAME} - Suggestions")
        self.attributes("-topmost", True)
        self.geometry(self._center_geometry(450, min(600, 50 + 40 * min(len(suggestions), 10))))
        self.resizable(False, False)

        # Bind keyboard events
        if self.config.get("enable_keyboard_nav", True):
            self.bind("<Up>", self._select_previous)
            self.bind("<Down>", self._select_next)
            self.bind("<Return>", self._select_current)
            self.bind("<Escape>", lambda e: self.on_ignore())

        # Main container
        container = ctk.CTkFrame(self)
        container.pack(fill="both", expand=True, padx=10, pady=10)

        # Phrase display
        display_phrase = phrase[:50] + ('...' if len(phrase) > 50 else '')
        label = ctk.CTkLabel(container, text=f"Did you mean instead of: '{display_phrase}'")
        label.pack(pady=(0, 10))

        # Dictionary definition
        if hasattr(self.master, 'ge') and phrase.lower() in self.master.ge.dictionary and self.config.get(
                "show_definitions", True):
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

        # Web search button if enabled
        if self.config.get("enable_web_search", True):
            search_frame = ctk.CTkFrame(container)
            search_frame.pack(fill="x", pady=(0, 10))
            ctk.CTkButton(
                search_frame,
                text=f"Search on {self.config.get('search_engine', 'Google')}",
                command=self.search_web,
                fg_color="#4285F4",
                hover_color="#3367D6",
                width=200
            ).pack(pady=5)

        # Suggestions scrollable area
        self.scroll_frame = ctk.CTkScrollableFrame(container, height=min(300, 35 * len(suggestions)))
        self.scroll_frame.pack(fill="both", expand=True)

        self.suggestion_buttons = []
        for i, sug in enumerate(suggestions[:20]):
            btn = ctk.CTkButton(
                self.scroll_frame,
                text=sug,
                command=lambda s=sug: self.select(s),
                anchor="w",
                width=420,
                height=30
            )
            btn.pack(pady=2, fill="x")
            self.suggestion_buttons.append(btn)

        # Bottom buttons
        btn_frame = ctk.CTkFrame(container)
        btn_frame.pack(fill="x", pady=(10, 0))

        ctk.CTkButton(
            btn_frame,
            text="Ignore",
            command=self.on_ignore,
            fg_color="gray",
            hover_color="darkgray",
            width=100
        ).pack(side="right", padx=5)

        if self.config.get("enable_word_learning", True):
            ctk.CTkButton(
                btn_frame,
                text="Learn Word",
                command=self.learn_word,
                fg_color="#5bc0de",
                hover_color="#46b8da",
                width=100
            ).pack(side="right", padx=5)

        self.after(timeout * 1000, self.destroy)
        self.focus_set()

    def _center_geometry(self, w: int, h: int) -> str:
        screen_width, screen_height = pyautogui.size()
        x = (screen_width - w) // 2
        y = (screen_height - h) // 2
        return f"{w}x{h}+{x}+{y}"

    def _select_previous(self, event=None):
        if len(self.suggestion_buttons) == 0:
            return
        if self.selected_index > 0:
            self.selected_index -= 1
        else:
            self.selected_index = len(self.suggestion_buttons) - 1
        self._highlight_selected()

    def _select_next(self, event=None):
        if len(self.suggestion_buttons) == 0:
            return
        if self.selected_index < len(self.suggestion_buttons) - 1:
            self.selected_index += 1
        else:
            self.selected_index = 0
        self._highlight_selected()

    def _select_current(self, event=None):
        if 0 <= self.selected_index < len(self.suggestion_buttons):
            self.select(self.suggestions[self.selected_index])

    def _highlight_selected(self):
        for i, btn in enumerate(self.suggestion_buttons):
            if i == self.selected_index:
                btn.configure(fg_color="#3a7ebf", hover_color="#1f538d")
            else:
                btn.configure(fg_color=("gray85", "gray25"), hover_color=("gray75", "gray35"))

    def select(self, correction: str):
        self.callback(self.phrase, correction)
        self.destroy()

    def on_ignore(self):
        if self.ignore_callback:
            self.ignore_callback(self.phrase)
        self.destroy()

    def learn_word(self):
        if hasattr(self.master, 'ge'):
            self.master.ge.learn_word(self.phrase)
            if self.config.get("show_notifications", True):
                self.master.show_notification(f"Learned word: {self.phrase}")
        self.destroy()

    def search_web(self):
        """Open web browser with search query"""
        search_engine = SEARCH_ENGINES.get(self.config.get("search_engine", "Google"), GOOGLE_SEARCH_URL)
        webbrowser.open_new_tab(search_engine + self.phrase.replace(" ", "+"))
        if self.config.get("clear_buffer_after_search", True):
            self.on_ignore()


# --- History Window ---
class HistoryWindow(ctk.CTkToplevel):
    def __init__(self, history: List[Dict]):
        super().__init__()
        self.title(f"{APP_NAME} - Correction History")
        self.geometry("800x600")
        self.history = history

        # Main container
        container = ctk.CTkFrame(self)
        container.pack(fill="both", expand=True, padx=10, pady=10)

        # Title
        ctk.CTkLabel(container, text="Correction History", font=("Arial", 16, "bold")).pack(pady=(0, 10))

        # Search frame
        search_frame = ctk.CTkFrame(container)
        search_frame.pack(fill="x", pady=(0, 10))

        self.search_var = ctk.StringVar()
        search_entry = ctk.CTkEntry(
            search_frame,
            textvariable=self.search_var,
            placeholder_text="Search history...",
            width=300
        )
        search_entry.pack(side="left", padx=5)
        search_entry.bind("<KeyRelease>", lambda e: self.filter_history())

        ctk.CTkButton(
            search_frame,
            text="Clear History",
            command=self.clear_history,
            fg_color="#d9534f",
            hover_color="#c9302c",
            width=120
        ).pack(side="right", padx=5)

        ctk.CTkButton(
            search_frame,
            text="Refresh",
            command=self.refresh_history,
            width=80
        ).pack(side="right", padx=5)

        # History list
        self.history_frame = ctk.CTkScrollableFrame(container)
        self.history_frame.pack(fill="both", expand=True)

        self.refresh_history()

    def filter_history(self):
        search_term = self.search_var.get().lower()
        for widget in self.history_frame.winfo_children():
            widget.destroy()

        for entry in self.history:
            if (search_term in entry["original"].lower() or
                    search_term in entry["correction"].lower() or
                    not search_term):
                self._add_history_entry(entry)

    def refresh_history(self):
        for widget in self.history_frame.winfo_children():
            widget.destroy()

        for entry in self.history:
            self._add_history_entry(entry)

    def _add_history_entry(self, entry: Dict):
        entry_frame = ctk.CTkFrame(self.history_frame)
        entry_frame.pack(fill="x", pady=2)

        timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(entry["timestamp"]))
        ctk.CTkLabel(
            entry_frame,
            text=timestamp,
            width=150
        ).pack(side="left", padx=5)

        ctk.CTkLabel(
            entry_frame,
            text=f"'{entry['original']}' → '{entry['correction']}'",
            width=400,
            anchor="w"
        ).pack(side="left", padx=5)

        if entry.get("learned", False):
            ctk.CTkLabel(
                entry_frame,
                text="✓",
                text_color="green",
                width=20
            ).pack(side="right", padx=5)

    def clear_history(self):
        if hasattr(self.master, 'ge'):
            self.master.ge.history = []
            self.master.ge.save_history()
            self.refresh_history()


# --- Main Application ---
class GrammarPalApp:
    def __init__(self):
        setup_logging()
        self.config = ConfigManager()
        self.keywords = self.load_keywords()
        self.ge = GrammarEngine(self.keywords)
        self.ge.config = self.config  # Pass config to grammar engine

        # State variables
        self.typed_chars: List[str] = []
        self.phrase_buffer: Deque[str] = deque(maxlen=self.config.get("max_phrase_length", 10))
        self.suggestion_active: bool = False
        self.listener_running: bool = False
        self.last_focused_window: Optional[str] = None
        self.force_suggest_mode: bool = False
        self.web_search_mode: bool = False

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
        self.root.geometry("800x850")

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

        ctk.CTkLabel(stats_grid, text="Dictionary words:", font=("Arial", 12)).grid(row=1, column=0, sticky="w", padx=5,
                                                                                    pady=2)
        self.dict_label = ctk.CTkLabel(stats_grid, text=str(len(self.ge.dictionary)), font=("Arial", 12, "bold"))
        self.dict_label.grid(row=1, column=1, sticky="w", padx=5, pady=2)

        ctk.CTkLabel(stats_grid, text="History items:", font=("Arial", 12)).grid(row=2, column=0, sticky="w", padx=5,
                                                                                 pady=2)
        self.history_label = ctk.CTkLabel(stats_grid, text=str(len(self.ge.history)), font=("Arial", 12, "bold"))
        self.history_label.grid(row=2, column=1, sticky="w", padx=5, pady=2)

        # Dictionary frame
        dict_frame = ctk.CTkFrame(main_frame)
        dict_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(dict_frame, text="Dictionary Tool", font=("Arial", 14, "bold")).pack(anchor="w", padx=10,
                                                                                          pady=(0, 5))

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

        # History frame
        history_frame = ctk.CTkFrame(main_frame)
        history_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(history_frame, text="History Tools", font=("Arial", 14, "bold")).pack(anchor="w", padx=10, pady=5)

        history_btn_frame = ctk.CTkFrame(history_frame)
        history_btn_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkButton(
            history_btn_frame,
            text="View History",
            command=self.view_history,
            width=150
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            history_btn_frame,
            text="Export History",
            command=self.export_history,
            width=150
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            history_btn_frame,
            text="Clear History",
            command=self.clear_history,
            width=150,
            fg_color="#d9534f",
            hover_color="#c9302c"
        ).pack(side="left", padx=5)

        # Web search frame
        web_frame = ctk.CTkFrame(main_frame)
        web_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(web_frame, text="Web Search", font=("Arial", 14, "bold")).pack(anchor="w", padx=10, pady=5)

        web_search_frame = ctk.CTkFrame(web_frame)
        web_search_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(web_search_frame, text="Search Engine:", font=("Arial", 12)).pack(side="left", padx=5)
        self.search_engine_var = ctk.StringVar(value=self.config.get("search_engine", "Google"))
        search_engine_menu = ctk.CTkOptionMenu(
            web_search_frame,
            values=list(SEARCH_ENGINES.keys()),
            variable=self.search_engine_var,
            command=self.change_search_engine,
            width=120
        )
        search_engine_menu.pack(side="left", padx=5)

        web_keyword_frame = ctk.CTkFrame(web_frame)
        web_keyword_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(web_keyword_frame, text="Search Keyword:", font=("Arial", 12)).pack(side="left", padx=5)
        self.web_keyword_var = ctk.StringVar(value=self.config.get("web_search_keyword", "search"))
        web_keyword_entry = ctk.CTkEntry(
            web_keyword_frame,
            textvariable=self.web_keyword_var,
            width=120
        )
        web_keyword_entry.pack(side="left", padx=5)
        ctk.CTkButton(
            web_keyword_frame,
            text="Set",
            command=self.update_web_keyword,
            width=50
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
        self.hotkey_var = ctk.StringVar(value=self.config.get("hotkey_force_suggest", DEFAULT_HOTKEY))
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

        # Second row of checkboxes
        check_frame2 = ctk.CTkFrame(settings_frame)
        check_frame2.pack(fill="x", padx=10, pady=5)

        self.start_minimized_var = ctk.BooleanVar(value=self.config.get("start_minimized", True))
        ctk.CTkCheckBox(
            check_frame2,
            text="Start minimized",
            variable=self.start_minimized_var,
            command=self.toggle_start_minimized
        ).pack(side="left", padx=10)

        self.enable_sound_var = ctk.BooleanVar(value=self.config.get("enable_sound", True))
        ctk.CTkCheckBox(
            check_frame2,
            text="Enable sounds",
            variable=self.enable_sound_var,
            command=self.toggle_enable_sound
        ).pack(side="left", padx=10)

        self.enable_history_var = ctk.BooleanVar(value=self.config.get("enable_history", True))
        ctk.CTkCheckBox(
            check_frame2,
            text="Enable history",
            variable=self.enable_history_var,
            command=self.toggle_enable_history
        ).pack(side="left", padx=10)

        self.enable_web_var = ctk.BooleanVar(value=self.config.get("enable_web_search", True))
        ctk.CTkCheckBox(
            check_frame2,
            text="Enable web search",
            variable=self.enable_web_var,
            command=self.toggle_enable_web_search
        ).pack(side="left", padx=10)

        # Third row of checkboxes
        check_frame3 = ctk.CTkFrame(settings_frame)
        check_frame3.pack(fill="x", padx=10, pady=5)

        self.auto_save_history_var = ctk.BooleanVar(value=self.config.get("auto_save_history", True))
        ctk.CTkCheckBox(
            check_frame3,
            text="Auto-save history",
            variable=self.auto_save_history_var,
            command=self.toggle_auto_save_history
        ).pack(side="left", padx=10)

        self.word_learning_var = ctk.BooleanVar(value=self.config.get("enable_word_learning", True))
        ctk.CTkCheckBox(
            check_frame3,
            text="Word learning",
            variable=self.word_learning_var,
            command=self.toggle_word_learning
        ).pack(side="left", padx=10)

        self.show_notifications_var = ctk.BooleanVar(value=self.config.get("show_notifications", True))
        ctk.CTkCheckBox(
            check_frame3,
            text="Show notifications",
            variable=self.show_notifications_var,
            command=self.toggle_show_notifications
        ).pack(side="left", padx=10)

        self.keyboard_nav_var = ctk.BooleanVar(value=self.config.get("enable_keyboard_nav", True))
        ctk.CTkCheckBox(
            check_frame3,
            text="Keyboard navigation",
            variable=self.keyboard_nav_var,
            command=self.toggle_keyboard_nav
        ).pack(side="left", padx=10)

        self.clear_buffer_var = ctk.BooleanVar(value=self.config.get("clear_buffer_after_search", True))
        ctk.CTkCheckBox(
            check_frame3,
            text="Clear buffer after search",
            variable=self.clear_buffer_var,
            command=self.toggle_clear_buffer
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
            pystray.MenuItem("View History", self.view_history),
            pystray.MenuItem("Quick Search", self.quick_web_search),
            pystray.MenuItem("Exit", self.quit_app)
        )
        icon = pystray.Icon(APP_NAME, image, f"{APP_NAME} {VERSION}", menu)
        threading.Thread(target=icon.run, daemon=True).start()
        return icon

    def show_from_tray(self):
        self.root.after(0, self.root.deiconify())

    def minimize_to_tray(self):
        self.root.withdraw()

    def change_theme(self, choice):
        ctk.set_appearance_mode(choice)
        self.config.set("theme", choice)

    def change_search_engine(self, choice):
        self.config.set("search_engine", choice)

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

    def toggle_enable_history(self):
        self.config.set("enable_history", self.enable_history_var.get())

    def toggle_enable_web_search(self):
        self.config.set("enable_web_search", self.enable_web_var.get())

    def toggle_auto_save_history(self):
        self.config.set("auto_save_history", self.auto_save_history_var.get())

    def toggle_word_learning(self):
        self.config.set("enable_word_learning", self.word_learning_var.get())

    def toggle_show_notifications(self):
        self.config.set("show_notifications", self.show_notifications_var.get())

    def toggle_keyboard_nav(self):
        self.config.set("enable_keyboard_nav", self.keyboard_nav_var.get())

    def toggle_clear_buffer(self):
        self.config.set("clear_buffer_after_search", self.clear_buffer_var.get())

    def update_hotkey(self):
        new_hotkey = self.hotkey_var.get().strip().lower()
        if new_hotkey:
            try:
                keyboard.remove_hotkey(self.config.get("hotkey_force_suggest", DEFAULT_HOTKEY))
                keyboard.add_hotkey(new_hotkey, self.force_suggestion)
                self.config.set("hotkey_force_suggest", new_hotkey)
                self.show_notification(f"Hotkey updated to {new_hotkey}")
            except Exception as e:
                logging.error(f"Error setting hotkey: {e}")
                self.status_label.configure(text="Status: Hotkey Error", text_color="red")
                self.root.after(3000, lambda: self.status_label.configure(text="Status: Running", text_color="green"))

    def update_web_keyword(self):
        new_keyword = self.web_keyword_var.get().strip().lower()
        if new_keyword:
            self.config.set("web_search_keyword", new_keyword)
            self.show_notification(f"Web search keyword updated to '{new_keyword}'")

    def force_suggestion(self):
        if self.phrase_buffer:
            phrase = ' '.join(self.phrase_buffer)
            self.force_suggest_mode = True
            self.on_phrase_completed(phrase)

    def quick_web_search(self):
        """Trigger web search from tray menu"""
        if self.phrase_buffer:
            phrase = ' '.join(self.phrase_buffer)
            self.web_search_mode = True
            self.on_phrase_completed(phrase)

    def add_to_dictionary(self):
        word = self.dict_word.get().strip().lower()
        definition = self.dict_def.get().strip()

        if not word:
            self.show_notification("Error: Word cannot be empty", is_error=True)
            return

        if not definition:
            self.show_notification("Error: Definition cannot be empty", is_error=True)
            return

        try:
            self.ge.dictionary[word] = definition
            self.ge.save_dictionary()
            self.dict_label.configure(text=str(len(self.ge.dictionary)))
            self.show_notification(f"Added '{word}' to dictionary")
            self.dict_word.set("")
            self.dict_def.set("")
        except Exception as e:
            logging.error(f"Error adding to dictionary: {e}")
            self.show_notification("Error saving dictionary", is_error=True)

    def lookup_word(self):
        word = self.dict_word.get().strip().lower()
        if word in self.ge.dictionary:
            self.dict_def.set(self.ge.dictionary[word])
            self.show_notification(f"Found definition for '{word}'")
        else:
            self.show_notification(f"'{word}' not in dictionary", is_warning=True)

    def remove_from_dictionary(self):
        word = self.dict_word.get().strip().lower()
        if word in self.ge.dictionary:
            del self.ge.dictionary[word]
            self.ge.save_dictionary()
            self.dict_label.configure(text=str(len(self.ge.dictionary)))
            self.show_notification(f"Removed '{word}' from dictionary")
            self.dict_word.set("")
            self.dict_def.set("")
        else:
            self.show_notification(f"'{word}' not in dictionary", is_warning=True)

    def export_dictionary(self):
        try:
            export_file = "grammarpal_dictionary_export.json"
            with open(export_file, 'w', encoding='utf-8') as f:
                json.dump(self.ge.dictionary, f, indent=2, ensure_ascii=False)
            self.show_notification(f"Dictionary exported to {export_file}")
        except Exception as e:
            logging.error(f"Error exporting dictionary: {e}")
            self.show_notification("Error exporting dictionary", is_error=True)

    def view_history(self):
        HistoryWindow(self.ge.history)

    def export_history(self):
        try:
            export_file = "grammarpal_history_export.json"
            with open(export_file, 'w', encoding='utf-8') as f:
                json.dump(self.ge.history, f, indent=2, ensure_ascii=False)
            self.show_notification(f"History exported to {export_file}")
        except Exception as e:
            logging.error(f"Error exporting history: {e}")
            self.show_notification("Error exporting history", is_error=True)

    def clear_history(self):
        self.ge.history = []
        self.ge.save_history()
        self.history_label.configure(text="0")
        self.show_notification("History cleared")

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
        self.ge.config = self.config
        self.keywords_label.configure(text=str(len(self.keywords)))
        self.show_notification("Keywords reloaded")

    def start_keyboard_listener(self):
        if not self.listener_running:
            keyboard.unhook_all()
            keyboard.on_press(self.on_key_press)
            keyboard.add_hotkey(
                self.config.get("hotkey_force_suggest", DEFAULT_HOTKEY),
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
                    # Check for web search keyword
                    if (self.config.get("enable_web_search", True) and
                            word.lower() == self.config.get("web_search_keyword", "search")):
                        self.web_search_mode = True
                        return

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
        if self.suggestion_active and not self.force_suggest_mode and not self.web_search_mode:
            return

        if len(phrase) < self.config.get("min_phrase_length", MIN_PHRASE_LENGTH):
            return

        # Skip learned words
        if (not self.web_search_mode and
                self.config.get("enable_word_learning", True) and
                self.ge.normalize_text(phrase) in self.ge.learned_words):
            return

        if self.web_search_mode:
            self.handle_web_search(phrase)
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
                    self.clear_phrase_buffer,
                    self.config.config,
                    timeout
                ),
                daemon=True
            ).start()
            self.force_suggest_mode = False

    def handle_web_search(self, phrase: str):
        """Handle web search functionality"""
        search_engine = SEARCH_ENGINES.get(self.config.get("search_engine", "Google"), GOOGLE_SEARCH_URL)
        webbrowser.open_new_tab(search_engine + phrase.replace(" ", "+"))

        if self.config.get("clear_buffer_after_search", True):
            self.phrase_buffer.clear()
            self.typed_chars.clear()

        self.web_search_mode = False
        self.suggestion_active = False

    def clear_phrase_buffer(self, phrase: str):
        """Clear the phrase buffer when a suggestion is ignored"""
        words_to_remove = phrase.split()
        # Remove the words from the buffer
        for word in words_to_remove:
            try:
                self.phrase_buffer.remove(word)
            except ValueError:
                pass
        self.suggestion_active = False

    def apply_correction(self, original: str, correction: str):
        try:
            logging.info(f"Applying correction: '{original}' → '{correction}'")

            # Play sound if enabled
            if self.config.get("enable_sound", True):
                try:
                    winsound.Beep(BEEP_FREQ, BEEP_DUR)
                except Exception as e:
                    logging.warning(f"Couldn't play sound: {e}")

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

            # Add to history
            if self.config.get("enable_history", True):
                self.ge.add_to_history(original, correction)
                self.history_label.configure(text=str(len(self.ge.history)))

            # Clear buffers
            self.typed_chars.clear()
            self.phrase_buffer.clear()

        except Exception as e:
            logging.error(f"Correction failed: {e}")
            self.show_notification("Correction failed", is_error=True)
        finally:
            self.suggestion_active = False

    def show_notification(self, message: str, is_error: bool = False, is_warning: bool = False):
        """Show a status notification"""
        if not self.config.get("show_notifications", True):
            return

        if is_error:
            self.status_label.configure(text=f"Error: {message}", text_color="red")
        elif is_warning:
            self.status_label.configure(text=f"Warning: {message}", text_color="orange")
        else:
            self.status_label.configure(text=message, text_color="blue")

        self.root.after(3000, lambda: self.status_label.configure(text="Status: Running", text_color="green"))

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
