"""
Component Comparator AI
-----------------------
A Tkinter application for comparing electronic component specification sheets (PDFs)
using Generative AI. Users can load two spec sheets, and the application will
extract text and images to provide a comparative analysis using a selected AI model.
The conversation history can be downloaded as a Word document.
"""
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
import fitz  # PyMuPDF, for PDF processing
import re # For table formatting
# Try to import specific fitz error, fall back if not found (e.g. older PyMuPDF)
try:
    from fitz.errors import FitzError
except ImportError:
    FitzError = Exception # Fallback to generic Exception if specific error not found
import os
import shutil
import google.generativeai as genai
from google.api_core import exceptions as google_exceptions
# Import specific GenAI exceptions if needed, e.g., genai.types.BlockedPromptException
try:
    from google.generativeai.types import BlockedPromptException, StopCandidateException
except ImportError:
    BlockedPromptException = Exception # Fallback
    StopCandidateException = Exception # Fallback

from PIL import Image, ImageTk, UnidentifiedImageError # Pillow for image handling
import docx # For downloading history
from dotenv import load_dotenv # For loading .env files

class ComponentComparatorAI:
    """
    Main application class for the Component Comparator AI.
    Manages the UI, file loading, PDF processing, AI interaction,
    and conversation history.
    """
    def __init__(self, root):
        """
        Initializes the application UI and internal state.

        Args:
            root (tk.Tk): The root Tkinter window.
        """
        self.root = root
        self.root.title("Component Comparator AI")
        self.root.geometry("850x700")

        if load_dotenv():
            print("DEBUG: Loaded environment variables from .env file.")
        else:
            print("DEBUG: No .env file found or python-dotenv not available/failed to load.")

        self.spec_sheet_1_path = None
        self.spec_sheet_1_text = None
        self.spec_sheet_1_image_paths = []
        self.spec_sheet_2_path = None
        self.spec_sheet_2_text = None
        self.spec_sheet_2_image_paths = []

        self.model = None
        self.chat_session = None
        self.conversation_log = []
        self.ai_history = []
        self.api_key_configured = False
        self.model_options_list = []
        self.placeholder_text = "Select AI Model (after loading files)"
        self.model_initializing = False

        self.temp_image_dir = "temp_images"
        self._create_temp_image_dir()

        self._setup_ui(root)
        self._configure_ai()

        self.update_conversation_history(
            "System: Welcome! Please load two PDF specification sheets to compare. " +
            "After loading both files, you will be able to select an AI model.",
            role="system"
        )

        if not self.api_key_configured:
             self._update_ui_for_ai_status(api_key_configured=False, model_initialized=False)
        else:
            self._update_ui_for_ai_status(api_key_configured=True, model_initialized=False)


    def _create_temp_image_dir(self):
        """Creates the temporary image directory if it doesn't exist."""
        if not os.path.exists(self.temp_image_dir):
            try:
                os.makedirs(self.temp_image_dir)
                print(f"DEBUG: Created temporary image directory: {self.temp_image_dir}")
            except OSError as e:
                print(f"Critical Error: Cannot create temporary directory {self.temp_image_dir}: {e}")
                self.update_conversation_history(f"System: Error - Cannot create temp image folder: {e}", role="error")


    def _setup_ui(self, root):
        """Configures the Tkinter UI elements and layout."""
        self.spec_sheet_1_label = ttk.Label(root, text="File 1: None")
        self.spec_sheet_1_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.load_spec_sheet_1_button = ttk.Button(root, text="Load Spec Sheet 1", command=self.load_spec_sheet_1)
        self.load_spec_sheet_1_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.spec_sheet_2_label = ttk.Label(root, text="File 2: None")
        self.spec_sheet_2_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.load_spec_sheet_2_button = ttk.Button(root, text="Load Spec Sheet 2", command=self.load_spec_sheet_2)
        self.load_spec_sheet_2_button.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        self.model_label = ttk.Label(root, text="Select AI Model:")
        self.model_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.model_var = tk.StringVar()
        self.model_combobox = ttk.Combobox(root, textvariable=self.model_var)

        self.model_options_list = [
            "models/gemini-1.0-pro-vision-latest", "models/gemini-pro-vision",
            "models/gemini-1.5-flash-latest", "models/gemini-1.5-flash",
            "models/gemini-1.5-flash-002", "models/gemini-1.5-flash-8b",
            "models/gemini-1.5-flash-8b-001", "models/gemini-1.5-flash-8b-latest",
            "models/gemini-2.5-flash-preview-04-17", "models/gemini-2.5-flash-preview-05-20",
            "models/gemini-2.5-flash-preview-04-17-thinking", "models/gemini-2.0-flash-exp",
            "models/gemini-2.0-flash", "models/gemini-2.0-flash-001",
            "models/gemini-2.0-flash-exp-image-generation", "models/gemini-2.0-flash-lite-001",
            "models/gemini-2.0-flash-lite", "models/gemini-2.0-flash-lite-preview-02-05",
            "models/gemini-2.0-flash-lite-preview", "models/gemini-2.0-flash-thinking-exp-01-21",
            "models/gemini-2.0-flash-thinking-exp", "models/gemini-2.0-flash-thinking-exp-1219",
            "models/learnlm-2.0-flash-experimental", "models/gemma-3-1b-it",
            "models/gemma-3-4b-it", "models/gemma-3-12b-it",
            "models/gemma-3-27b-it", "models/gemma-3n-e4b-it"
        ]
        self.model_combobox['values'] = self.model_options_list

        self.model_combobox.set(self.placeholder_text)
        self.model_combobox.state(["disabled"])

        self.model_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.model_combobox.bind("<<ComboboxSelected>>", self._on_model_selected)

        self.history_label = ttk.Label(root, text="Conversation History:")
        self.history_label.grid(row=3, column=0, columnspan=2, padx=10, pady=(10,0), sticky="w")
        self.conversation_history = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=15, width=80)
        self.conversation_history.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

        self.conversation_history.tag_configure("user_message", foreground="blue", font=('Arial', 10))
        self.conversation_history.tag_configure("ai_message", foreground="#008800", font=('Arial', 10)) # Darker green
        self.conversation_history.tag_configure("system_message", foreground="#550055", font=('Arial', 10, 'italic')) # Darker purple
        self.conversation_history.tag_configure("error_message", foreground="red", font=('Arial', 10, 'bold'))

        self.conversation_history.config(state=tk.DISABLED)


        self.input_label = ttk.Label(root, text="Your Message:")
        self.input_label.grid(row=5, column=0, columnspan=2, padx=10, pady=(10,0), sticky="w")
        self.user_input_entry = ttk.Entry(root, width=70)
        self.user_input_entry.grid(row=6, column=0, padx=10, pady=5, sticky="ew")
        self.send_button = ttk.Button(root, text="Send", command=self.send_user_query)
        self.send_button.grid(row=6, column=1, padx=(0,10), pady=5, sticky="ew")

        self.download_history_button = ttk.Button(root, text="Download History", command=self.download_history)
        self.download_history_button.grid(row=7, column=0, padx=10, pady=10, sticky="ew")
        self.clear_all_button = ttk.Button(root, text="Clear All", command=self.clear_all)
        self.clear_all_button.grid(row=7, column=1, padx=(0,10), pady=10, sticky="ew")

        root.grid_columnconfigure(0, weight=3); root.grid_columnconfigure(1, weight=1)
        root.grid_rowconfigure(4, weight=1)

    def load_spec_sheet_1(self):
        filepath = filedialog.askopenfilename(title="Select Spec Sheet 1 (PDF)", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if filepath:
            self.spec_sheet_1_path = filepath
            self.spec_sheet_1_label.config(text=f"File 1: {os.path.basename(filepath)}")
            if self.spec_sheet_1_path and self.spec_sheet_2_path:
                self.model_combobox.config(state='readonly')
                self.update_conversation_history("System: Both spec sheets loaded. Please select an AI model.", role="system")
        else:
            self.spec_sheet_1_label.config(text=f"File 1: {os.path.basename(self.spec_sheet_1_path) if self.spec_sheet_1_path else 'None'}")

    def load_spec_sheet_2(self):
        filepath = filedialog.askopenfilename(title="Select Spec Sheet 2 (PDF)", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if filepath:
            self.spec_sheet_2_path = filepath
            self.spec_sheet_2_label.config(text=f"File 2: {os.path.basename(filepath)}")
            if self.spec_sheet_1_path and self.spec_sheet_2_path:
                self.model_combobox.config(state='readonly')
                self.update_conversation_history("System: Both spec sheets loaded. Please select an AI model.", role="system")
        else:
            self.spec_sheet_2_label.config(text=f"File 2: {os.path.basename(self.spec_sheet_2_path) if self.spec_sheet_2_path else 'None'}")

    def _format_ai_response(self, text_response: str) -> str:
        """Applies basic formatting to AI responses, especially for tables."""
        lines = text_response.split('\n')
        formatted_lines = []
        # Regex to detect typical markdown table separator lines (e.g., |---|---| or |:--|--:|)
        separator_pattern = re.compile(r"^\s*\|?[-:|\s]+\|?\s*$") # Made pipe at end optional

        for line in lines:
            if '|' in line:
                # For lines with table data or separators
                processed_line = line
                if separator_pattern.match(line):
                    processed_line = processed_line.replace("-", "—") # Em-dash

                # Add spacing around pipes, try to handle existing spaces gracefully
                processed_line = re.sub(r'\s*\|\s*', '  |  ', processed_line).strip()
                formatted_lines.append(processed_line)
            else:
                formatted_lines.append(line)
        return "\n".join(formatted_lines)

    def update_conversation_history(self, message, role="system"):
        # This method is for updating the UI display and the self.conversation_log (for download).
        # The raw message for AI history (self.ai_history) should be added by the calling method
        # (send_to_ai, send_user_query) using _add_to_ai_history.

        if hasattr(self, 'conversation_history') and self.conversation_history:
            self.conversation_history.config(state=tk.NORMAL)

            tag_to_apply = "system_message"
            display_message = message # Default to original message

            if role == "user":
                tag_to_apply = "user_message"
            elif role == "ai":
                tag_to_apply = "ai_message"
                display_message = self._format_ai_response(message) # Format AI messages for display
            elif role == "error":
                tag_to_apply = "error_message"

            self.conversation_history.insert(tk.END, display_message + "\n", tag_to_apply)
            self.conversation_history.see(tk.END)
            self.conversation_history.config(state=tk.DISABLED)

        # For self.conversation_log (used for download), store the message as it's passed
        # (which includes prefixes like "User: ", "AI: ", and for AI, it's the raw unformatted one).
        # The formatting by _format_ai_response is only for the ScrolledText widget.
        # This means downloaded AI responses will be raw, not table-formatted. This is a choice.
        # If formatted download is needed, `message` here for AI role should be the formatted one.
        # For now, `conversation_log` gets the potentially raw `message`.
        self.conversation_log.append(message)


    def _update_ui_for_ai_status(self, api_key_configured=None, model_initialized=None):
        if not hasattr(self, 'send_button'): return
        is_api_key_ready = api_key_configured if api_key_configured is not None else self.api_key_configured
        is_model_ready = model_initialized if model_initialized is not None else (self.model is not None)
        can_perform_ai_ops = is_api_key_ready and is_model_ready

        self.send_button.config(state=tk.NORMAL if can_perform_ai_ops else tk.DISABLED)
        self.user_input_entry.config(state=tk.NORMAL if can_perform_ai_ops else tk.DISABLED)

        if hasattr(self, 'model_combobox') and self.model_combobox.cget('state') != 'disabled':
            self.model_combobox.config(state="readonly")


    def _configure_ai(self):
        try:
            api_key = os.environ.get("GOOGLE_API_KEY")
            if not api_key:
                print("DEBUG: GOOGLE_API_KEY not found in environment for _configure_ai.")
                self.api_key_configured = False
            else:
                genai.configure(api_key=api_key)
                self.api_key_configured = True
        except Exception as e:
            self.update_conversation_history(f"System: Error configuring Generative AI SDK - {e}", role="error")
            self.api_key_configured = False
        finally:
            self._update_ui_for_ai_status(api_key_configured=self.api_key_configured, model_initialized=(self.model is not None))


    def _on_model_selected(self, event=None):
        if event: print(f"DEBUG: _on_model_selected triggered. Event type: {event.type}, Widget: {event.widget}")
        else: print(f"DEBUG: _on_model_selected triggered programmatically or without event object.")

        selected_model_name = self.model_combobox.get()
        print(f"DEBUG: Current Combobox value via get(): '{selected_model_name}'")

        if selected_model_name == self.placeholder_text:
            self.update_conversation_history("System: Please select a valid AI model to proceed.", role="system")
            self._update_ui_for_ai_status(model_initialized=False)
            return

        previous_model_name = self.model.model_name if self.model else None

        is_different_model = self.model and self.model.model_name != selected_model_name
        is_first_meaningful_selection = not self.model and self.conversation_log and \
                                       any(not log.startswith("System: Welcome!") for log in self.conversation_log)

        if is_different_model or is_first_meaningful_selection:
            log_msg = f"System: Changing model"
            if previous_model_name: log_msg += f" from {previous_model_name}"
            log_msg += f" to {selected_model_name}. Clearing previous AI context and conversation."
            self.update_conversation_history(log_msg, role="system")
            self.clear_all(clear_files=False)
            self.update_conversation_history(f"System: AI Model selected: {selected_model_name}", role="system") # Re-log after clear
        else:
             self.update_conversation_history(f"System: AI Model selected: {selected_model_name}", role="system")

        self.model_initializing = True
        model_successfully_initialized = self._initialize_model(selected_model_name)
        self.model_initializing = False

        if model_successfully_initialized:
            if self.spec_sheet_1_path and self.spec_sheet_2_path:
                self.update_conversation_history(f"System: Starting analysis with {selected_model_name}...", role="system")
                self.check_and_process_spec_sheets()
            else:
                 self.update_conversation_history("System: Model initialized. Please load both spec sheets if you haven't already.", role="system")
        else:
            self.update_conversation_history(f"System: Failed to initialize {selected_model_name}. See logs for details.", role="error")


    def _initialize_model(self, model_name=None):
        source_log = "explicitly passed"
        if model_name is None:
            model_name = self.model_combobox.get()
            source_log = f"fetched from Combobox: '{model_name}'"

        if model_name == self.placeholder_text:
            self.update_conversation_history("System: Attempted to initialize with placeholder. Please select a valid model.", role="system")
            self.model = None; self.chat_session = None
            self._update_ui_for_ai_status(model_initialized=False)
            return False

        self.update_conversation_history(f"System: _initialize_model called ({source_log}). Target model: '{model_name}'", role="system")

        if not model_name or model_name not in self.model_options_list:
            log_msg = f"System: Invalid or empty model name ('{model_name}') for initialization."
            if self.model_options_list: log_msg += f" Available: {len(self.model_options_list)}."
            else: log_msg += " Model options list N/A."
            self.update_conversation_history(log_msg + " Aborted.", role="error"); self.model = None; self.chat_session = None
            self._update_ui_for_ai_status(model_initialized=False); return False

        if self.model and self.model.model_name == model_name and self.api_key_configured:
            self.update_conversation_history(f"System: Model '{model_name}' is already active.", role="system")
            self._update_ui_for_ai_status(model_initialized=True); return True

        if not self.api_key_configured:
            self.update_conversation_history("System: Cannot initialize model - API key not configured.", role="error")
            self.model = None; self.chat_session = None
            self._update_ui_for_ai_status(model_initialized=False); return False

        self.update_conversation_history(f"System: Attempting to initialize AI model: {model_name}...", role="system")
        try:
            if not any("Generative AI configured successfully." in log for log in self.conversation_log):
                 self.update_conversation_history("System: Generative AI configured successfully.", role="system")

            self.model = genai.GenerativeModel(model_name); self.chat_session = None
            self.update_conversation_history(f"System: Successfully initialized AI model: {model_name}", role="system")
            self._update_ui_for_ai_status(model_initialized=True); return True
        except Exception as e:
            self.model = None; self.chat_session = None
            self.update_conversation_history(f"System: Error initializing AI model {model_name}: {e}", role="error")
            self._update_ui_for_ai_status(model_initialized=False); return False

    def get_selected_model_name(self): return self.model_var.get()

    def _add_to_ai_history(self, role: str, text_content: str):
        self.ai_history.append({'role': role, 'parts': [text_content]})
        print(f"DEBUG: Added to AI history: Role={role}, Content='{text_content[:50]}...'")

    def _convert_log_to_gemini_history(self):
        return [entry for entry in self.ai_history if entry['role'] in ['user', 'model']]

    def send_user_query(self):
        if not self.model or not self.api_key_configured:
            print("DEBUG: send_user_query: model or API key not configured. Attempting _initialize_model.")
            if not self._initialize_model(): return
        if not self.model:
            self.update_conversation_history("System: AI Model is not available. Cannot send message.", role="error"); return

        user_text_for_display = self.user_input_entry.get().strip()
        if not user_text_for_display: return

        self.update_conversation_history(f"User: {user_text_for_display}", role="user")
        self.user_input_entry.delete(0, tk.END)

        user_text_for_ai = user_text_for_display

        active_model_name = self.model.model_name if self.model else "Unknown Model"
        try:
            self.send_button.config(state=tk.DISABLED); self.user_input_entry.config(state=tk.DISABLED)

            if not self.chat_session:
                self.update_conversation_history(f"System: Starting new chat session with {active_model_name}...", role="system")
                try:
                    history_for_init = self._convert_log_to_gemini_history()
                    print(f"DEBUG: Starting chat with history: {history_for_init}")
                    self.chat_session = self.model.start_chat(history=history_for_init)
                except Exception as e:
                    self.update_conversation_history(f"System: Error starting chat session with {active_model_name}: {e}", role="error")
                    self._update_ui_for_ai_status()
                    return

            self._add_to_ai_history('user', user_text_for_ai)

            self.update_conversation_history(f"System: Sending to AI ({active_model_name})...", role="system");
            response = self.chat_session.send_message(user_text_for_ai)

            self._add_to_ai_history('model', response.text)
            # Pass raw response.text to update_conversation_history, it will be formatted if role is 'ai'
            self.update_conversation_history(f"AI ({active_model_name}): {response.text}", role="ai")

        except Exception as e:
            error_message = f"System: Error during AI interaction with {active_model_name}: {e}"
            self.update_conversation_history(error_message, role="error")
            print(f"DEBUG: AI Error in send_user_query: {error_message}")
            if isinstance(e, (google_exceptions.PermissionDenied, google_exceptions.Unauthenticated)):
                self.api_key_configured = False
            if isinstance(e, (google_exceptions.InvalidArgument, ValueError, BlockedPromptException, StopCandidateException, google_exceptions.NotFound, google_exceptions.PermissionDenied)):
                self.update_conversation_history(f"System: Resetting current AI model ({active_model_name}) due to error.", role="system")
                self.model = None; self.chat_session = None; self.ai_history = []
        finally: self._update_ui_for_ai_status()


    def download_history(self):
        if not self.conversation_log: self.update_conversation_history("System: History empty.", role="system"); return
        try: filepath = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")], title="Save History")
        except Exception as e: self.update_conversation_history(f"System: Error opening save dialog: {e}", role="error"); return
        if not filepath: self.update_conversation_history("System: Download cancelled.", role="system"); return
        try:
            doc = docx.Document(); doc.add_heading("Component Comparator AI Chat History", level=1)
            if self.spec_sheet_1_path: doc.add_paragraph(f"Spec Sheet 1: {os.path.basename(self.spec_sheet_1_path)}")
            if self.spec_sheet_2_path: doc.add_paragraph(f"Spec Sheet 2: {os.path.basename(self.spec_sheet_2_path)}")
            model_name = self.model.model_name if self.model and hasattr(self.model, 'model_name') else "N/A"
            doc.add_paragraph(f"AI Model (last used): {model_name}"); doc.add_paragraph("-" * 20)
            for entry in self.conversation_log: doc.add_paragraph(entry)
            doc.save(filepath); self.update_conversation_history(f"System: History downloaded to {filepath}", role="system")
        except Exception as e: self.update_conversation_history(f"System: Error downloading: {e}", role="error"); print(f"Error saving .docx: {e}")

    def clear_all(self, clear_files=True):
        print(f"DEBUG: clear_all called with clear_files={clear_files}")
        if clear_files:
            self.spec_sheet_1_path = None; self.spec_sheet_1_text = None; self.spec_sheet_1_image_paths = []
            self.spec_sheet_2_path = None; self.spec_sheet_2_text = None; self.spec_sheet_2_image_paths = []
            if hasattr(self, 'spec_sheet_1_label'): self.spec_sheet_1_label.config(text="File 1: None")
            if hasattr(self, 'spec_sheet_2_label'): self.spec_sheet_2_label.config(text="File 2: None")

            if hasattr(self, 'model_combobox'):
                self.model_combobox.set(self.placeholder_text)
                self.model_combobox.state(["disabled"])

            if os.path.exists(self.temp_image_dir):
                try: shutil.rmtree(self.temp_image_dir); print(f"DEBUG: Deleted temp image directory: {self.temp_image_dir}")
                except OSError as e: print(f"Error deleting temp image directory {self.temp_image_dir}: {e}")
            self._create_temp_image_dir()

        if hasattr(self, 'conversation_history'):
            self.conversation_history.config(state=tk.NORMAL); self.conversation_history.delete(1.0, tk.END)

        self.conversation_log = []
        self.ai_history = []

        if clear_files:
             self.update_conversation_history(
                "System: Welcome! Please load two PDF specification sheets to compare. " +
                "After loading both files, you will be able to select an AI model.", role="system"
             )

        if hasattr(self, 'user_input_entry'): self.user_input_entry.delete(0, tk.END)

        self.model = None
        self.chat_session = None

        self._configure_ai()

        if not clear_files and self.spec_sheet_1_path and self.spec_sheet_2_path:
            if hasattr(self, 'model_combobox'): self.model_combobox.config(state='readonly')
            self.update_conversation_history("System: AI context cleared. Loaded files remain. Select a model to re-analyze or start a new chat.", role="system")
        elif clear_files:
            if hasattr(self, 'model_combobox'): self.model_combobox.state(['disabled'])

        self._update_ui_for_ai_status(api_key_configured=self.api_key_configured, model_initialized=False)

        print("DEBUG: Clear All logic finished.")


    def extract_text_from_pdf(self, filepath):
        if not filepath or not os.path.exists(filepath): self.update_conversation_history(f"System: PDF not found: {os.path.basename(filepath or 'Unknown')}", role="error"); return ""
        try:
            self.update_conversation_history(f"System: Extracting text from {os.path.basename(filepath)}...", role="system")
            with fitz.open(filepath) as doc: text = "".join(page.get_text() for page in doc)
            self.update_conversation_history(f"System: Text extraction successful: {os.path.basename(filepath)}.", role="system"); return text
        except Exception as e: self.update_conversation_history(f"System: Error extracting text from {os.path.basename(filepath)}: {e}", role="error"); return ""

    def extract_images_from_pdf(self, filepath, output_folder):
        if not filepath or not os.path.exists(filepath): self.update_conversation_history(f"System: PDF not found: {os.path.basename(filepath or 'Unknown')}", role="error"); return []
        paths = []
        try:
            self.update_conversation_history(f"System: Extracting images from {os.path.basename(filepath)}...", role="system")
            if not os.path.exists(output_folder): os.makedirs(output_folder)
            with fitz.open(filepath) as doc:
                for i, page in enumerate(doc):
                    for j, img_info in enumerate(page.get_images(full=True)):
                        xref = img_info[0]
                        try: base = doc.extract_image(xref)
                        except Exception as e: self.update_conversation_history(f"System: Error extracting img xref {xref} pg {i+1}. Skip. Err: {e}", role="error"); continue
                        img_bytes, ext = base["image"], base["ext"]
                        path = os.path.join(output_folder, f"pg{i+1}_img{j+1}.{ext}")
                        try:
                            with open(path, "wb") as f: f.write(img_bytes)
                            paths.append(path)
                        except IOError as e: self.update_conversation_history(f"System: IOError saving image {path}. Error: {e}", role="error")
            msg = f"System: Extracted {len(paths)} images from {os.path.basename(filepath)}." if paths else f"System: No images found in {os.path.basename(filepath)}."
            self.update_conversation_history(msg, role="system"); return paths
        except Exception as e: self.update_conversation_history(f"System: Error extracting images from {os.path.basename(filepath)}: {e}", role="error"); return []

    def check_and_process_spec_sheets(self):
        if not (self.spec_sheet_1_path and self.spec_sheet_2_path): return

        if not self.api_key_configured:
            self.update_conversation_history("System: API Key not configured. Cannot process specs.", role="error"); return

        if not self.model:
            self.update_conversation_history("System: AI Model not selected/initialized. Please select a model to start analysis.", role="system")
            return

        self.update_conversation_history("System: Both spec sheets loaded and model active. Clearing previous analysis results...", role="system")
        self.conversation_log = []
        self.ai_history = []
        if hasattr(self, 'conversation_history'):
            self.conversation_history.config(state=tk.NORMAL); self.conversation_history.delete(1.0, tk.END)

        if self.api_key_configured: self.update_conversation_history("System: Generative AI is configured.", role="system")
        if self.model: self.update_conversation_history(f"System: Model '{self.model.model_name}' is active.", role="system")

        self.process_spec_sheets()


    def process_spec_sheets(self):
        if not self.model or not self.api_key_configured or not self.spec_sheet_1_path or not self.spec_sheet_2_path:
            self.update_conversation_history("System: Pre-requisites for processing not met (files, API key, or model).", role="error"); return

        self.update_conversation_history("System: Starting analysis of spec sheets...", role="system")
        self.spec_sheet_1_text = self.extract_text_from_pdf(self.spec_sheet_1_path)
        if not self.spec_sheet_1_text: self.update_conversation_history(f"System: Halting. Text extraction failed for {os.path.basename(self.spec_sheet_1_path)}.", role="error"); return
        spec1_img_folder = os.path.join(self.temp_image_dir, f"{os.path.splitext(os.path.basename(self.spec_sheet_1_path))[0]}_imgs_{len(os.listdir(self.temp_image_dir))}")
        self.spec_sheet_1_image_paths = self.extract_images_from_pdf(self.spec_sheet_1_path, spec1_img_folder)

        self.spec_sheet_2_text = self.extract_text_from_pdf(self.spec_sheet_2_path)
        if not self.spec_sheet_2_text: self.update_conversation_history(f"System: Halting. Text extraction failed for {os.path.basename(self.spec_sheet_2_path)}.", role="error"); return
        spec2_img_folder = os.path.join(self.temp_image_dir, f"{os.path.splitext(os.path.basename(self.spec_sheet_2_path))[0]}_imgs_{len(os.listdir(self.temp_image_dir))}")
        self.spec_sheet_2_image_paths = self.extract_images_from_pdf(self.spec_sheet_2_path, spec2_img_folder)

        summary_msg_ui = (f"System: Analysis inputs:\n- Spec 1: {os.path.basename(self.spec_sheet_1_path)} ({len(self.spec_sheet_1_image_paths)} images)\n"
                       f"- Spec 2: {os.path.basename(self.spec_sheet_2_path)} ({len(self.spec_sheet_2_image_paths)} images)")
        self.update_conversation_history(summary_msg_ui, role="system")

        initial_prompt_text_for_ai_log = "User: Analyze the following two component specification sheets.\n"
        initial_prompt_text_for_ai_log += f"Spec 1 ({os.path.basename(self.spec_sheet_1_path)}): {self.spec_sheet_1_text[:200]}...\n"
        if self.spec_sheet_1_image_paths: initial_prompt_text_for_ai_log += f"({len(self.spec_sheet_1_image_paths)} images included)\n"
        initial_prompt_text_for_ai_log += f"Spec 2 ({os.path.basename(self.spec_sheet_2_path)}): {self.spec_sheet_2_text[:200]}...\n"
        if self.spec_sheet_2_image_paths: initial_prompt_text_for_ai_log += f"({len(self.spec_sheet_2_image_paths)} images included)\n"
        initial_prompt_text_for_ai_log += "Request: Identify component type, crucial parameters, pin compatibility, and key differences."

        self._add_to_ai_history('user', initial_prompt_text_for_ai_log)


        prompt_parts_for_genai = ["You are an expert electronics component analyst...", "\n--- Spec Sheet 1 Text ---", self.spec_sheet_1_text]
        for img_path in self.spec_sheet_1_image_paths:
            try: prompt_parts_for_genai.append(Image.open(img_path))
            except Exception as e: self.update_conversation_history(f"System: Error loading image {img_path}. Skip. Err: {e}", role="error")
        prompt_parts_for_genai.extend(["\n--- Spec Sheet 2 Text ---", self.spec_sheet_2_text])
        for img_path in self.spec_sheet_2_image_paths:
            try: prompt_parts_for_genai.append(Image.open(img_path))
            except Exception as e: self.update_conversation_history(f"System: Error loading image {img_path}. Skip. Err: {e}", role="error")
        prompt_parts_for_genai.append("\n--- Analysis Request ---"
                            "\n1. Identify component type for each."
                            "\n2. List crucial parameters for comparison."
                            "\n3. Compare pin-to-pin compatibility (compatible, potentially, or not, and why)."
                            "\n4. List key spec differences (electrical, physical) structured."
                            "\nPresent analysis clearly and concisely.")
        self.send_to_ai(prompt_parts_for_genai, is_initial_analysis=True)

    def send_to_ai(self, prompt_parts, is_initial_analysis=False):
        if not self.model:
            self.update_conversation_history("System: AI model not available for sending request.", role="error"); return

        active_model_name = self.model.model_name if hasattr(self.model, 'model_name') else "Unknown Model"
        try:
            self.send_button.config(state=tk.DISABLED); self.user_input_entry.config(state=tk.DISABLED)
            self.update_conversation_history(f"System: Sending request to AI ({active_model_name})... May take time.", role="system")

            response = self.model.generate_content(prompt_parts, request_options={'timeout': 600})

            raw_ai_response_text = ""
            if response.prompt_feedback and response.prompt_feedback.block_reason:
                raw_ai_response_text = f"AI Error - Prompt was blocked. Reason: {response.prompt_feedback.block_reason}"
                self.update_conversation_history(f"System: {raw_ai_response_text}", role="error")
            elif not response.candidates or not hasattr(response, 'text') or not response.text:
                raw_ai_response_text = "AI response was empty or had no content."
                self.update_conversation_history(f"System: AI ({active_model_name}): {raw_ai_response_text}", role="system")
            else:
                raw_ai_response_text = response.text
                self.update_conversation_history(f"AI ({active_model_name}): {raw_ai_response_text}", role="ai") # Formatted for display

            if is_initial_analysis:
                self._add_to_ai_history('model', raw_ai_response_text) # Log raw response for AI history
                self.chat_session = None

        except Exception as e:
            error_message = f"System: Error during AI content generation with {active_model_name}: {e}"
            self.update_conversation_history(error_message, role="error")
            print(f"DEBUG: AI Error: {error_message}")
            self._add_to_ai_history('model', f"Error during generation: {e}")

            if isinstance(e, (google_exceptions.PermissionDenied, google_exceptions.Unauthenticated)):
                self.api_key_configured = False
            if isinstance(e, (google_exceptions.InvalidArgument, ValueError, BlockedPromptException, StopCandidateException, google_exceptions.NotFound, google_exceptions.PermissionDenied)):
                self.update_conversation_history(f"System: Current AI model instance ({active_model_name}) has been reset due to a critical error.", role="system")
                self.model = None; self.chat_session = None; self.ai_history = []
        finally:
            self._update_ui_for_ai_status()

def main():
    try:
        s = ttk.Style(); available_themes = s.theme_names()
        if "clam" in available_themes: s.theme_use("clam")
        elif "aqua" in available_themes: s.theme_use("aqua")
        elif "vista" in available_themes: s.theme_use("vista")
    except Exception: pass
    root = tk.Tk()
    app = ComponentComparatorAI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
