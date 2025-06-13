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
        self.root.geometry("850x700") # Set a default window size

        # --- Internal State Variables ---
        self.spec_sheet_1_path = None
        self.spec_sheet_1_text = None
        self.spec_sheet_1_image_paths = []
        self.spec_sheet_2_path = None
        self.spec_sheet_2_text = None
        self.spec_sheet_2_image_paths = []

        self.model = None
        self.chat_session = None
        self.conversation_log = []
        self.api_key_configured = False
        self.model_options_list = [] # Initialize model options list

        self.temp_image_dir = "temp_images"
        if not os.path.exists(self.temp_image_dir):
            try:
                os.makedirs(self.temp_image_dir)
            except OSError as e:
                error_message = f"Critical Error: Cannot create temporary directory {self.temp_image_dir}: {e}"
                print(error_message)

        self._setup_ui(root)
        self._configure_ai()
        if self.api_key_configured:
            self._initialize_model()


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
        self.model_combobox = ttk.Combobox(root, textvariable=self.model_var, state="readonly")

        self.model_options_list = [ # Assign to instance variable
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
        if self.model_options_list:
            self.model_combobox.current(0)
        else:
            print("Error: model_options_list is empty. Cannot set default Combobox selection.")
            self.model_combobox.set("No models available")

        self.model_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.model_combobox.bind("<<ComboboxSelected>>", self._on_model_selected)

        self.history_label = ttk.Label(root, text="Conversation History:")
        self.history_label.grid(row=3, column=0, columnspan=2, padx=10, pady=(10,0), sticky="w")
        self.conversation_history = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=15, width=80)
        self.conversation_history.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")
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
            self.check_and_process_spec_sheets()
        else:
            self.spec_sheet_1_label.config(text=f"File 1: {os.path.basename(self.spec_sheet_1_path) if self.spec_sheet_1_path else 'None'}")

    def load_spec_sheet_2(self):
        filepath = filedialog.askopenfilename(title="Select Spec Sheet 2 (PDF)", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if filepath:
            self.spec_sheet_2_path = filepath
            self.spec_sheet_2_label.config(text=f"File 2: {os.path.basename(filepath)}")
            self.check_and_process_spec_sheets()
        else:
            self.spec_sheet_2_label.config(text=f"File 2: {os.path.basename(self.spec_sheet_2_path) if self.spec_sheet_2_path else 'None'}")

    def update_conversation_history(self, message, is_internal_log_message=False):
        if hasattr(self, 'conversation_history') and self.conversation_history:
            self.conversation_history.config(state=tk.NORMAL)
            self.conversation_history.insert(tk.END, message + "\n")
            self.conversation_history.see(tk.END)
            self.conversation_history.config(state=tk.DISABLED)
        if not is_internal_log_message:
            self.conversation_log.append(message)

    def _update_ui_for_ai_status(self, api_key_configured=None, model_initialized=None):
        if not hasattr(self, 'send_button'): return
        is_api_key_ready = api_key_configured if api_key_configured is not None else self.api_key_configured
        is_model_ready = model_initialized if model_initialized is not None else (self.model is not None)
        can_perform_ai_ops = is_api_key_ready and is_model_ready
        self.send_button.config(state=tk.NORMAL if can_perform_ai_ops else tk.DISABLED)
        self.user_input_entry.config(state=tk.NORMAL if can_perform_ai_ops else tk.DISABLED)
        if hasattr(self, 'model_combobox'): self.model_combobox.config(state="readonly")

    def _configure_ai(self):
        try:
            api_key = os.environ.get("GOOGLE_API_KEY")
            if not api_key:
                self.update_conversation_history("System: Error - GOOGLE_API_KEY environment variable not set. AI features disabled.")
                self.api_key_configured = False
            else:
                genai.configure(api_key=api_key)
                self.update_conversation_history("System: Generative AI configured successfully.")
                self.api_key_configured = True
        except Exception as e:
            self.update_conversation_history(f"System: Error configuring Generative AI SDK - {e}")
            self.api_key_configured = False
        finally:
            self._update_ui_for_ai_status(model_initialized=(self.model is not None))

    def _on_model_selected(self, event=None):
        selected_model_name = self.model_var.get()
        model_initialized_successfully = self._initialize_model(selected_model_name) # Pass explicit name
        if model_initialized_successfully:
            self.check_and_process_spec_sheets()

    def _initialize_model(self, model_name=None):
        # Determine the source of model_name and log it
        if model_name is None:
            fetched_model_name = self.model_combobox.get()
            self.update_conversation_history(f"System: _initialize_model called without explicit model_name. Fetched from Combobox: '{fetched_model_name}'")
            model_name = fetched_model_name # Use the fetched name
        else:
            self.update_conversation_history(f"System: _initialize_model called with explicit model_name: '{model_name}'")

        # Validate the obtained model_name against the instance's model_options_list
        if not model_name or model_name not in self.model_options_list:
            self.update_conversation_history(f"System: Invalid or empty model name ('{model_name}') for initialization. Available options: {len(self.model_options_list)}. Initialization aborted.")
            self.model = None
            self.chat_session = None
            self._update_ui_for_ai_status(model_initialized=False)
            return False

        # Proceed if model_name is valid
        if self.model and self.model.model_name == model_name and self.api_key_configured:
            self._update_ui_for_ai_status(model_initialized=True); return True

        if not self.api_key_configured:
            self.update_conversation_history("System: Cannot initialize model - API key not configured.")
            self.model = None; self.chat_session = None # Ensure reset
            self._update_ui_for_ai_status(model_initialized=False); return False

        self.update_conversation_history(f"System: Attempting to initialize AI model: {model_name}...")
        try:
            self.model = genai.GenerativeModel(model_name); self.chat_session = None
            self.update_conversation_history(f"System: Successfully initialized AI model: {model_name}")
            self._update_ui_for_ai_status(model_initialized=True); return True
        except Exception as e:
            self.model = None; self.chat_session = None
            self.update_conversation_history(f"System: Error initializing AI model {model_name}: {e}")
            self._update_ui_for_ai_status(model_initialized=False); return False

    def get_selected_model_name(self): return self.model_var.get()

    def send_user_query(self):
        if not self.model or not self.api_key_configured:
            if not self._initialize_model():
                self.update_conversation_history("System: AI Model not initialized. Select model & ensure API key is set."); return
        if not self.model: return

        user_text = self.user_input_entry.get().strip()
        if not user_text: return
        self.update_conversation_history(f"User: {user_text}"); self.user_input_entry.delete(0, tk.END)
        try:
            self.send_button.config(state=tk.DISABLED); self.user_input_entry.config(state=tk.DISABLED)
            if not self.chat_session:
                self.update_conversation_history(f"System: Starting new chat session with {self.model.model_name}...")
                try: self.chat_session = self.model.start_chat(history=[])
                except Exception as e: self.update_conversation_history(f"System: Error starting chat: {e}"); self._update_ui_for_ai_status(); return
            self.update_conversation_history(f"System: Sending to AI ({self.model.model_name})..."); response = self.chat_session.send_message(user_text)
            self.update_conversation_history(f"AI ({self.model.model_name}): {response.text}")
        except Exception as e:
            self.update_conversation_history(f"System: Error during AI interaction: {e}")
            if isinstance(e, google_exceptions.PermissionDenied): self.api_key_configured = False
            if isinstance(e, (google_exceptions.InvalidArgument, ValueError, BlockedPromptException, StopCandidateException)): self.model = None; self.chat_session = None
        finally: self._update_ui_for_ai_status()

    def download_history(self):
        if not self.conversation_log: self.update_conversation_history("System: History empty."); return
        try: filepath = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")], title="Save History")
        except Exception as e: self.update_conversation_history(f"System: Error opening save dialog: {e}"); return
        if not filepath: self.update_conversation_history("System: Download cancelled."); return
        try:
            doc = docx.Document(); doc.add_heading("Component Comparator AI Chat History", level=1)
            if self.spec_sheet_1_path: doc.add_paragraph(f"Spec Sheet 1: {os.path.basename(self.spec_sheet_1_path)}")
            if self.spec_sheet_2_path: doc.add_paragraph(f"Spec Sheet 2: {os.path.basename(self.spec_sheet_2_path)}")
            model_name = self.model.model_name if self.model and hasattr(self.model, 'model_name') else "N/A"
            doc.add_paragraph(f"AI Model (last used): {model_name}"); doc.add_paragraph("-" * 20)
            for entry in self.conversation_log: doc.add_paragraph(entry)
            doc.save(filepath); self.update_conversation_history(f"System: History downloaded to {filepath}")
        except Exception as e: self.update_conversation_history(f"System: Error downloading: {e}"); print(f"Error saving .docx: {e}")

    def clear_all(self):
        self.spec_sheet_1_label.config(text="File 1: None"); self.spec_sheet_1_path = None; self.spec_sheet_1_text = None; self.spec_sheet_1_image_paths = []
        self.spec_sheet_2_label.config(text="File 2: None"); self.spec_sheet_2_path = None; self.spec_sheet_2_text = None; self.spec_sheet_2_image_paths = []
        if self.model_options_list: self.model_combobox.current(0) # Reset to first option if list exists
        else: self.model_combobox.set("") # Clear if no options
        current_model_name_before_reset = self.model_var.get() # Get before self.model is None
        self.model = None; self.chat_session = None; self.conversation_log = []
        if hasattr(self, 'conversation_history'): self.conversation_history.config(state=tk.NORMAL); self.conversation_history.delete(1.0, tk.END)
        self._configure_ai()
        if self.api_key_configured: self._initialize_model(current_model_name_before_reset)
        else: self._update_ui_for_ai_status(model_initialized=False)
        if hasattr(self, 'user_input_entry'): self.user_input_entry.delete(0, tk.END)
        try:
            if os.path.exists(self.temp_image_dir): shutil.rmtree(self.temp_image_dir)
            os.makedirs(self.temp_image_dir)
        except OSError as e: self.update_conversation_history(f"System: Error cleaning temp dir: {e}")
        print("Clear All: App state reset.")

    def extract_text_from_pdf(self, filepath):
        if not filepath or not os.path.exists(filepath): self.update_conversation_history(f"System: PDF not found: {os.path.basename(filepath or 'Unknown')}"); return ""
        try:
            self.update_conversation_history(f"System: Extracting text from {os.path.basename(filepath)}...")
            with fitz.open(filepath) as doc: text = "".join(page.get_text() for page in doc)
            self.update_conversation_history(f"System: Text extraction successful: {os.path.basename(filepath)}."); return text
        except Exception as e: self.update_conversation_history(f"System: Error extracting text from {os.path.basename(filepath)}: {e}"); return ""

    def extract_images_from_pdf(self, filepath, output_folder):
        if not filepath or not os.path.exists(filepath): self.update_conversation_history(f"System: PDF not found: {os.path.basename(filepath or 'Unknown')}"); return []
        paths = []
        try:
            self.update_conversation_history(f"System: Extracting images from {os.path.basename(filepath)}...")
            if not os.path.exists(output_folder): os.makedirs(output_folder)
            with fitz.open(filepath) as doc:
                for i, page in enumerate(doc):
                    for j, img_info in enumerate(page.get_images(full=True)):
                        xref = img_info[0]
                        try: base = doc.extract_image(xref)
                        except Exception as e: self.update_conversation_history(f"System: Error extracting img xref {xref} pg {i+1}. Skip. Err: {e}"); continue
                        img_bytes, ext = base["image"], base["ext"]
                        path = os.path.join(output_folder, f"pg{i+1}_img{j+1}.{ext}")
                        try:
                            with open(path, "wb") as f: f.write(img_bytes)
                            paths.append(path)
                        except IOError as e: self.update_conversation_history(f"System: IOError saving image {path}. Error: {e}")
            msg = f"System: Extracted {len(paths)} images from {os.path.basename(filepath)}." if paths else f"System: No images found in {os.path.basename(filepath)}."
            self.update_conversation_history(msg); return paths
        except Exception as e: self.update_conversation_history(f"System: Error extracting images from {os.path.basename(filepath)}: {e}"); return []

    def check_and_process_spec_sheets(self):
        if not (self.spec_sheet_1_path and self.spec_sheet_2_path): return
        self.update_conversation_history("System: Both spec sheets loaded. Verifying AI model status...")
        self.conversation_log = []
        if hasattr(self, 'conversation_history'):
            self.conversation_history.config(state=tk.NORMAL); self.conversation_history.delete(1.0, tk.END)
        self._configure_ai()
        if not self.api_key_configured:
            self.update_conversation_history("System: API Key not configured. Cannot process specs."); return
        if not self.model:
            self.update_conversation_history("System: No AI model active. Attempting to initialize from selection...")
            if not self._initialize_model():
                self.update_conversation_history("System: AI model initialization failed. Please select a model or ensure API key is correct to process specs.")
                return
        self.process_spec_sheets()

    def process_spec_sheets(self):
        if not self.model:
            self.update_conversation_history("System: Critical - process_spec_sheets called without initialized model.")
            if not self._initialize_model(): self._update_ui_for_ai_status(model_initialized=False); return
        if not self.spec_sheet_1_path or not self.spec_sheet_2_path or not self.api_key_configured:
            self.update_conversation_history("System: Pre-requisites not met for processing (files, API key, or model)."); return

        self.update_conversation_history("System: Starting analysis of spec sheets...")
        self.spec_sheet_1_text = self.extract_text_from_pdf(self.spec_sheet_1_path)
        if not self.spec_sheet_1_text: self.update_conversation_history(f"System: Halting. Text extraction failed for {os.path.basename(self.spec_sheet_1_path)}."); return
        spec1_img_folder = os.path.join(self.temp_image_dir, f"{os.path.splitext(os.path.basename(self.spec_sheet_1_path))[0]}_imgs_{len(os.listdir(self.temp_image_dir))}")
        self.spec_sheet_1_image_paths = self.extract_images_from_pdf(self.spec_sheet_1_path, spec1_img_folder)

        self.spec_sheet_2_text = self.extract_text_from_pdf(self.spec_sheet_2_path)
        if not self.spec_sheet_2_text: self.update_conversation_history(f"System: Halting. Text extraction failed for {os.path.basename(self.spec_sheet_2_path)}."); return
        spec2_img_folder = os.path.join(self.temp_image_dir, f"{os.path.splitext(os.path.basename(self.spec_sheet_2_path))[0]}_imgs_{len(os.listdir(self.temp_image_dir))}")
        self.spec_sheet_2_image_paths = self.extract_images_from_pdf(self.spec_sheet_2_path, spec2_img_folder)

        summary_msg = (f"System: Analysis inputs:\n- Spec 1: {os.path.basename(self.spec_sheet_1_path)} ({len(self.spec_sheet_1_image_paths)} images)\n"
                       f"- Spec 2: {os.path.basename(self.spec_sheet_2_path)} ({len(self.spec_sheet_2_image_paths)} images)")
        self.update_conversation_history(summary_msg, is_internal_log_message=True)

        prompt_parts = ["You are an expert electronics component analyst...", "\n--- Spec Sheet 1 Text ---", self.spec_sheet_1_text]
        for img_path in self.spec_sheet_1_image_paths:
            try: prompt_parts.append(Image.open(img_path))
            except Exception as e: self.update_conversation_history(f"System: Error loading image {img_path}. Skip. Err: {e}")
        prompt_parts.extend(["\n--- Spec Sheet 2 Text ---", self.spec_sheet_2_text])
        for img_path in self.spec_sheet_2_image_paths:
            try: prompt_parts.append(Image.open(img_path))
            except Exception as e: self.update_conversation_history(f"System: Error loading image {img_path}. Skip. Err: {e}")
        prompt_parts.append("\n--- Analysis Request ---"
                            "\n1. Identify component type for each."
                            "\n2. List crucial parameters for comparison."
                            "\n3. Compare pin-to-pin compatibility (compatible, potentially, or not, and why)."
                            "\n4. List key spec differences (electrical, physical) structured."
                            "\nPresent analysis clearly and concisely.")
        self.send_to_ai(prompt_parts, is_initial_analysis=True)

    def send_to_ai(self, prompt_parts, is_initial_analysis=False):
        if not self.model: self.update_conversation_history("System: AI model not available."); return
        model_name_for_log = self.model.model_name if hasattr(self.model, 'model_name') else "Unknown Model"
        try:
            self.send_button.config(state=tk.DISABLED); self.user_input_entry.config(state=tk.DISABLED)
            self.update_conversation_history(f"System: Sending request to AI ({model_name_for_log})... May take time.")
            response = self.model.generate_content(prompt_parts)
            if response.prompt_feedback and response.prompt_feedback.block_reason:
                self.update_conversation_history(f"System: AI Error - Prompt blocked. Reason: {response.prompt_feedback.block_reason}")
            elif not response.candidates or not response.text:
                self.update_conversation_history(f"System: AI ({model_name_for_log}): Received no content or empty response.")
            else: self.update_conversation_history(f"AI ({model_name_for_log}): {response.text}")
            if is_initial_analysis: self.chat_session = None
        except Exception as e:
            self.update_conversation_history(f"System: Error during AI content generation: {e}")
            if isinstance(e, google_exceptions.PermissionDenied): self.api_key_configured = False
            if isinstance(e, (google_exceptions.InvalidArgument, ValueError, BlockedPromptException, StopCandidateException)): self.model = None; self.chat_session = None
        finally: self._update_ui_for_ai_status()

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
