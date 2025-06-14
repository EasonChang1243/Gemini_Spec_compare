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

class Tooltip:
    """
    Create a tooltip for a given widget.
    """
    def __init__(self, widget, text_callback):
        self.widget = widget
        self.text_callback = text_callback
        self.tip_window = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        """Display tooltip"""
        if self.tip_window or not hasattr(self, 'text_callback'):
            return

        text = self.text_callback()
        if not text:
            return

        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 1

        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")

        label = tk.Label(tw, text=text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hide_tip(self, event=None):
        """Hide tooltip"""
        if self.tip_window:
            self.tip_window.destroy()
        self.tip_window = None

class ComponentComparatorAI:
    """
    Main application class for the Component Comparator AI.
    Manages the UI, file loading, PDF processing, AI interaction,
    and conversation history.
    """
    def __init__(self, root):
        """
        Initializes the application UI and internal state.
        """
        self.root = root
        self.root.title("Component Comparator AI")
        self.root.geometry("850x870")

        if load_dotenv(): print("DEBUG: Loaded environment variables from .env file.")
        else: print("DEBUG: No .env file found or python-dotenv not available/failed to load.")

        self.spec_sheet_1_path = None; self.spec_sheet_1_text = None; self.spec_sheet_1_image_paths = []
        self.mfg_pn_var_1 = tk.StringVar()
        self.spec_sheet_2_path = None; self.spec_sheet_2_text = None; self.spec_sheet_2_image_paths = []
        self.mfg_pn_var_2 = tk.StringVar()

        self.model = None; self.chat_session = None; self.conversation_log = []; self.ai_history = []
        self.api_key_configured = False; self.model_options_list = []
        self.placeholder_text = "Select AI Model (after loading files)"; self.model_initializing = False
        self.start_comparison_button = None
        self.upload_image_button = None
        self.pending_user_image_path = None
        self.pending_user_image_pil = None
        self.translate_to_chinese_var = tk.BooleanVar(value=True)


        self.temp_image_dir = "temp_images"; self._create_temp_image_dir()
        self._setup_ui(root); self._configure_ai()

        self.update_conversation_history(
            "System: Welcome! Please load two PDF specification sheets to compare. " +
            "After loading both files, you will be able to select an AI model.", role="system"
        )
        self._update_ui_for_ai_status(api_key_configured=self.api_key_configured, model_initialized=False)


    def _create_temp_image_dir(self):
        if not os.path.exists(self.temp_image_dir):
            try: os.makedirs(self.temp_image_dir); print(f"DEBUG: Created temp dir: {self.temp_image_dir}")
            except OSError as e: print(f"Critical Error creating temp dir {self.temp_image_dir}: {e}"); self.update_conversation_history(f"System: Error creating temp folder: {e}", role="error")

    def _setup_ui(self, root):
        current_row = 0
        # File 1 & MFG P/N 1
        self.spec_sheet_1_label = ttk.Label(root, text="File 1: None")
        self.spec_sheet_1_label.grid(row=current_row, column=0, padx=10, pady=2, sticky="w")
        self.load_spec_sheet_1_button = ttk.Button(root, text="Load Spec Sheet 1", command=self.load_spec_sheet_1)
        self.load_spec_sheet_1_button.grid(row=current_row, column=1, padx=5, pady=2, sticky="ew")
        current_row += 1
        self.mfg_pn_label_1 = ttk.Label(root, text="MFG P/N 1:")
        self.mfg_pn_label_1.grid(row=current_row, column=0, sticky=tk.W, padx=10, pady=2)
        self.mfg_pn_entry_1 = ttk.Entry(root, textvariable=self.mfg_pn_var_1, width=30)
        self.mfg_pn_entry_1.grid(row=current_row, column=1, sticky=tk.EW, padx=5, pady=2)
        current_row += 1
        # File 2 & MFG P/N 2
        self.spec_sheet_2_label = ttk.Label(root, text="File 2: None")
        self.spec_sheet_2_label.grid(row=current_row, column=0, padx=10, pady=2, sticky="w")
        self.load_spec_sheet_2_button = ttk.Button(root, text="Load Spec Sheet 2", command=self.load_spec_sheet_2)
        self.load_spec_sheet_2_button.grid(row=current_row, column=1, padx=5, pady=2, sticky="ew")
        current_row += 1
        self.mfg_pn_label_2 = ttk.Label(root, text="MFG P/N 2:")
        self.mfg_pn_label_2.grid(row=current_row, column=0, sticky=tk.W, padx=10, pady=2)
        self.mfg_pn_entry_2 = ttk.Entry(root, textvariable=self.mfg_pn_var_2, width=30)
        self.mfg_pn_entry_2.grid(row=current_row, column=1, sticky=tk.EW, padx=5, pady=2)
        current_row += 1
        # Model Selection
        ttk.Label(root, text="Select AI Model:").grid(row=current_row, column=0, padx=10, pady=5, sticky="w")
        self.model_var = tk.StringVar()
        self.model_combobox = ttk.Combobox(root, textvariable=self.model_var)
        self.model_options_list = ["models/gemini-1.0-pro-vision-latest", "models/gemini-pro-vision","models/gemini-1.5-flash-latest", "models/gemini-1.5-flash","models/gemini-1.5-flash-002", "models/gemini-1.5-flash-8b","models/gemini-1.5-flash-8b-001", "models/gemini-1.5-flash-8b-latest","models/gemini-2.5-flash-preview-04-17", "models/gemini-2.5-flash-preview-05-20","models/gemini-2.5-flash-preview-04-17-thinking", "models/gemini-2.0-flash-exp","models/gemini-2.0-flash", "models/gemini-2.0-flash-001","models/gemini-2.0-flash-exp-image-generation", "models/gemini-2.0-flash-lite-001","models/gemini-2.0-flash-lite", "models/gemini-2.0-flash-lite-preview-02-05","models/gemini-2.0-flash-lite-preview", "models/gemini-2.0-flash-thinking-exp-01-21","models/gemini-2.0-flash-thinking-exp", "models/gemini-2.0-flash-thinking-exp-1219","models/learnlm-2.0-flash-experimental", "models/gemma-3-1b-it","models/gemma-3-4b-it", "models/gemma-3-12b-it","models/gemma-3-27b-it", "models/gemma-3n-e4b-it"]
        self.model_combobox['values'] = self.model_options_list
        self.model_combobox.set(self.placeholder_text); self.model_combobox.state(["disabled"])
        self.model_combobox.grid(row=current_row, column=1, padx=5, pady=5, sticky="ew")
        self.model_combobox.bind("<<ComboboxSelected>>", self._on_model_selected)
        def get_tooltip_text(): return self.model_combobox.get() if self.model_combobox.get() != self.placeholder_text else None
        self.combobox_tooltip = Tooltip(self.model_combobox, get_tooltip_text)
        current_row += 1
        # Conversation History
        ttk.Label(root, text="Conversation History:").grid(row=current_row, column=0, columnspan=2, padx=10, pady=(10,0), sticky="w")
        current_row += 1
        self.conversation_history = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=10, width=80)
        self.conversation_history.grid(row=current_row, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")
        self.conversation_history.tag_configure("user_message", foreground="blue", font=('Arial', 10))
        self.conversation_history.tag_configure("ai_message", foreground="#008800", font=('Arial', 10))
        self.conversation_history.tag_configure("system_message", foreground="#550055", font=('Arial', 10, 'italic'))
        self.conversation_history.tag_configure("error_message", foreground="red", font=('Arial', 10, 'bold'))
        self.conversation_history.config(state=tk.DISABLED)
        root.grid_rowconfigure(current_row, weight=1)
        current_row += 1
        # Comparison Treeview
        self.treeview_frame = ttk.LabelFrame(root, text="Comparison Details")
        self.treeview_frame.grid(row=current_row, column=0, columnspan=2, sticky="nsew", padx=10, pady=5)
        columns = ("parameter", "component1", "component2", "notes")
        self.comparison_treeview = ttk.Treeview(self.treeview_frame, columns=columns, show="headings", height=8)
        self.comparison_treeview.heading("parameter", text="Parameter"); self.comparison_treeview.column("parameter", anchor=tk.W, width=150)
        self.comparison_treeview.heading("component1", text="Comp 1 Val"); self.comparison_treeview.column("component1", anchor=tk.W, width=200)
        self.comparison_treeview.heading("component2", text="Comp 2 Val"); self.comparison_treeview.column("component2", anchor=tk.W, width=200)
        self.comparison_treeview.heading("notes", text="Notes"); self.comparison_treeview.column("notes", anchor=tk.W, width=250)
        vsb = ttk.Scrollbar(self.treeview_frame, orient="vertical", command=self.comparison_treeview.yview)
        hsb = ttk.Scrollbar(self.treeview_frame, orient="horizontal", command=self.comparison_treeview.xview)
        self.comparison_treeview.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.comparison_treeview.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns"); hsb.grid(row=1, column=0, sticky="ew")
        self.treeview_frame.grid_rowconfigure(0, weight=1); self.treeview_frame.grid_columnconfigure(0, weight=1)
        root.grid_rowconfigure(current_row, weight=1)
        current_row += 1

        # --- User Input Controls Frame ---
        self.user_input_controls_frame = ttk.Frame(root)
        self.user_input_controls_frame.grid(row=current_row, column=0, columnspan=2, sticky="ew", padx=5, pady=0)
        self.user_input_controls_frame.columnconfigure(0, weight=1)
        self.user_input_controls_frame.columnconfigure(1, weight=0)
        self.user_input_controls_frame.columnconfigure(2, weight=0)
        self.user_input_controls_frame.columnconfigure(3, weight=0)

        ttk.Label(self.user_input_controls_frame, text="Your Message:").grid(row=0, column=0, columnspan=4, padx=5, pady=(5,0), sticky="w")

        self.user_input_entry = ttk.Entry(self.user_input_controls_frame, width=70)
        self.user_input_entry.grid(row=1, column=0, padx=(5,0), pady=5, sticky="ew")

        self.translate_to_chinese_checkbutton = ttk.Checkbutton(
            self.user_input_controls_frame,
            text="AI replies in Chinese",
            variable=self.translate_to_chinese_var,
            onvalue=True,
            offvalue=False
        )
        self.translate_to_chinese_checkbutton.grid(row=1, column=1, padx=5, pady=5, sticky="e")

        self.upload_image_button = ttk.Button(self.user_input_controls_frame, text="Attach Image", command=self.on_upload_image)
        self.upload_image_button.grid(row=1, column=2, padx=5, pady=5, sticky="e")

        self.send_button = ttk.Button(self.user_input_controls_frame, text="Send", command=self.send_user_query)
        self.send_button.grid(row=1, column=3, padx=(0,5), pady=5, sticky="e")
        current_row += 1

        # --- Action Buttons Frame ---
        self.action_buttons_frame = ttk.Frame(root)
        self.action_buttons_frame.grid(row=current_row, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        self.action_buttons_frame.columnconfigure(0, weight=1)
        self.action_buttons_frame.columnconfigure(1, weight=1)
        self.action_buttons_frame.columnconfigure(2, weight=1)
        self.start_comparison_button = ttk.Button(self.action_buttons_frame, text="Start Detailed Comparison", command=self.on_start_detailed_comparison, state=tk.DISABLED)
        self.start_comparison_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, pady=5)
        self.download_history_button = ttk.Button(self.action_buttons_frame, text="Download History", command=self.download_history)
        self.download_history_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, pady=5)
        self.clear_all_button = ttk.Button(self.action_buttons_frame, text="Clear All", command=self.clear_all)
        self.clear_all_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, pady=5)

        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=0)

    def _parse_initial_analysis_response(self, response_text: str) -> dict:
        """Parses the structured AI response from the initial analysis."""
        data = {
            "component1_type": "Unknown",
            "component2_type": "Unknown",
            "functionally_similar": "Unknown",
            "is_similar_flag": False, # Default to False
            "mfg_pn1": "Not Found",
            "mfg_pn2": "Not Found"
        }
        lines = response_text.split('\n')
        for line in lines:
            if line.startswith("Component1_Type:"):
                data["component1_type"] = line.split(":", 1)[1].strip()
            elif line.startswith("Component2_Type:"):
                data["component2_type"] = line.split(":", 1)[1].strip()
            elif line.startswith("Functionally_Similar:"):
                similarity_text = line.split(":", 1)[1].strip()
                data["functionally_similar"] = similarity_text
                if similarity_text.lower().startswith("yes"):
                    data["is_similar_flag"] = True
            elif line.startswith("MFG_PN1:"):
                data["mfg_pn1"] = line.split(":", 1)[1].strip()
            elif line.startswith("MFG_PN2:"):
                data["mfg_pn2"] = line.split(":", 1)[1].strip()

        print(f"DEBUG: Parsed initial analysis: {data}")
        return data

    def on_upload_image(self):
        filetypes = [('Image files', '*.png *.jpg *.jpeg *.bmp *.gif *.webp'), ('All files', '*.*')]
        filepath = filedialog.askopenfilename(title="Select an Image for AI Analysis", filetypes=filetypes)
        if filepath:
            try:
                img = Image.open(filepath)
                self.pending_user_image_path = filepath
                self.pending_user_image_pil = img
                image_name = os.path.basename(filepath)
                self.update_conversation_history(f"System: Image '{image_name}' attached. It will be sent with your next message.", role="system")
            except FileNotFoundError:
                self.update_conversation_history(f"System: Error - Image file not found at {filepath}", role="error")
                self.pending_user_image_path = None; self.pending_user_image_pil = None
            except UnidentifiedImageError:
                self.update_conversation_history(f"System: Error - Cannot identify image file. Not a valid image format? File: {filepath}", role="error")
                self.pending_user_image_path = None; self.pending_user_image_pil = None
            except Exception as e:
                self.update_conversation_history(f"System: Error processing image {filepath}: {e}", role="error")
                self.pending_user_image_path = None; self.pending_user_image_pil = None

    def on_start_detailed_comparison(self):
        self.update_conversation_history("System: 'Start Detailed Comparison' initiated...", role="system")
        if hasattr(self, 'start_comparison_button'): self.start_comparison_button.config(state=tk.DISABLED)

        if not self.model:
            self.update_conversation_history("System: No AI model initialized. Please select a model.", role="error")
            return
        if not (self.spec_sheet_1_text and self.spec_sheet_2_text):
            self.update_conversation_history("System: Both spec sheets must be loaded and processed.", role="error")
            return

        mfg_pn1 = self.mfg_pn_var_1.get() if hasattr(self, 'mfg_pn_var_1') else "N/A"
        mfg_pn2 = self.mfg_pn_var_2.get() if hasattr(self, 'mfg_pn_var_2') else "N/A"

        user_prompt_for_history = (
            f"User: Detailed comparison request for MFG P/N 1: {mfg_pn1 if mfg_pn1 else 'N/A'} "
            f"(from {os.path.basename(self.spec_sheet_1_path or 'File 1')}) "
            f"vs MFG P/N 2: {mfg_pn2 if mfg_pn2 else 'N/A'} "
            f"(from {os.path.basename(self.spec_sheet_2_path or 'File 2')}). "
            "Focus: Crucial parameters, differences table, operating temp, SMT compatibility."
        )

        detailed_prompt_parts_for_genai = [
            f"Please perform a detailed comparison of two electronic components previously analyzed (initial analysis provided context on component types, text, and images).",
            f"Component 1 is identified by MFG P/N: {mfg_pn1 if mfg_pn1 else 'N/A'}.",
            f"Component 2 is identified by MFG P/N: {mfg_pn2 if mfg_pn2 else 'N/A'}.",
            "Focus on the following aspects for your detailed comparison:",
            "1. Crucial electrical and physical parameters relevant for comparing these specific component types (list them).",
            "2. List all key specification differences in a clear, concise markdown table format.",
            "3. Explicitly state their Operating Temperature ranges.",
            "4. Assess SMT Compatibility: Can Component 2's package (based on its description in provided text/images) likely be SMT'd onto Component 1's typical PCB footprint? Consider common package names and pin counts. State any assumptions clearly."
        ]

        self.update_conversation_history("System: Sending detailed comparison request to AI. This may take some time.", role="system")
        print(f"DEBUG: Detailed Comparison Prompt Parts being sent to AI (text only shown):\n{detailed_prompt_parts_for_genai}")

        ai_response_text = self.send_to_ai(
            detailed_prompt_parts_for_genai,
            is_initial_analysis=False,
            user_prompt_for_history=user_prompt_for_history
        )

        if ai_response_text and not ai_response_text.startswith("AI Error:"):
            self._populate_comparison_treeview(ai_response_text)
        else:
            if not ai_response_text:
                 self.update_conversation_history("System: Failed to get detailed comparison from AI (no response).", role="error")

        if self.model and hasattr(self, 'start_comparison_button'):
            self.start_comparison_button.config(state=tk.NORMAL)


    def _parse_markdown_table(self, markdown_text: str) -> list:
        table_data = []
        lines = markdown_text.split('\n')
        header_pattern = re.compile(r"^\s*\|([^|]+)\|([^|]+)\|([^|]+)\|([^|]+)?\|?\s*$")
        separator_pattern = re.compile(r"^\s*\|?[-:|\s]+\|?\s*$")
        in_table_block = False
        for line in lines:
            line = line.strip()
            if not line:
                if in_table_block: in_table_block = False
                continue
            if line.startswith('|') and line.endswith('|'):
                in_table_block = True
                if separator_pattern.match(line): continue
                match = header_pattern.match(line)
                if match:
                    parts = [p.strip() for p in match.groups()]
                    if len(parts) >= 3:
                        table_data.append({
                            "parameter": parts[0], "component1": parts[1],
                            "component2": parts[2], "notes": parts[3] if len(parts) > 3 and parts[3] is not None else ""
                        })
            elif in_table_block: in_table_block = False
        print(f"DEBUG: Parsed table data: {table_data}"); return table_data

    def _populate_comparison_treeview(self, ai_response_text: str):
        if hasattr(self, 'comparison_treeview'):
            for item in self.comparison_treeview.get_children(): self.comparison_treeview.delete(item)
        else: self.update_conversation_history("System: Treeview not found.", role="error"); return
        parsed_data = self._parse_markdown_table(ai_response_text)
        if not parsed_data: self.update_conversation_history("System: No table data parsed for Treeview.", role="system"); return
        pn1 = self.mfg_pn_var_1.get() or (os.path.basename(self.spec_sheet_1_path) if self.spec_sheet_1_path else "Comp 1")
        pn2 = self.mfg_pn_var_2.get() or (os.path.basename(self.spec_sheet_2_path) if self.spec_sheet_2_path else "Comp 2")
        self.comparison_treeview.heading("component1", text=f"{pn1[:25]}{'...' if len(pn1)>25 else ''}")
        self.comparison_treeview.heading("component2", text=f"{pn2[:25]}{'...' if len(pn2)>25 else ''}")
        for row in parsed_data:
            self.comparison_treeview.insert("", tk.END, values=(
                row.get("parameter", ""), row.get("component1", ""),
                row.get("component2", ""), row.get("notes", "")))
        self.update_conversation_history("System: Detailed comparison table populated.", role="system")

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
        lines = text_response.split('\n')
        formatted_lines = []
        separator_pattern = re.compile(r"^\s*\|?[-:|\s]+\|?\s*$")
        for line in lines:
            if '|' in line:
                processed_line = line
                if separator_pattern.match(line): processed_line = processed_line.replace("-", "—")
                processed_line = re.sub(r'\s*\|\s*', '  |  ', processed_line).strip()
                formatted_lines.append(processed_line)
            else: formatted_lines.append(line)
        return "\n".join(formatted_lines)

    def update_conversation_history(self, message, role="system"):
        if hasattr(self, 'conversation_history') and self.conversation_history:
            self.conversation_history.config(state=tk.NORMAL)
            tag_to_apply = {"user": "user_message", "ai": "ai_message", "error": "error_message"}.get(role, "system_message")
            display_message = self._format_ai_response(message) if role == "ai" else message
            self.conversation_history.insert(tk.END, display_message + "\n", tag_to_apply)
            self.conversation_history.see(tk.END)
            self.conversation_history.config(state=tk.DISABLED)
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
            if not api_key: print("DEBUG: GOOGLE_API_KEY not found for _configure_ai."); self.api_key_configured = False
            else: genai.configure(api_key=api_key); self.api_key_configured = True
        except Exception as e: self.update_conversation_history(f"System: Error configuring AI SDK: {e}", role="error"); self.api_key_configured = False
        finally: self._update_ui_for_ai_status(api_key_configured=self.api_key_configured, model_initialized=(self.model is not None))

    def _on_model_selected(self, event=None):
        if event: print(f"DEBUG: _on_model_selected. Event: {event.type}, Widget: {event.widget}")
        else: print(f"DEBUG: _on_model_selected programmatically.")
        selected_model_name = self.model_combobox.get(); print(f"DEBUG: Combobox get(): '{selected_model_name}'")
        if selected_model_name == self.placeholder_text:
            self.update_conversation_history("System: Select a valid AI model.", role="system"); self._update_ui_for_ai_status(model_initialized=False); return
        previous_model_name = self.model.model_name if self.model else None
        is_diff_model = self.model and self.model.model_name != selected_model_name
        is_first_select_with_history = not self.model and self.conversation_log and any(not log.startswith("System: Welcome!") for log in self.conversation_log)
        if is_diff_model or is_first_select_with_history:
            log_msg = f"System: Changing model";
            if previous_model_name: log_msg += f" from {previous_model_name}"
            log_msg += f" to {selected_model_name}. Clearing context."
            self.update_conversation_history(log_msg, role="system"); self.clear_all(clear_files=False)
            self.update_conversation_history(f"System: AI Model selected: {selected_model_name}", role="system")
        else: self.update_conversation_history(f"System: AI Model selected: {selected_model_name}", role="system")
        self.model_initializing = True
        model_init_ok = self._initialize_model(selected_model_name)
        self.model_initializing = False
        if model_init_ok:
            if self.spec_sheet_1_path and self.spec_sheet_2_path:
                self.update_conversation_history(f"System: Starting analysis with {selected_model_name}...", role="system")
                self.check_and_process_spec_sheets()
            else: self.update_conversation_history("System: Model initialized. Load both spec sheets.", role="system")
        else: self.update_conversation_history(f"System: Failed to init {selected_model_name}. Check logs.", role="error")

    def _initialize_model(self, model_name=None):
        source_log = "explicitly";
        if model_name is None: model_name = self.model_combobox.get(); source_log = f"from Combobox: '{model_name}'"
        if model_name == self.placeholder_text:
            self.update_conversation_history("System: Select valid model.", role="system"); self.model=None; self.chat_session=None; self._update_ui_for_ai_status(model_initialized=False); return False
        self.update_conversation_history(f"System: _initialize_model ({source_log}). Target: '{model_name}'", role="system")
        if not model_name or model_name not in self.model_options_list:
            self.update_conversation_history(f"System: Invalid model ('{model_name}'). Aborted.", role="error"); self.model=None; self.chat_session=None; self._update_ui_for_ai_status(model_initialized=False); return False
        if self.model and self.model.model_name == model_name and self.api_key_configured:
            self.update_conversation_history(f"System: Model '{model_name}' already active.", role="system"); self._update_ui_for_ai_status(model_initialized=True); return True
        if not self.api_key_configured:
            self.update_conversation_history("System: Cannot init model - API key not set.", role="error"); self.model=None; self.chat_session=None; self._update_ui_for_ai_status(model_initialized=False); return False
        self.update_conversation_history(f"System: Initializing model: {model_name}...", role="system")
        try:
            if not any("Generative AI configured successfully." in log for log in self.conversation_log): self.update_conversation_history("System: Generative AI configured successfully.", role="system")
            self.model = genai.GenerativeModel(model_name); self.chat_session = None
            self.update_conversation_history(f"System: Successfully initialized model: {model_name}", role="system"); self._update_ui_for_ai_status(model_initialized=True); return True
        except Exception as e: self.model=None; self.chat_session=None; self.update_conversation_history(f"System: Error initializing model {model_name}: {e}", role="error"); self._update_ui_for_ai_status(model_initialized=False); return False

    def get_selected_model_name(self): return self.model_var.get()
    def _add_to_ai_history(self,role:str,text_content:str): self.ai_history.append({'role':role,'parts':[text_content]}); print(f"DEBUG: AI history add: {role}, '{text_content[:50]}...'")
    def _convert_log_to_gemini_history(self): return [e for e in self.ai_history if e['role'] in ('user','model')]

    def send_user_query(self):
        if not self.model or not self.api_key_configured:
            if not self._initialize_model(): return
        if not self.model: self.update_conversation_history("System: AI Model N/A.", role="error"); return

        user_text = self.user_input_entry.get().strip()
        # self.update_conversation_history is called after constructing prompt_parts_for_ai
        self.user_input_entry.delete(0, tk.END)

        prompt_parts_for_ai = []
        image_sent_this_turn = False
        log_message_for_user_turn = user_text

        if self.pending_user_image_pil:
            if user_text: # Text and image
                prompt_parts_for_ai.append(user_text)
                prompt_parts_for_ai.append(self.pending_user_image_pil)
            else: # Image only
                prompt_parts_for_ai.append(self.pending_user_image_pil)

            image_filename = os.path.basename(self.pending_user_image_path)
            self.update_conversation_history(f"User: {user_text} [Image: {image_filename}]", role="user")
            log_message_for_user_turn = f"{user_text} [Image: {image_filename}]"
            image_sent_this_turn = True
        elif user_text: # Text only
            prompt_parts_for_ai.append(user_text)
            self.update_conversation_history(f"User: {user_text}", role="user")
        else:
            self.update_conversation_history("System: Cannot send empty message.", role="system"); return

        active_model_name = self.model.model_name
        try:
            self.send_button.config(state=tk.DISABLED); self.user_input_entry.config(state=tk.DISABLED)
            if not self.chat_session:
                self.update_conversation_history(f"System: Starting new chat with {active_model_name}...", role="system")
                try: self.chat_session = self.model.start_chat(history=self._convert_log_to_gemini_history())
                except Exception as e: self.update_conversation_history(f"System: Error starting chat: {e}", role="error"); self._update_ui_for_ai_status(); return

            self._add_to_ai_history('user', log_message_for_user_turn)

            self.update_conversation_history(f"System: Sending to AI ({active_model_name})...", role="system")

            final_prompt_parts_for_sending = list(prompt_parts_for_ai) # Use a copy
            if self.translate_to_chinese_var.get():
                translation_instruction = " Please provide your entire response in Chinese."
                appended_to_text = False
                for i in range(len(final_prompt_parts_for_sending) - 1, -1, -1):
                    if isinstance(final_prompt_parts_for_sending[i], str):
                        final_prompt_parts_for_sending[i] += translation_instruction
                        appended_to_text = True; break
                if not appended_to_text: final_prompt_parts_for_sending.append(translation_instruction)

            response = self.chat_session.send_message(final_prompt_parts_for_sending)
            self._add_to_ai_history('model', response.text)
            self.update_conversation_history(f"AI ({active_model_name}): {response.text}", role="ai")

            if image_sent_this_turn:
                self.pending_user_image_path = None; self.pending_user_image_pil = None
        except Exception as e:
            err_msg=f"System: Error with AI ({active_model_name}): {e}"; self.update_conversation_history(err_msg,role="error"); print(f"DEBUG: {err_msg}")
            if isinstance(e,(google_exceptions.PermissionDenied,google_exceptions.Unauthenticated)): self.api_key_configured=False
            if isinstance(e,(google_exceptions.InvalidArgument,ValueError,BlockedPromptException,StopCandidateException,google_exceptions.NotFound,google_exceptions.PermissionDenied)):
                self.update_conversation_history(f"System: Resetting model ({active_model_name}) due to error.",role="system"); self.model=None; self.chat_session=None; self.ai_history=[]
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
            self.spec_sheet_1_path=None; self.spec_sheet_1_text=None; self.spec_sheet_1_image_paths=[]
            self.spec_sheet_2_path=None; self.spec_sheet_2_text=None; self.spec_sheet_2_image_paths=[]
            if hasattr(self,'spec_sheet_1_label'): self.spec_sheet_1_label.config(text="File 1: None")
            if hasattr(self,'spec_sheet_2_label'): self.spec_sheet_2_label.config(text="File 2: None")
            if hasattr(self,'mfg_pn_var_1'): self.mfg_pn_var_1.set("")
            if hasattr(self,'mfg_pn_var_2'): self.mfg_pn_var_2.set("")
            if hasattr(self,'model_combobox'): self.model_combobox.set(self.placeholder_text); self.model_combobox.state(["disabled"])
            if os.path.exists(self.temp_image_dir):
                try: shutil.rmtree(self.temp_image_dir); print(f"DEBUG: Deleted temp dir: {self.temp_image_dir}")
                except OSError as e: print(f"Error deleting temp dir {self.temp_image_dir}: {e}")
            self._create_temp_image_dir()

        self.pending_user_image_path = None
        self.pending_user_image_pil = None

        if hasattr(self,'conversation_history'): self.conversation_history.config(state=tk.NORMAL); self.conversation_history.delete(1.0, tk.END)
        if hasattr(self,'comparison_treeview'):
            for item in self.comparison_treeview.get_children(): self.comparison_treeview.delete(item)
        self.conversation_log=[]; self.ai_history=[]
        if clear_files: self.update_conversation_history("System: Welcome! Load PDFs to start.",role="system")
        if hasattr(self,'user_input_entry'): self.user_input_entry.delete(0,tk.END)
        self.model=None; self.chat_session=None
        self._configure_ai()
        if not clear_files and self.spec_sheet_1_path and self.spec_sheet_2_path:
            if hasattr(self,'model_combobox'): self.model_combobox.config(state='readonly')
            self.update_conversation_history("System: AI context cleared. Files remain. Select model.",role="system")
        elif clear_files:
             if hasattr(self,'model_combobox'): self.model_combobox.state(['disabled'])
        if hasattr(self, 'start_comparison_button'): self.start_comparison_button.config(state=tk.DISABLED)
        self._update_ui_for_ai_status(api_key_configured=self.api_key_configured,model_initialized=False)
        print("DEBUG: Clear All finished.")

    def extract_text_from_pdf(self, filepath):
        if not filepath or not os.path.exists(filepath): self.update_conversation_history(f"System: PDF not found: {os.path.basename(filepath or 'Unknown')}", role="error"); return ""
        try:
            self.update_conversation_history(f"System: Extracting text from {os.path.basename(filepath)}...", role="system")
            with fitz.open(filepath) as doc: text = "".join(page.get_text() for page in doc)
            self.update_conversation_history(f"System: Text extraction OK: {os.path.basename(filepath)}.", role="system"); return text
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
        if not self.api_key_configured: self.update_conversation_history("System: API Key not configured.", role="error"); return
        if not self.model: self.update_conversation_history("System: AI Model not selected. Please select a model.", role="system"); return
        self.update_conversation_history("System: Files and model active. Clearing old results...", role="system")
        self.conversation_log = []; self.ai_history = []
        if hasattr(self, 'conversation_history'): self.conversation_history.config(state=tk.NORMAL); self.conversation_history.delete(1.0, tk.END)
        if hasattr(self, 'comparison_treeview'):
            for item in self.comparison_treeview.get_children(): self.comparison_treeview.delete(item)
        if self.api_key_configured: self.update_conversation_history("System: AI Configured.", role="system")
        if self.model: self.update_conversation_history(f"System: Model '{self.model.model_name}' active.", role="system")
        self.process_spec_sheets()

    def process_spec_sheets(self):
        if not self.model or not self.api_key_configured or not self.spec_sheet_1_path or not self.spec_sheet_2_path:
            self.update_conversation_history("System: Pre-reqs not met (files, API key, model).", role="error"); return
        self.update_conversation_history("System: Starting initial analysis...", role="system")
        self.spec_sheet_1_text = self.extract_text_from_pdf(self.spec_sheet_1_path)
        if not self.spec_sheet_1_text: self.update_conversation_history(f"System: Halting. Text extract fail: {os.path.basename(self.spec_sheet_1_path)}.", role="error"); return
        s1_f = os.path.join(self.temp_image_dir, f"{os.path.splitext(os.path.basename(self.spec_sheet_1_path))[0]}_imgs_{len(os.listdir(self.temp_image_dir))}")
        self.spec_sheet_1_image_paths = self.extract_images_from_pdf(self.spec_sheet_1_path, s1_f)
        self.spec_sheet_2_text = self.extract_text_from_pdf(self.spec_sheet_2_path)
        if not self.spec_sheet_2_text: self.update_conversation_history(f"System: Halting. Text extract fail: {os.path.basename(self.spec_sheet_2_path)}.", role="error"); return
        s2_f = os.path.join(self.temp_image_dir, f"{os.path.splitext(os.path.basename(self.spec_sheet_2_path))[0]}_imgs_{len(os.listdir(self.temp_image_dir))}")
        self.spec_sheet_2_image_paths = self.extract_images_from_pdf(self.spec_sheet_2_path, s2_f)

        initial_analysis_prompt_text = (
            "You are an expert electronics component analyst. Analyze the following two component specification sheets.\n\n"
            "**Instructions for AI:**\n"
            "1. For Component 1 (described first), identify its specific component type.\n"
            "2. For Component 2 (described second), identify its specific component type.\n"
            "3. Assess if Component 1 and Component 2 are functionally similar (e.g., both are dual N-channel MOSFETs, or one is an LDO regulator and the other a switching regulator, or one is a TVS diode and the other a Zener diode). Your assessment should be based on their primary function.\n"
            "4. For Component 1, find and extract the first complete Manufacturer Part Number (MFG P/N) listed in its 'Order Information' or equivalent section. If multiple are listed, provide only the first one. If none is explicitly found, state 'Not Found'.\n"
            "5. For Component 2, find and extract the first complete Manufacturer Part Number (MFG P/N) listed in its 'Order Information' or equivalent section. If multiple are listed, provide only the first one. If none is explicitly found, state 'Not Found'.\n\n"
            "**Output Format:**\n"
            "Please provide your response *only* in the following structured format, using these exact labels:\n"
            "Component1_Type: [Type for component 1]\n"
            "Component2_Type: [Type for component 2]\n"
            "Functionally_Similar: [Yes/No, brief explanation]\n"
            "MFG_PN1: [MFG P/N for component 1 or 'Not Found']\n"
            "MFG_PN2: [MFG P/N for component 2 or 'Not Found']\n\n"
            "**Component 1 Data:**\n"
            f"Text Content:\n{self.spec_sheet_1_text}\n\n"
            "**Component 2 Data:**\n"
            f"Text Content:\n{self.spec_sheet_2_text}\n"
        )

        prompt_parts_for_genai = [initial_analysis_prompt_text]
        # Add images for component 1
        for img_path in self.spec_sheet_1_image_paths:
            try: prompt_parts_for_genai.append(Image.open(img_path))
            except Exception as e: self.update_conversation_history(f"System: Error loading image {img_path} for Comp 1. Skip. Err: {e}", role="error")
        # Add images for component 2
        prompt_parts_for_genai.append("\n--- End of Component 1 Images, Start of Component 2 Images (if any) ---") # Separator for clarity if needed
        for img_path in self.spec_sheet_2_image_paths:
            try: prompt_parts_for_genai.append(Image.open(img_path))
            except Exception as e: self.update_conversation_history(f"System: Error loading image {img_path} for Comp 2. Skip. Err: {e}", role="error")

        user_prompt_for_history_log = "User: Initial component type identification and MFG P/N extraction for spec sheets."
        self.send_to_ai(prompt_parts_for_genai, is_initial_analysis=True, user_prompt_for_history=user_prompt_for_history_log)


    def send_to_ai(self, prompt_parts, is_initial_analysis=False, user_prompt_for_history=None):
        if not self.model: self.update_conversation_history("System: AI model N/A.", role="error"); return None
        active_model_name = self.model.model_name
        raw_ai_response_text = ""

        final_prompt_parts = list(prompt_parts) # Work with a copy
        if self.translate_to_chinese_var.get():
            translation_instruction = " Please provide your entire response in Chinese."
            appended_to_text = False
            for i in range(len(final_prompt_parts) - 1, -1, -1):
                if isinstance(final_prompt_parts[i], str):
                    final_prompt_parts[i] += translation_instruction
                    appended_to_text = True; break
            if not appended_to_text: final_prompt_parts.append(translation_instruction)

        try:
            self.send_button.config(state=tk.DISABLED); self.user_input_entry.config(state=tk.DISABLED)
            if hasattr(self, 'start_comparison_button'): self.start_comparison_button.config(state=tk.DISABLED)
            self.update_conversation_history(f"System: Sending to AI ({active_model_name})... May take time.", role="system")

            if is_initial_analysis and user_prompt_for_history:
                 self._add_to_ai_history('user', user_prompt_for_history)

            response = self.model.generate_content(final_prompt_parts, request_options={'timeout': 600})

            if response.prompt_feedback and response.prompt_feedback.block_reason:
                raw_ai_response_text = f"AI Error - Prompt was blocked. Reason: {response.prompt_feedback.block_reason}"
                self.update_conversation_history(f"System: {raw_ai_response_text}", role="error")
            elif not response.candidates or not hasattr(response, 'text') or not response.text:
                raw_ai_response_text = "AI response empty/no content."
                self.update_conversation_history(f"System: AI ({active_model_name}): {raw_ai_response_text}", role="system")
            else:
                raw_ai_response_text = response.text
                self.update_conversation_history(f"AI ({active_model_name}): {raw_ai_response_text}", role="ai") # Display formatted

            self._add_to_ai_history('model', raw_ai_response_text) # Log model's raw response or error

            if is_initial_analysis:
                self.chat_session = None
                # Parse the response and update UI elements
                parsed_info = self._parse_initial_analysis_response(raw_ai_response_text)

                self.update_conversation_history(f"System: Initial Analysis Parsed Data:", role="system")
                self.update_conversation_history(f"  Component 1 Type: {parsed_info['component1_type']}", role="system")
                self.update_conversation_history(f"  Component 2 Type: {parsed_info['component2_type']}", role="system")
                self.update_conversation_history(f"  Functionally Similar: {parsed_info['functionally_similar']}", role="system")

                if hasattr(self, 'mfg_pn_var_1'):
                    self.mfg_pn_var_1.set(parsed_info['mfg_pn1'] if parsed_info['mfg_pn1'] != "Not Found" else "")
                    self.update_conversation_history(f"  MFG P/N 1 set to: {self.mfg_pn_var_1.get() or 'Not Found'}", role="system")
                if hasattr(self, 'mfg_pn_var_2'):
                    self.mfg_pn_var_2.set(parsed_info['mfg_pn2'] if parsed_info['mfg_pn2'] != "Not Found" else "")
                    self.update_conversation_history(f"  MFG P/N 2 set to: {self.mfg_pn_var_2.get() or 'Not Found'}", role="system")

                if hasattr(self, 'start_comparison_button'):
                    if parsed_info["is_similar_flag"] and not ("AI Error" in raw_ai_response_text or "empty/no content" in raw_ai_response_text) :
                        self.start_comparison_button.config(state=tk.NORMAL)
                        self.update_conversation_history("System: Components appear functionally similar. 'Start Detailed Comparison' enabled.", role="system")
                    else:
                        self.start_comparison_button.config(state=tk.DISABLED)
                        self.update_conversation_history("System: Components may not be functionally similar or analysis incomplete. Detailed comparison not enabled.", role="system")
            return raw_ai_response_text
        except Exception as e:
            err_msg = f"System: Error with AI ({active_model_name}): {e}"
            self.update_conversation_history(err_msg, role="error"); print(f"DEBUG: {err_msg}")
            self._add_to_ai_history('model', f"Error: {e}")
            if hasattr(self, 'start_comparison_button'): self.start_comparison_button.config(state=tk.DISABLED)
            if isinstance(e, (google_exceptions.PermissionDenied,google_exceptions.Unauthenticated)): self.api_key_configured=False
            if isinstance(e, (google_exceptions.InvalidArgument, ValueError, BlockedPromptException, StopCandidateException, google_exceptions.NotFound, google_exceptions.PermissionDenied)):
                self.update_conversation_history(f"System: Resetting model ({active_model_name}) due to error.", role="system")
                self.model = None; self.chat_session = None; self.ai_history = []
            return f"AI Error: {e}"
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

[end of main.py]
