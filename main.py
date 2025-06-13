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

        self.model = None # Stores the initialized GenerativeModel instance
        self.chat_session = None # Stores the current chat session with the AI
        self.conversation_log = [] # Internal list to store conversation strings for download
        self.api_key_configured = False # Flag to track API key status

        self.temp_image_dir = "temp_images" # Folder for storing extracted images temporarily
        if not os.path.exists(self.temp_image_dir):
            try:
                os.makedirs(self.temp_image_dir)
            except OSError as e:
                # This is an early critical error. We'll try to show it in the UI if it's up.
                # If UI isn't up yet, this print will go to console.
                error_message = f"Critical Error: Cannot create temporary directory {self.temp_image_dir}: {e}"
                print(error_message)
                if hasattr(self, 'conversation_history'): # Check if UI is partially initialized
                    self.update_conversation_history(error_message)
                # Consider disabling file loading buttons if this fails, as image extraction will fail.

        self._setup_ui(root)
        self._configure_ai() # Attempt to configure AI and update UI state accordingly

    def _setup_ui(self, root):
        """Configures the Tkinter UI elements and layout."""
        # --- UI Elements ---
        # File loading section
        self.spec_sheet_1_label = ttk.Label(root, text="File 1: None")
        self.spec_sheet_1_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.load_spec_sheet_1_button = ttk.Button(root, text="Load Spec Sheet 1", command=self.load_spec_sheet_1)
        self.load_spec_sheet_1_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.spec_sheet_2_label = ttk.Label(root, text="File 2: None")
        self.spec_sheet_2_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.load_spec_sheet_2_button = ttk.Button(root, text="Load Spec Sheet 2", command=self.load_spec_sheet_2)
        self.load_spec_sheet_2_button.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # AI Model selection
        self.model_label = ttk.Label(root, text="Select AI Model:")
        self.model_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.model_var = tk.StringVar()
        self.model_combobox = ttk.Combobox(root, textvariable=self.model_var, state="readonly") # Readonly state
        self.model_combobox['values'] = (
            "gemini-1.5-flash-latest",
            "gemini-1.5-pro-latest",
            "gemini-1.0-pro",
            # Gemma models might require different setup or might not support all features.
            # Disabling them for now to simplify error handling related to model capabilities.
            # "gemma-7b",
            # "gemma-2b",
        )
        self.model_combobox.current(0)
        self.model_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        # Conversation history
        self.history_label = ttk.Label(root, text="Conversation History:")
        self.history_label.grid(row=3, column=0, columnspan=2, padx=10, pady=(10,0), sticky="w") # Added more padding
        self.conversation_history = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=15, width=80) # Increased width
        self.conversation_history.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")
        self.conversation_history.config(state=tk.DISABLED)

        # User input
        self.input_label = ttk.Label(root, text="Your Message:")
        self.input_label.grid(row=5, column=0, columnspan=2, padx=10, pady=(10,0), sticky="w") # Added more padding
        self.user_input_entry = ttk.Entry(root, width=70) # Increased width
        self.user_input_entry.grid(row=6, column=0, padx=10, pady=5, sticky="ew")
        self.send_button = ttk.Button(root, text="Send", command=self.send_user_query)
        self.send_button.grid(row=6, column=1, padx=(0,10), pady=5, sticky="ew") # Adjusted padx

        # Action buttons
        self.download_history_button = ttk.Button(root, text="Download History", command=self.download_history)
        self.download_history_button.grid(row=7, column=0, padx=10, pady=10, sticky="ew")
        self.clear_all_button = ttk.Button(root, text="Clear All", command=self.clear_all)
        self.clear_all_button.grid(row=7, column=1, padx=(0,10), pady=10, sticky="ew") # Adjusted padx

        # Configure column weights for responsiveness
        root.grid_columnconfigure(0, weight=3) # Give more weight to the text area column
        root.grid_columnconfigure(1, weight=1)
        root.grid_rowconfigure(4, weight=1) # Conversation history expands vertically

    def load_spec_sheet_1(self):
        """Handles the 'Load Spec Sheet 1' button action.
        Opens a file dialog for PDF selection, updates the UI label,
        and triggers processing if both sheets are loaded.
        """
        filepath = filedialog.askopenfilename(
            title="Select Spec Sheet 1 (PDF)",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filepath: # File selected
            self.spec_sheet_1_path = filepath
            self.spec_sheet_1_label.config(text=f"File 1: {os.path.basename(filepath)}")
            print(f"Spec sheet 1 loaded: {self.spec_sheet_1_path}")
            self.check_and_process_spec_sheets()
        else: # Dialog cancelled
            # Keep existing path if any, or ensure label shows "None" if path was None
            current_filename = os.path.basename(self.spec_sheet_1_path) if self.spec_sheet_1_path else "None"
            self.spec_sheet_1_label.config(text=f"File 1: {current_filename}")
            # No error message needed for cancellation.
            print("Spec sheet 1 loading cancelled or no file selected.")


    def load_spec_sheet_2(self):
        """Handles the 'Load Spec Sheet 2' button action.
        Similar to load_spec_sheet_1.
        """
        filepath = filedialog.askopenfilename(
            title="Select Spec Sheet 2 (PDF)",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filepath: # File selected
            self.spec_sheet_2_path = filepath
            self.spec_sheet_2_label.config(text=f"File 2: {os.path.basename(filepath)}")
            print(f"Spec sheet 2 loaded: {self.spec_sheet_2_path}")
            self.check_and_process_spec_sheets()
        else: # Dialog cancelled
            current_filename = os.path.basename(self.spec_sheet_2_path) if self.spec_sheet_2_path else "None"
            self.spec_sheet_2_label.config(text=f"File 2: {current_filename}")
            print("Spec sheet 2 loading cancelled or no file selected.")

    def update_conversation_history(self, message, is_internal_log_message=False):
        """
        Updates the ScrolledText widget with a new message and appends to the internal log.

        Args:
            message (str): The message to display and log.
            is_internal_log_message (bool): If True, this message might be treated differently
                                           for logging vs. display. Currently, all messages
                                           are displayed and logged.
        """
        # Ensure UI component exists before trying to update it (robustness)
        if hasattr(self, 'conversation_history') and self.conversation_history:
            self.conversation_history.config(state=tk.NORMAL)
            self.conversation_history.insert(tk.END, message + "\n")
            self.conversation_history.see(tk.END) # Auto-scroll to the latest message
            self.conversation_history.config(state=tk.DISABLED)

        # For now, all messages shown in UI are added to the downloadable log.
        self.conversation_log.append(message)

    def _update_ui_for_ai_status(self):
        """Enables or disables UI elements based on AI configuration status."""
        # Ensure UI components exist
        if not hasattr(self, 'send_button'):
            return # UI not fully initialized yet

        if self.api_key_configured:
            self.send_button.config(state=tk.NORMAL)
            self.user_input_entry.config(state=tk.NORMAL)
            # Model combobox should be interactable if API key is there
            self.model_combobox.config(state="readonly")
        else:
            self.send_button.config(state=tk.DISABLED)
            self.user_input_entry.config(state=tk.DISABLED)
            # Allow model selection even if API key is missing,
            # error will be handled upon sending.
            self.model_combobox.config(state="readonly")


    def _configure_ai(self):
        """
        Configures the Generative AI SDK with the API key from environment variables.
        Updates UI elements based on success or failure.
        """
        try:
            # This check should ideally happen before UI that depends on it is active.
            # However, __init__ order means update_conversation_history might be called early.
            api_key = os.environ.get("GOOGLE_API_KEY")
            if not api_key:
                self.update_conversation_history("System: Error - GOOGLE_API_KEY environment variable not set. AI features disabled.")
                self.api_key_configured = False
            else:
                genai.configure(api_key=api_key)
                self.update_conversation_history("System: Generative AI configured successfully.")
                self.api_key_configured = True
        except Exception as e: # Catch any exception during genai.configure
            self.update_conversation_history(f"System: Error configuring Generative AI SDK - {e}")
            self.api_key_configured = False
        finally:
            # Update UI state regardless of success/failure of AI configuration
            self._update_ui_for_ai_status()


    def get_selected_model_name(self):
        """Returns the currently selected AI model name from the Combobox."""
        return self.model_var.get()

    def send_user_query(self):
        """
        Handles sending user's text query to the AI.
        Initializes model and chat session if needed.
        Updates conversation history with user message and AI response.
        Manages UI state of the send button during AI interaction.
        """
        if not self.api_key_configured:
            self.update_conversation_history("System: AI is not configured. Cannot send message. Please set GOOGLE_API_KEY.")
            return

        user_text = self.user_input_entry.get().strip()
        if not user_text:
            return # Do nothing if input is empty

        self.update_conversation_history(f"User: {user_text}")
        self.user_input_entry.delete(0, tk.END) # Clear input field

        selected_model_name = self.get_selected_model_name()
        if not selected_model_name:
            self.update_conversation_history("System: Please select an AI model first.")
            return

        try:
            self.send_button.config(state=tk.DISABLED) # Disable send button during processing
            self.user_input_entry.config(state=tk.DISABLED) # Disable input during processing

            # Initialize or change model if necessary
            if not self.model or self.model.model_name != selected_model_name:
                self.update_conversation_history(f"System: Initializing model {selected_model_name}...")
                self.model = genai.GenerativeModel(selected_model_name)
                self.chat_session = self.model.start_chat(history=[]) # New model = new session
                self.update_conversation_history(f"System: Started new chat session with {selected_model_name}.")
            elif not self.chat_session: # Model exists, but no active chat session
                 self.chat_session = self.model.start_chat(history=[])
                 self.update_conversation_history(f"System: Resumed chat session with {selected_model_name}.")

            self.update_conversation_history(f"System: Sending to AI ({selected_model_name})...")
            response = self.chat_session.send_message(user_text)
            self.update_conversation_history(f"AI ({selected_model_name}): {response.text}")

        except google_exceptions.InvalidArgument as e:
            self.update_conversation_history(f"System: AI Error - Invalid argument (check prompt): {e}")
        except google_exceptions.PermissionDenied as e:
            self.update_conversation_history(f"System: AI Error - Permission denied (check API key/permissions): {e}")
            self.api_key_configured = False # Potentially reset API key status
            self._update_ui_for_ai_status()
        except google_exceptions.ServiceUnavailable as e:
            self.update_conversation_history(f"System: AI Error - Service unavailable. Please try again later: {e}")
        except google_exceptions.GoogleAPIError as e: # Catch other specific Google API errors
            self.update_conversation_history(f"System: Google API Error during AI interaction: {e}")
        except BlockedPromptException as e:
            self.update_conversation_history(f"System: AI Error - Your prompt was blocked. Details: {e}")
        except StopCandidateException as e:
            self.update_conversation_history(f"System: AI Error - Generation stopped unexpectedly. Details: {e}")
        except ValueError as e: # Catches model initialization errors or other value errors
             self.update_conversation_history(f"System: Error during AI setup (e.g. invalid model name, API issue): {e}")
             self.model = None # Reset model state as it might be corrupted
             self.chat_session = None
        except Exception as e: # General fallback for other GenAI errors or unexpected issues
            self.update_conversation_history(f"System: An unexpected error occurred during AI interaction: {e}")
        finally:
            # Re-enable send button and input field if API key is still considered configured
            if self.api_key_configured:
                self.send_button.config(state=tk.NORMAL)
                self.user_input_entry.config(state=tk.NORMAL)
            else: # If API key failed, ensure they remain disabled
                self.send_button.config(state=tk.DISABLED)
                self.user_input_entry.config(state=tk.DISABLED)


    def download_history(self):
        """
        Handles the 'Download History' button action.
        Prompts user for save location and saves conversation log as a .docx file.
        Includes basic error handling for file operations.
        """
        if not self.conversation_log:
            self.update_conversation_history("System: Conversation history is empty. Nothing to download.")
            return

        try:
            filepath = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")],
                title="Save Conversation History"
            )
        except Exception as e: # Handle potential error with filedialog itself
            self.update_conversation_history(f"System: Error opening save dialog: {e}")
            return

        if not filepath: # User cancelled the dialog
            self.update_conversation_history("System: Download cancelled by user.")
            return

        try:
            doc = docx.Document()
            doc.add_heading("Component Comparator AI - Chat History", level=1)

            # Add metadata
            if self.spec_sheet_1_path:
                doc.add_paragraph(f"Spec Sheet 1: {os.path.basename(self.spec_sheet_1_path)}")
            if self.spec_sheet_2_path:
                doc.add_paragraph(f"Spec Sheet 2: {os.path.basename(self.spec_sheet_2_path)}")

            # Try to get model name safely
            model_name_to_log = "N/A"
            if self.model and hasattr(self.model, 'model_name'):
                model_name_to_log = self.model.model_name
            doc.add_paragraph(f"AI Model Used (last for query/analysis): {model_name_to_log}")
            doc.add_paragraph("-" * 20) # Visual separator

            for entry in self.conversation_log:
                doc.add_paragraph(entry)

            doc.save(filepath)
            self.update_conversation_history(f"System: Conversation history downloaded to {filepath}")
        except IOError as e: # Specific error for file I/O issues
            self.update_conversation_history(f"System: Error saving file (IOError): {e}")
            print(f"Error saving .docx (IOError): {e}") # Also print to console for debugging
        except PermissionError as e: # Specific error for permission issues
            self.update_conversation_history(f"System: Error saving file - Permission denied: {e}")
            print(f"Error saving .docx (PermissionError): {e}")
        except Exception as e: # Catch other potential errors from docx library or unexpected issues
            self.update_conversation_history(f"System: Error downloading history: {e}")
            print(f"Error saving .docx: {e}")
        # No longer printing "Download history button clicked" here, UI message is enough.

    def clear_all(self):
        """
        Resets the application state: clears file paths, extracted data,
        AI model/chat session, conversation UI, and internal log.
        Also cleans up the temporary image directory.
        """
        # Reset file paths and related UI labels
        self.spec_sheet_1_label.config(text="File 1: None")
        self.spec_sheet_1_path = None
        self.spec_sheet_1_text = None
        self.spec_sheet_1_image_paths = []

        self.spec_sheet_2_label.config(text="File 2: None")
        self.spec_sheet_2_path = None
        self.spec_sheet_2_text = None
        self.spec_sheet_2_image_paths = []

        # Reset AI state
        self.model_combobox.current(0) # Reset to default model selection
        self.model = None
        self.chat_session = None
        self.conversation_log = [] # Clear the internal conversation log

        # Clear UI conversation history display
        if hasattr(self, 'conversation_history') and self.conversation_history:
            self.conversation_history.config(state=tk.NORMAL)
            self.conversation_history.delete(1.0, tk.END)
            # Don't re-disable here, _configure_ai will handle it via update_conversation_history

        # Re-run AI configuration (checks API key, updates UI history and button states)
        self._configure_ai()

        # Clear user input field
        if hasattr(self, 'user_input_entry') and self.user_input_entry:
            self.user_input_entry.delete(0, tk.END)

        # Clear and recreate temporary image directory
        try:
            if os.path.exists(self.temp_image_dir):
                shutil.rmtree(self.temp_image_dir)
            os.makedirs(self.temp_image_dir) # Recreate for next session
        except OSError as e:
            # If UI is available, show error there, otherwise print.
            err_msg = f"System: Error cleaning temporary image directory: {e}"
            if hasattr(self, 'conversation_history') and self.conversation_history:
                self.update_conversation_history(err_msg)
            else:
                print(err_msg)

        print("Clear All: Application state reset.") # Console message for debugging


    def extract_text_from_pdf(self, filepath):
        """
        Extracts all text from a given PDF file using PyMuPDF (fitz).

        Args:
            filepath (str): The path to the PDF file.

        Returns:
            str: The concatenated text from all pages, or an empty string on failure.
        """
        if not filepath or not os.path.exists(filepath):
            self.update_conversation_history(f"System: PDF file not found for text extraction: {os.path.basename(filepath if filepath else 'Unknown')}")
            return ""
        try:
            self.update_conversation_history(f"System: Extracting text from {os.path.basename(filepath)}...")
            # Using 'with' statement ensures the document is properly closed.
            with fitz.open(filepath) as pdf_doc:
                full_text = ""
                for page in pdf_doc: # Iterate through pages
                    full_text += page.get_text()
            self.update_conversation_history(f"System: Text extraction successful for {os.path.basename(filepath)}.")
            return full_text
        except FitzError as e: # Specific error for PyMuPDF
            error_msg = f"System: FitzError extracting text from {os.path.basename(filepath)} - {e}"
            print(error_msg) # Also print to console for debugging
            self.update_conversation_history(error_msg)
            return ""
        except FileNotFoundError: # Should be caught by the initial check, but as a safeguard.
            error_msg = f"System: File not found error during text extraction: {os.path.basename(filepath)}"
            print(error_msg)
            self.update_conversation_history(error_msg)
            return ""
        except Exception as e: # General fallback for any other unexpected errors
            error_msg = f"System: Unexpected error extracting text from {os.path.basename(filepath)} - {e}"
            print(error_msg)
            self.update_conversation_history(error_msg)
            return ""

    def extract_images_from_pdf(self, filepath, output_folder):
        """
        Extracts all images from a given PDF file and saves them to the output_folder.

        Args:
            filepath (str): The path to the PDF file.
            output_folder (str): The directory to save extracted images.

        Returns:
            list: A list of paths to the extracted image files, or an empty list on failure.
        """
        if not filepath or not os.path.exists(filepath):
            self.update_conversation_history(f"System: PDF file not found for image extraction: {os.path.basename(filepath if filepath else 'Unknown')}")
            return []

        extracted_image_paths = []
        try:
            self.update_conversation_history(f"System: Extracting images from {os.path.basename(filepath)}...")
            # Ensure output folder exists
            if not os.path.exists(output_folder):
                try:
                    os.makedirs(output_folder)
                except OSError as e:
                    self.update_conversation_history(f"System: Cannot create output folder for images: {output_folder}. Error: {e}")
                    return []


            with fitz.open(filepath) as pdf_doc: # Using 'with' for resource management
                for page_num in range(len(pdf_doc)):
                    page = pdf_doc.load_page(page_num) # Load page explicitly
                    image_list = page.get_images(full=True)
                    for img_index, img_info in enumerate(image_list):
                        xref = img_info[0]
                        try:
                            base_image = pdf_doc.extract_image(xref)
                        except Exception as e: # Handle error extracting a specific image
                            self.update_conversation_history(f"System: Error extracting image xref {xref} on page {page_num+1} from {os.path.basename(filepath)}. Skipping this image. Error: {e}")
                            continue # Skip to next image

                        image_bytes = base_image["image"]
                        image_ext = base_image["ext"]
                        image_filename = f"page{page_num+1}_img{img_index+1}.{image_ext}"
                        image_path = os.path.join(output_folder, image_filename)

                        try:
                            with open(image_path, "wb") as img_file:
                                img_file.write(image_bytes)
                            extracted_image_paths.append(image_path)
                        except IOError as e:
                             self.update_conversation_history(f"System: IOError saving image {image_filename} to {output_folder}. Error: {e}")

            if extracted_image_paths:
                self.update_conversation_history(f"System: Image extraction successful for {os.path.basename(filepath)}. Found {len(extracted_image_paths)} images.")
            else:
                self.update_conversation_history(f"System: No images successfully extracted from {os.path.basename(filepath)}.")
            return extracted_image_paths
        except FitzError as e: # Specific PyMuPDF error
            error_msg = f"System: FitzError extracting images from {os.path.basename(filepath)} - {e}"
            print(error_msg)
            self.update_conversation_history(error_msg)
            return []
        except FileNotFoundError: # Should be caught by initial check
            error_msg = f"System: File not found error during image extraction: {os.path.basename(filepath)}"
            print(error_msg)
            self.update_conversation_history(error_msg)
            return []
        except Exception as e: # General fallback
            error_msg = f"System: Unexpected error extracting images from {os.path.basename(filepath)} - {e}"
            print(error_msg)
            self.update_conversation_history(error_msg)
            return []

    def check_and_process_spec_sheets(self):
        """
        Checks if both spec sheets are loaded. If so, resets relevant AI state
        and initiates the automated analysis via process_spec_sheets.
        Ensures AI is configured before proceeding.
        """
        if self.spec_sheet_1_path and self.spec_sheet_2_path:
            self.update_conversation_history("System: Both spec sheets loaded. Preparing for analysis...")
            # Reset AI state for a new comparison.
            # This ensures that if a user loads new files, it's a fresh analysis.
            self.model = None
            self.chat_session = None

            # Clear previous conversation log for the new analysis context.
            # This is a design choice: each "Process Specs" starts a new log.
            # If appending to a master log is desired, this line would be removed.
            self.conversation_log = []
            if hasattr(self, 'conversation_history') and self.conversation_history: # Clear UI too
                self.conversation_history.config(state=tk.NORMAL)
                self.conversation_history.delete(1.0, tk.END)
                # No need to disable here, update_conversation_history will.

            self._configure_ai() # Re-check API key and update UI (especially button states)

            if not self.api_key_configured:
                 self.update_conversation_history("System: Cannot start analysis. AI is not configured (check GOOGLE_API_KEY).")
                 return # Stop if AI isn't ready

            self.process_spec_sheets() # Proceed to process

    def process_spec_sheets(self):
        """
        Orchestrates the extraction of data from loaded PDFs and sends it
        for an initial analysis to the AI. This is called after checks ensure
        both files are loaded and AI is configured.
        """
        # Defensive check, though check_and_process_spec_sheets should ensure this.
        if not self.spec_sheet_1_path or not self.spec_sheet_2_path:
            self.update_conversation_history("System: One or both spec sheets are not loaded. Cannot process.")
            return
        if not self.api_key_configured: # Another check for safety
            self.update_conversation_history("System: AI not configured. Cannot process spec sheets.")
            return

        # 1. Extract text and images
        self.spec_sheet_1_text = self.extract_text_from_pdf(self.spec_sheet_1_path)
        if not self.spec_sheet_1_text:
            self.update_conversation_history(f"System: Halting analysis. Failed to extract text from {os.path.basename(self.spec_sheet_1_path)}.")
            return
        # Create a unique subfolder for spec1 images to avoid name clashes if re-loading same file
        spec1_img_folder = os.path.join(self.temp_image_dir, f"{os.path.splitext(os.path.basename(self.spec_sheet_1_path))[0]}_images_run{len(os.listdir(self.temp_image_dir))}")
        self.spec_sheet_1_image_paths = self.extract_images_from_pdf(self.spec_sheet_1_path, spec1_img_folder)

        self.spec_sheet_2_text = self.extract_text_from_pdf(self.spec_sheet_2_path)
        if not self.spec_sheet_2_text:
            self.update_conversation_history(f"System: Halting analysis. Failed to extract text from {os.path.basename(self.spec_sheet_2_path)}.")
            return
        # Unique subfolder for spec2 images
        spec2_img_folder = os.path.join(self.temp_image_dir, f"{os.path.splitext(os.path.basename(self.spec_sheet_2_path))[0]}_images_run{len(os.listdir(self.temp_image_dir))}")
        self.spec_sheet_2_image_paths = self.extract_images_from_pdf(self.spec_sheet_2_path, spec2_img_folder)

        # 2. Construct prompt_parts and log a summary
        # This message goes to UI and the downloadable log
        initial_prompt_summary_message = (
            "System: Starting new analysis with the following inputs:\n"
            f"- Spec Sheet 1: {os.path.basename(self.spec_sheet_1_path)} (Text extracted, {len(self.spec_sheet_1_image_paths)} images extracted)\n"
            f"- Spec Sheet 2: {os.path.basename(self.spec_sheet_2_path)} (Text extracted, {len(self.spec_sheet_2_image_paths)} images extracted)\n"
            "- Analysis Request: Identify component types, crucial parameters, pin-to-pin compatibility, and key differences."
        )
        self.update_conversation_history(initial_prompt_summary_message)

        prompt_parts_for_ai = [
            "You are an expert electronics component analyst. You will be given text and images from two component specification sheets.",
            "\n--- Spec Sheet 1 Text ---", self.spec_sheet_1_text,
        ]
        for img_path in self.spec_sheet_1_image_paths:
            try:
                img = Image.open(img_path)
                # It's good practice to verify images, but PIL's verify() is basic.
                # For more robust validation, one might need other libraries or checks.
                # img.verify() # This closes the file, so must reopen if used.
                # img = Image.open(img_path) # Re-open after verify
                prompt_parts_for_ai.append(img)
            except FileNotFoundError:
                self.update_conversation_history(f"System: Image file not found: {img_path}. Skipping this image.")
            except UnidentifiedImageError: # PIL specific error for bad image files
                 self.update_conversation_history(f"System: Cannot identify image file (possibly corrupted or unsupported format): {img_path}. Skipping this image.")
            except Exception as e: # General fallback for other image loading errors
                self.update_conversation_history(f"System: Error loading image {img_path} for AI. Skipping this image. Error: {e}")

        prompt_parts_for_ai.extend([
            "\n--- Spec Sheet 2 Text ---", self.spec_sheet_2_text,
        ])
        for img_path in self.spec_sheet_2_image_paths:
            try:
                img = Image.open(img_path)
                # img.verify()
                # img = Image.open(img_path)
                prompt_parts_for_ai.append(img)
            except FileNotFoundError:
                self.update_conversation_history(f"System: Image file not found: {img_path}. Skipping this image.")
            except UnidentifiedImageError:
                 self.update_conversation_history(f"System: Cannot identify image file: {img_path}. Skipping this image.")
            except Exception as e:
                self.update_conversation_history(f"System: Error loading image {img_path} for AI. Skipping. Error: {e}")

        prompt_parts_for_ai.append(
            "\n--- Analysis Request ---" # Clearer separation for the request
            "\nBased on the provided information for the two components:"
            "\n1. Identify the component type for each document (e.g., MOSFET, TVS Diode, Resistor, Capacitor, Inductor, Op-Amp, etc.)."
            "\n2. Based on these identified types, generate a prioritized list of parameters that are crucial for comparing such components."
            "\n3. Compare the two components for pin-to-pin compatibility. Clearly state if they are compatible, potentially compatible (with notes), or not compatible, and why."
            "\n4. List all key electrical and physical specification differences in a structured format (e.g., table or bullet points)."
            "\nPresent your analysis clearly and concisely." # Added conciseness instruction
        )

        self.send_to_ai(prompt_parts_for_ai, is_initial_analysis=True)

    def send_to_ai(self, prompt_parts, is_initial_analysis=False):
        """
        Sends the constructed prompt (with text and images) to the Generative AI model.
        Manages UI state of the send button during AI interaction.

        Args:
            prompt_parts (list): The list of parts (text, images) forming the prompt.
            is_initial_analysis (bool): True if this is the first analysis pass.
        """
        if not self.api_key_configured: # Should have been checked before, but good safeguard
            self.update_conversation_history("System: AI is not configured. Cannot send request.")
            return

        selected_model_name = self.get_selected_model_name()
        if not selected_model_name:
            self.update_conversation_history("System: Please select an AI model first.")
            return

        try:
            self.send_button.config(state=tk.DISABLED) # Disable send button
            self.user_input_entry.config(state=tk.DISABLED) # Disable input field

            # Initialize or change model if it's the initial analysis or model selection changed
            if is_initial_analysis or not self.model or self.model.model_name != selected_model_name:
                 self.update_conversation_history(f"System: Initializing AI model: {selected_model_name}...")
                 # Add generation_config for safety settings if available/desired
                 # safety_settings = [
                 #     {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
                 #     {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
                 # ]
                 # self.model = genai.GenerativeModel(selected_model_name, safety_settings=safety_settings)
                 self.model = genai.GenerativeModel(selected_model_name)
                 self.chat_session = None # Reset chat for new model/initial analysis

            self.update_conversation_history(f"System: Sending request to AI ({selected_model_name})... This may take some time.")

            # For initial analysis, use generate_content. For chat, use chat_session.send_message.
            # This method is now only called for the initial analysis or direct generation.
            # send_user_query handles chat session messages.
            response = self.model.generate_content(prompt_parts)

            # Process potential partial or no response
            if response.prompt_feedback and response.prompt_feedback.block_reason:
                self.update_conversation_history(f"System: AI Error - Prompt was blocked. Reason: {response.prompt_feedback.block_reason}")
            elif not response.candidates or not response.text: # Check if response.text is empty or candidates list is empty
                self.update_conversation_history(f"System: AI ({selected_model_name}): Received no content or an empty response.")
            else:
                self.update_conversation_history(f"AI ({selected_model_name}): {response.text}")

            if is_initial_analysis:
                # After the initial analysis, subsequent user queries will use the chat session.
                # We ensure chat_session is None now so send_user_query initializes it cleanly for follow-ups.
                self.chat_session = None

        except google_exceptions.InvalidArgument as e:
            self.update_conversation_history(f"System: AI Error - Invalid argument (often an issue with the prompt or image data): {e}")
        except google_exceptions.PermissionDenied as e:
            self.update_conversation_history(f"System: AI Error - Permission denied. This might be an API key issue or service permissions. {e}")
            self.api_key_configured = False # API key might be invalid
            self._update_ui_for_ai_status()
        except google_exceptions.ServiceUnavailable as e:
            self.update_conversation_history(f"System: AI Error - The service is currently unavailable. Please try again later: {e}")
        except google_exceptions.GoogleAPIError as e: # Catch other specific Google API errors
            self.update_conversation_history(f"System: Google API Error during AI content generation: {e}")
        except BlockedPromptException as e: # Specific exception for blocked prompts
            self.update_conversation_history(f"System: AI Error - Your prompt was blocked by safety settings. Details: {e}")
        except StopCandidateException as e: # If generation stops unexpectedly due to safety or other reasons
             self.update_conversation_history(f"System: AI Error - Content generation stopped unexpectedly. Details: {e}")
        except ValueError as e: # E.g. invalid model name during genai.GenerativeModel()
             self.update_conversation_history(f"System: Error during AI setup (check model name or API key configuration): {e}")
             self.model = None # Reset model as it might be misconfigured
             self.chat_session = None
        except Exception as e: # General fallback for any other unexpected errors
            self.update_conversation_history(f"System: An unexpected error occurred during AI content generation: {e}")
        finally:
            # Re-enable send button and input field if API key is still considered configured
            if self.api_key_configured:
                self.send_button.config(state=tk.NORMAL)
                self.user_input_entry.config(state=tk.NORMAL)
            else: # Ensure they remain disabled if API key failed
                self.send_button.config(state=tk.DISABLED)
                self.user_input_entry.config(state=tk.DISABLED)


def main():
    """Main function to create and run the Tkinter application."""
    # Attempt to set a more modern theme if available (for ttk widgets)
    # This is optional and platform-dependent.
    try:
        s = ttk.Style()
        available_themes = s.theme_names()
        # Prioritize modern themes
        if "clam" in available_themes: # Good cross-platform
            s.theme_use("clam")
        elif "aqua" in available_themes: # macOS
             s.theme_use("aqua")
        elif "vista" in available_themes: # Windows
            s.theme_use("vista")
        # 'alt', 'default', 'classic' are other fallbacks but usually less modern.
    except tk.TclError:
        # This can happen if Tkinter is not fully initialized or themes are unavailable.
        print("Could not set a custom ttk theme; using default.")
    except Exception: # Catch any other theming errors silently
        pass

    root = tk.Tk()
    # It's good practice to put the app instance in a variable.
    app = ComponentComparatorAI(root)
    root.mainloop()

if __name__ == "__main__":
    # This ensures that the main() function is called only when the script is executed directly,
    # not when imported as a module.
    main()
