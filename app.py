import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox, simpledialog
import docx
import re
from typing import List, Tuple, Optional
import unicodedata
import os
from llm_providers import LLMModelRegistry, ClaudeProvider, OpenAIProvider

# Try to import tkinterdnd2 for drag-and-drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    print("Note: tkinterdnd2 not available. Drag-and-drop will be disabled.")
    print("Install it with: pip install tkinterdnd2")


class WordProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word Document Processor with LLM API")
        self.root.geometry("1200x900")
        
        # API Configuration - Separate keys for Claude and OpenAI
        self.claude_api_key = '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'
        self.openai_api_key = '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'
        
        # Initialize LLM Registry
        self.llm_registry = LLMModelRegistry()
        
        # Register providers with their respective API keys
        try:
            if self.claude_api_key:
                claude_provider = ClaudeProvider(self.claude_api_key)
                self.llm_registry.register_provider("claude", claude_provider)
        except Exception as e:
            print(f"Warning: Could not initialize Claude provider: {e}")
        
        try:
            if self.openai_api_key:
                openai_provider = OpenAIProvider(self.openai_api_key)
                self.llm_registry.register_provider("openai", openai_provider)
        except Exception as e:
            print(f"Warning: Could not initialize OpenAI provider: {e}")
        
        # Default model (first available model)
        available_models = self.llm_registry.get_all_models()
        self.selected_model = available_models[0] if available_models else None
        
        # Development/Testing: Hardcoded file path (set to None to disable, or provide full path)
        # Example: self.hardcoded_file_path = r"C:\Users\bob\Documents\test_document.docx"
        #self.hardcoded_file_path = r"C:\Users\bob\Music\test.docx"  # Set this to your test file path for development
        self.hardcoded_file_path = None
        
        # Data storage
        self.full_text = ""
        self.extracted_text = ""
        self.masked_text = ""
        self.masking_changes = []  # List of individual occurrence changes
        self.current_changes = []  # Current list of changes for undo
        self.name_to_id = {}  # Maps normalized name to its unique ID
        self.name_occurrences = {}  # Maps normalized name to list of occurrences
        
        # Instructions storage
        self.instructions_file = "instructions.txt"
        self.instructions_dict = {}  # Dictionary to store instructions: {label: text}
        self.current_instruction_label = "basic"
        
        # Chat messages storage
        self.chat_file = "chat.txt"
        self.chat_dict = {}  # Dictionary to store chat messages: {label: text}
        self.current_chat_label = "basic"
        
        # Conversation history for chat functionality
        self.conversation_history = []  # List of messages: [{"role": "user"/"assistant", "content": "..."}]
        self.is_first_message = True  # Track if this is the first API call
        
        # Load saved instructions and chat messages
        self.load_instructions()
        self.load_chat_messages()
        
        # Create GUI
        self.create_widgets()
        
        # Auto-load hardcoded file if specified (for development/testing)
        if self.hardcoded_file_path:
            import os
            if os.path.exists(self.hardcoded_file_path):
                self.file_path_var.set(self.hardcoded_file_path)
                self.load_document(self.hardcoded_file_path)
            else:
                print(f"Warning: Hardcoded file path does not exist: {self.hardcoded_file_path}")
        
    def create_widgets(self):
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Tab 1: Text Extraction
        self.tab1 = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.tab1, text="1. Text Extraction")
        self.create_tab1()
        
        # Tab 2: Name Masking
        self.tab2 = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.tab2, text="2. Name Masking")
        self.create_tab2()
        
        # Tab 3: API & Results
        self.tab3 = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.tab3, text="3. API & Results")
        self.create_tab3()
    
    def create_tab1(self):
        """Create Tab 1: Text Extraction"""
        self.tab1.columnconfigure(1, weight=1)
        
        # File upload section
        ttk.Label(self.tab1, text="Word Document:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.file_path_var = tk.StringVar()
        
        # Create a frame for the file path entry with drag-and-drop support
        file_entry_frame = ttk.Frame(self.tab1)
        file_entry_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        file_entry_frame.columnconfigure(0, weight=1)
        
        self.file_path_entry = ttk.Entry(file_entry_frame, textvariable=self.file_path_var, width=50, state="readonly")
        self.file_path_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Enable drag-and-drop if available
        if DND_AVAILABLE:
            try:
                self.file_path_entry.drop_target_register(DND_FILES)
                self.file_path_entry.dnd_bind('<<Drop>>', self.on_file_drop)
                # Also enable drop on the tab itself
                self.tab1.drop_target_register(DND_FILES)
                self.tab1.dnd_bind('<<Drop>>', self.on_file_drop)
            except Exception as e:
                print(f"Warning: Could not enable drag-and-drop: {e}")
        
        ttk.Button(self.tab1, text="Browse", command=self.browse_file).grid(row=0, column=2, padx=5)
        
        # Add hint label for drag-and-drop
        if DND_AVAILABLE:
            hint_label = ttk.Label(self.tab1, text="(Drag and drop .docx file here)", font=("TkDefaultFont", 8), foreground="gray")
            hint_label.grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # Extraction range section
        ttk.Label(self.tab1, text="Start Word (wordA):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.start_word_var = tk.StringVar(value="commemoratifs")
        ttk.Entry(self.tab1, textvariable=self.start_word_var, width=30).grid(row=1, column=1, sticky=tk.W, padx=5)
        
        ttk.Label(self.tab1, text="End Word (wordB):").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.end_word_var = tk.StringVar(value="documents presentes")
        ttk.Entry(self.tab1, textvariable=self.end_word_var, width=30).grid(row=2, column=1, sticky=tk.W, padx=5)
        
        # Button frame for extraction buttons
        button_frame = ttk.Frame(self.tab1)
        button_frame.grid(row=2, column=2, padx=5, sticky=tk.W)
        ttk.Button(button_frame, text="Extract Text", command=self.extract_text).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Undo Extraction", command=self.undo_extraction).pack(side=tk.LEFT, padx=2)
        
        # Extracted text display (editable)
        extracted_label_frame = ttk.Frame(self.tab1)
        extracted_label_frame.grid(row=3, column=0, sticky=(tk.W, tk.N), pady=5)
        ttk.Label(extracted_label_frame, text="Extracted Text (editable):").pack(side=tk.LEFT)
        
        # Navigation button to next tab
        nav_frame = ttk.Frame(self.tab1)
        nav_frame.grid(row=3, column=2, padx=5, sticky=tk.N)
        ttk.Button(nav_frame, text="Send to Masking →", command=self.sync_to_masking).pack(side=tk.TOP, pady=2)
        
        self.extracted_text_area = scrolledtext.ScrolledText(self.tab1, height=15, width=80, wrap=tk.WORD)
        self.extracted_text_area.grid(row=3, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # Configure grid weights for resizing
        self.tab1.rowconfigure(3, weight=1)
    
    def create_tab2(self):
        """Create Tab 2: Name Masking"""
        self.tab2.columnconfigure(1, weight=1)
        
        # Navigation button to previous tab
        nav_frame_top = ttk.Frame(self.tab2)
        nav_frame_top.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=5)
        ttk.Button(nav_frame_top, text="← Back to Extraction", command=self.go_to_extraction_tab).pack(side=tk.LEFT, padx=5)
        
        # Names to mask section
        ttk.Label(self.tab2, text="Names/Surnames to Mask (comma-separated):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.names_var = tk.StringVar()
        ttk.Entry(self.tab2, textvariable=self.names_var, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(self.tab2, text="Apply Masking", command=self.apply_masking).grid(row=1, column=2, padx=5)
        
        # Masking preview
        ttk.Label(self.tab2, text="Masking Preview:").grid(row=2, column=0, sticky=(tk.W, tk.N), pady=5)
        self.masking_preview_area = scrolledtext.ScrolledText(self.tab2, height=10, width=80, wrap=tk.WORD)
        self.masking_preview_area.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # Changes list
        ttk.Label(self.tab2, text="Changes (click to undo):").grid(row=3, column=0, sticky=(tk.W, tk.N), pady=5)
        changes_frame = ttk.Frame(self.tab2)
        changes_frame.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.changes_listbox = tk.Listbox(changes_frame, height=5, width=50)
        self.changes_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.changes_listbox.bind('<Double-Button-1>', self.undo_change)
        ttk.Button(changes_frame, text="Undo Selected", command=self.undo_selected_change).pack(side=tk.LEFT, padx=5)
        
        # Navigation button to next tab
        nav_frame_bottom = ttk.Frame(self.tab2)
        nav_frame_bottom.grid(row=4, column=0, columnspan=3, sticky=tk.E, pady=10)
        ttk.Button(nav_frame_bottom, text="Continue to API →", command=self.go_to_api_tab).pack(side=tk.RIGHT, padx=5)
        
        # Configure grid weights for resizing
        self.tab2.rowconfigure(2, weight=1)
    
    def create_tab3(self):
        """Create Tab 3: API & Results"""
        self.tab3.columnconfigure(1, weight=1)
        
        # Navigation button to previous tab
        nav_frame_top = ttk.Frame(self.tab3)
        nav_frame_top.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=5)
        ttk.Button(nav_frame_top, text="← Back to Masking", command=self.go_to_masking_tab).pack(side=tk.LEFT, padx=5)
        
        # API instructions section
        # Model selection (top right)
        model_selection_frame = ttk.Frame(self.tab3)
        model_selection_frame.grid(row=1, column=0, columnspan=3, sticky=tk.E, padx=5, pady=5)
        ttk.Label(model_selection_frame, text="Model:").pack(side=tk.LEFT, padx=5)
        self.model_var = tk.StringVar()
        self.model_combo = ttk.Combobox(model_selection_frame, textvariable=self.model_var, 
                                        width=30, state="readonly")
        self.model_combo.pack(side=tk.LEFT, padx=5)
        self.model_combo.bind('<<ComboboxSelected>>', self.on_model_selected)
        self.update_model_combo()
        
        instructions_label_frame = ttk.Frame(self.tab3)
        instructions_label_frame.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
        ttk.Label(instructions_label_frame, text="LLM API Instructions:").pack(side=tk.LEFT, padx=5)
        
        # Instruction label selection
        ttk.Label(instructions_label_frame, text="Label:").pack(side=tk.LEFT, padx=5)
        self.instruction_label_var = tk.StringVar()
        self.instruction_label_combo = ttk.Combobox(instructions_label_frame, textvariable=self.instruction_label_var, 
                                                    width=20, state="readonly")
        self.instruction_label_combo.pack(side=tk.LEFT, padx=5)
        self.instruction_label_combo.bind('<<ComboboxSelected>>', self.on_instruction_label_selected)
        
        # Buttons for managing instructions
        buttons_frame = ttk.Frame(instructions_label_frame)
        buttons_frame.pack(side=tk.LEFT, padx=10)
        ttk.Button(buttons_frame, text="Save", command=self.save_instruction).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons_frame, text="Create New", command=self.create_new_instruction).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons_frame, text="Delete", command=self.delete_instruction).pack(side=tk.LEFT, padx=2)
        
        # Instruction text area (editable)
        ttk.Label(self.tab3, text="Instruction Text:").grid(row=3, column=0, sticky=(tk.W, tk.N), pady=5)
        self.instructions_text_area = scrolledtext.ScrolledText(self.tab3, height=5, width=80, wrap=tk.WORD)
        self.instructions_text_area.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # Send button
        ttk.Button(self.tab3, text="Send to LLM API", command=self.send_to_api).grid(row=4, column=1, columnspan=2, pady=5)
        
        # Update instruction combo and load default
        self.update_instruction_combo()
        if "basic" in self.instructions_dict:
            self.instruction_label_var.set("basic")
            self.on_instruction_label_selected()
        
        # Final text display
        ttk.Label(self.tab3, text="Final Text (from LLM):").grid(row=5, column=0, sticky=(tk.W, tk.N), pady=5)
        final_text_frame = ttk.Frame(self.tab3)
        final_text_frame.grid(row=5, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        self.final_text_area = scrolledtext.ScrolledText(final_text_frame, height=30, width=80, wrap=tk.WORD)
        self.final_text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Button(final_text_frame, text="Copy Final Text", command=self.copy_final_text).pack(side=tk.LEFT, padx=5)
        
        # Chat messages section (below final text)
        chat_label_frame = ttk.Frame(self.tab3)
        chat_label_frame.grid(row=6, column=0, columnspan=3, sticky=tk.W, pady=5)
        ttk.Label(chat_label_frame, text="Continue conversation:").pack(side=tk.LEFT, padx=5)
        
        # Chat message template selection
        chat_template_frame = ttk.Frame(self.tab3)
        chat_template_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        ttk.Label(chat_template_frame, text="Chat Message Label:").pack(side=tk.LEFT, padx=5)
        self.chat_label_var = tk.StringVar()
        self.chat_label_combo = ttk.Combobox(chat_template_frame, textvariable=self.chat_label_var, 
                                               width=20, state="readonly")
        self.chat_label_combo.pack(side=tk.LEFT, padx=5)
        self.chat_label_combo.bind('<<ComboboxSelected>>', self.on_chat_label_selected)
        
        # Buttons for managing chat messages
        chat_buttons_frame = ttk.Frame(chat_template_frame)
        chat_buttons_frame.pack(side=tk.LEFT, padx=10)
        ttk.Button(chat_buttons_frame, text="Save", command=self.save_chat_message).pack(side=tk.LEFT, padx=2)
        ttk.Button(chat_buttons_frame, text="Create New", command=self.create_new_chat_message).pack(side=tk.LEFT, padx=2)
        ttk.Button(chat_buttons_frame, text="Delete", command=self.delete_chat_message).pack(side=tk.LEFT, padx=2)
        
        # Chat input section
        chat_input_frame = ttk.Frame(self.tab3)
        chat_input_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), padx=5, pady=5)
        chat_input_frame.columnconfigure(0, weight=1)
        
        self.chat_input = ttk.Entry(chat_input_frame, width=60)
        self.chat_input.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5)
        self.chat_input.bind('<Return>', lambda e: self.send_chat_message())
        ttk.Button(chat_input_frame, text="Send", command=self.send_chat_message).grid(row=0, column=1, padx=5)
        ttk.Button(chat_input_frame, text="Clear History", command=self.clear_conversation_history).grid(row=0, column=2, padx=5)
        
        # Update chat combo and load default
        self.update_chat_combo()
        if "basic" in self.chat_dict:
            self.chat_label_var.set("basic")
            self.on_chat_label_selected()
        
        # Configure grid weights for resizing
        self.tab3.rowconfigure(2, weight=1)
        self.tab3.rowconfigure(4, weight=1)
        self.tab3.rowconfigure(5, weight=2)  # Give more weight to final text area
    
    def go_to_extraction_tab(self):
        """Navigate to Tab 1: Text Extraction"""
        self.notebook.select(0)
    
    def go_to_masking_tab(self):
        """Navigate to Tab 2: Name Masking"""
        # Sync extracted text to masking if needed (read from text area to get any edits)
        if self.extracted_text_area.get(1.0, tk.END).strip():
            # Get current text from the text area (may have been edited)
            current_text = self.extracted_text_area.get(1.0, tk.END).rstrip('\n')
            if current_text:
                # Update extracted_text with current text from widget
                self.extracted_text = current_text
                # Update masked_text if not already set or if text changed
                if not self.masked_text or self.masked_text != self.extracted_text:
                    self.masked_text = self.extracted_text
                    # Update the masking preview
                    self.masking_preview_area.delete(1.0, tk.END)
                    self.masking_preview_area.insert(1.0, self.masked_text)
        self.notebook.select(1)
    
    def go_to_api_tab(self):
        """Navigate to Tab 3: API & Results"""
        self.notebook.select(2)
        
    def normalize_text(self, text: str) -> str:
        """Normalize text to remove accents and convert to lowercase for comparison"""
        # Remove accents
        nfd = unicodedata.normalize('NFD', text)
        text_no_accents = ''.join(c for c in nfd if unicodedata.category(c) != 'Mn')
        return text_no_accents.lower()
    
    def find_word_ignore_case_accent(self, text: str, word: str) -> Optional[int]:
        """Find the first occurrence of a word ignoring case and accents"""
        normalized_word = self.normalize_text(word)
        normalized_text = self.normalize_text(text)
        
        # Find word boundaries
        pattern = r'\b' + re.escape(normalized_word) + r'\b'
        match = re.search(pattern, normalized_text)
        
        if match:
            # Build position mapping
            position_map = []
            for orig_pos, char in enumerate(text):
                nfd = unicodedata.normalize('NFD', char)
                normalized_char = ''.join(c for c in nfd if unicodedata.category(c) != 'Mn')
                normalized_lower = normalized_char.lower()
                for _ in range(len(normalized_lower)):
                    position_map.append(orig_pos)
            
            # Map normalized position to original
            norm_start = match.start()
            if norm_start < len(position_map):
                return position_map[norm_start]
        
        return None
    
    def on_file_drop(self, event):
        """Handle file drop event"""
        # Get the dropped file path(s)
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0].strip('{}')  # Remove curly braces that Windows adds
            # Validate it's a .docx file
            if file_path.lower().endswith('.docx'):
                self.file_path_var.set(file_path)
                self.load_document(file_path)
            else:
                messagebox.showwarning("Warning", "Please drop a .docx file (Word document).")
    
    def browse_file(self):
        """Open file dialog to select Word document"""
        file_path = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.load_document(file_path)
    
    def load_document(self, file_path: str):
        """Load and extract text from Word document"""
        try:
            doc = docx.Document(file_path)
            self.full_text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
            
            # Automatically extract the entire document content initially
            self.extracted_text = self.full_text
            self.masked_text = self.full_text
            
            # Clear any existing masking data
            self.masking_changes = []
            self.current_changes = []
            self.name_to_id = {}
            self.name_occurrences = {}
            
            # Display the full text in the extracted text area
            self.extracted_text_area.delete(1.0, tk.END)
            self.extracted_text_area.insert(1.0, self.extracted_text)
            
            # Clear masking preview and changes list
            self.masking_preview_area.delete(1.0, tk.END)
            self.changes_listbox.delete(0, tk.END)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load document: {str(e)}")
    
    def extract_text(self):
        """Extract text between start and end words"""
        if not self.full_text:
            messagebox.showwarning("Warning", "Please load a document first.")
            return
        
        start_word = self.start_word_var.get().strip()
        end_word = self.end_word_var.get().strip()
        
        if not start_word or not end_word:
            messagebox.showwarning("Warning", "Please enter both start and end words.")
            return
        
        # Find start position
        start_pos = self.find_word_ignore_case_accent(self.full_text, start_word)
        if start_pos is None:
            messagebox.showwarning("Warning", f"Start word '{start_word}' not found in document.")
            return
        
        # Find end position (search from start position)
        text_from_start = self.full_text[start_pos:]
        end_pos_relative = self.find_word_ignore_case_accent(text_from_start, end_word)
        if end_pos_relative is None:
            messagebox.showwarning("Warning", f"End word '{end_word}' not found after start word.")
            return
        
        # Exclude the end word from extraction - stop at the start of the end word
        end_pos = start_pos + end_pos_relative
        
        # Extract text (excluding the end word)
        self.extracted_text = self.full_text[start_pos:end_pos]
        self.masked_text = self.extracted_text
        self.masking_changes = []
        self.current_changes = []
        self.name_to_id = {}
        self.name_occurrences = {}
        
        # Display extracted text
        self.extracted_text_area.delete(1.0, tk.END)
        self.extracted_text_area.insert(1.0, self.extracted_text)
        
        # Clear masking preview
        self.masking_preview_area.delete(1.0, tk.END)
        self.masking_preview_area.insert(1.0, self.extracted_text)
        
        # Clear changes list
        self.changes_listbox.delete(0, tk.END)
        
    
    def undo_extraction(self):
        """Undo text extraction - clears extracted text and all masking"""
        if not self.extracted_text:
            return
        
        # Confirm with user
        if not messagebox.askyesno("Confirm", "This will clear the extracted text and all masking. Continue?"):
            return
        
        # Clear extracted text
        self.extracted_text = ""
        self.masked_text = ""
        
        # Clear all masking data
        self.masking_changes = []
        self.current_changes = []
        self.name_to_id = {}
        self.name_occurrences = {}
        
        # Clear all text areas
        self.extracted_text_area.delete(1.0, tk.END)
        self.masking_preview_area.delete(1.0, tk.END)
        self.final_text_area.delete(1.0, tk.END)
        
        # Clear changes list
        self.changes_listbox.delete(0, tk.END)
        
    
    def sync_to_masking(self):
        """Sync the edited extracted text to the masking preview"""
        # Get the current text from the extracted text area
        edited_text = self.extracted_text_area.get(1.0, tk.END).rstrip('\n')
        
        if not edited_text.strip():
            messagebox.showwarning("Warning", "Extracted text area is empty.")
            return
        
        # Update the extracted_text variable with the edited text
        self.extracted_text = edited_text
        
        # Clear any existing masking when syncing new text
        # Ask user if they want to keep existing masking
        if self.current_changes:
            if not messagebox.askyesno("Confirm", "This will clear all existing masking. Continue?"):
                return
            # Clear all masking data
            self.masking_changes = []
            self.current_changes = []
            self.name_to_id = {}
            self.name_occurrences = {}
            self.changes_listbox.delete(0, tk.END)
        
        # Update masked_text to match the edited extracted text
        self.masked_text = self.extracted_text
        
        # Update the masking preview
        self.masking_preview_area.delete(1.0, tk.END)
        self.masking_preview_area.insert(1.0, self.masked_text)
        
        # Navigate to masking tab
        self.go_to_masking_tab()
    
    def find_name_ignore_case_accent(self, text: str, name: str) -> List[Tuple[int, int, str]]:
        """Find all occurrences of a name ignoring case and accents.
        Returns list of (start_pos, end_pos, original_text) tuples.
        Handles both single-word and multi-word names (e.g., "John" or "John Smith")."""
        normalized_name = self.normalize_text(name)
        results = []
        
        # Check if name contains spaces (multi-word name)
        name_has_spaces = ' ' in name
        
        if name_has_spaces:
            # For multi-word names, match the entire phrase with word boundaries at start and end
            # Escape the name and replace spaces with \s+ to handle multiple spaces
            escaped_name = re.escape(normalized_name)
            # Replace escaped spaces with \s+ to handle variable whitespace
            pattern = r'\b' + escaped_name.replace(r'\ ', r'\s+') + r'\b'
        else:
            # For single-word names, use word boundaries
            pattern = r'\b' + re.escape(normalized_name) + r'\b'
        
        # Normalize the text for searching
        normalized_text = self.normalize_text(text)
        matches = list(re.finditer(pattern, normalized_text))
        
        # Build position mapping
        position_map = []
        for orig_pos, char in enumerate(text):
            nfd = unicodedata.normalize('NFD', char)
            normalized_char = ''.join(c for c in nfd if unicodedata.category(c) != 'Mn')
            normalized_lower = normalized_char.lower()
            for _ in range(len(normalized_lower)):
                position_map.append(orig_pos)
        
        # Map back to original positions
        for match in matches:
            norm_start = match.start()
            norm_end = match.end()
            
            # Find corresponding positions in original text
            if norm_start < len(position_map) and norm_end <= len(position_map):
                orig_start = position_map[norm_start] if norm_start < len(position_map) else position_map[-1]
                # For end position, find the corresponding position in original text
                if norm_end > 0 and norm_end <= len(position_map):
                    orig_end = position_map[norm_end - 1] + 1
                else:
                    orig_end = orig_start
                
                # For multi-word names, we need to find the exact end of the phrase in original text
                if name_has_spaces:
                    # Find the end of the full phrase by looking for the last word
                    # Start from orig_start and find the full phrase
                    phrase_end = orig_start
                    words = name.split()
                    word_count = 0
                    i = orig_start
                    
                    # Skip leading whitespace
                    while i < len(text) and text[i].isspace():
                        i += 1
                    
                    # Try to match all words in the phrase
                    for word_idx, word in enumerate(words):
                        if i >= len(text):
                            break
                        
                        # Find the start of this word
                        word_start = i
                        # Find the end of this word (alphanumeric + apostrophes/hyphens)
                        while i < len(text) and (text[i].isalnum() or text[i] in "'-"):
                            i += 1
                        word_end = i
                        
                        # Check if this word matches (case/accent insensitive)
                        word_text = text[word_start:word_end]
                        if self.normalize_text(word_text) == self.normalize_text(word):
                            word_count += 1
                            phrase_end = word_end
                            # Skip whitespace between words
                            while i < len(text) and text[i].isspace():
                                i += 1
                        else:
                            break
                    
                    # If we matched all words, use the phrase_end
                    if word_count == len(words):
                        orig_end = phrase_end
                    else:
                        # Fallback: use the normalized position mapping
                        if norm_end > 0 and norm_end <= len(position_map):
                            orig_end = position_map[norm_end - 1] + 1
                else:
                    # For single-word names, find the end of the word
                    if orig_start < len(text):
                        word_end = orig_start
                        while word_end < len(text) and (text[word_end].isalnum() or text[word_end] in "'-"):
                            word_end += 1
                        orig_end = word_end
                
                # Extract the original text
                if orig_start < len(text) and orig_end <= len(text):
                    original_text = text[orig_start:orig_end]
                    results.append((orig_start, orig_end, original_text))
        
        return results
    
    def map_normalized_to_original(self, text: str, normalized_pos: int) -> Optional[int]:
        """Map a position in normalized text back to original text position"""
        # Build normalized text character by character and track positions
        normalized_chars = []
        position_map = []  # Maps normalized index to original index
        
        for orig_pos, char in enumerate(text):
            nfd = unicodedata.normalize('NFD', char)
            normalized_char = ''.join(c for c in nfd if unicodedata.category(c) != 'Mn')
            normalized_lower = normalized_char.lower()
            
            # Map each normalized character position back to original
            for _ in range(len(normalized_lower)):
                position_map.append(orig_pos)
                normalized_chars.append(normalized_lower[_] if _ < len(normalized_lower) else '')
        
        if normalized_pos < len(position_map):
            return position_map[normalized_pos]
        return None
    
    def apply_masking(self):
        """Apply masking to names/surnames"""
        if not self.extracted_text:
            messagebox.showwarning("Warning", "Please extract text first.")
            return
        
        names_input = self.names_var.get().strip()
        if not names_input:
            messagebox.showwarning("Warning", "Please enter names/surnames to mask.")
            return
        
        # Parse names (comma-separated)
        names = [name.strip() for name in names_input.split(",") if name.strip()]
        if not names:
            messagebox.showwarning("Warning", "No valid names found.")
            return
        
        # Always search in extracted_text to get correct positions
        # But we need to account for already-masked text when searching
        # So we'll search in extracted_text and filter out already-masked areas
        new_changes = []
        total_occurrences = 0
        
        # Build a set of already-masked positions to avoid re-masking
        masked_positions = set()
        for change in self.current_changes:
            for pos in range(change['position'], change['position'] + change['length']):
                masked_positions.add(pos)
        
        # Mask each name (handling accents and case)
        for name in names:
            # Normalize name for consistent lookup
            normalized_name = self.normalize_text(name)
            
            # Check if this name is already masked
            if normalized_name in self.name_to_id:
                # Name already masked, skip
                continue
            
            # Find all occurrences in extracted_text (ignoring case and accents)
            # IMPORTANT: Always search in extracted_text, never in masked_text
            occurrences = self.find_name_ignore_case_accent(self.extracted_text, name)
            
            # Filter out occurrences that are already masked
            new_occurrences = []
            for start_pos, end_pos, original_text in occurrences:
                # Check if this occurrence overlaps with already-masked text
                is_masked = any(pos in masked_positions for pos in range(start_pos, end_pos))
                if not is_masked:
                    # Additional check: verify the text at this position matches what we expect
                    actual_text = self.extracted_text[start_pos:end_pos]
                    # The actual_text might have different case/accents, so we normalize for comparison
                    normalized_actual = self.normalize_text(actual_text)
                    normalized_name = self.normalize_text(name)
                    normalized_original = self.normalize_text(original_text)
                    
                    # Only add if it matches (accounting for normalization)
                    if (normalized_actual == normalized_name or 
                        normalized_actual == normalized_original or 
                        actual_text == original_text):
                        # Double-check: ensure this position doesn't conflict with any existing change
                        conflicts = False
                        for existing_change in self.current_changes:
                            ex_start = existing_change['position']
                            ex_end = ex_start + existing_change['length']
                            # Check for any overlap
                            if not (end_pos <= ex_start or ex_end <= start_pos):
                                conflicts = True
                                break
                        
                        if not conflicts:
                            new_occurrences.append((start_pos, end_pos, original_text))
            
            if not new_occurrences:
                continue
            
            # Assign a unique ID for this name (based on number of already masked names)
            name_id = len(self.name_to_id) + 1
            masked = f"[NAME_{name_id}]"
            
            # Store the mapping
            self.name_to_id[normalized_name] = {
                'id': name_id,
                'masked': masked,
                'original_name': name  # Keep original for display
            }
            
            # Store all occurrences for this name
            name_occurrence_list = []
            for start_pos, end_pos, original_text in new_occurrences:
                change_info = {
                    'original': original_text,
                    'masked': masked,
                    'position': start_pos,
                    'length': len(original_text),
                    'normalized_name': normalized_name
                }
                new_changes.append(change_info)
                name_occurrence_list.append(change_info)
                total_occurrences += 1
                
                # Mark these positions as masked
                for pos in range(start_pos, end_pos):
                    masked_positions.add(pos)
            
            self.name_occurrences[normalized_name] = name_occurrence_list
        
        if not new_changes:
            return
        
        # Add to changes list
        self.current_changes.extend(new_changes)
        self.masking_changes.extend(new_changes)
        
        # Rebuild masked text from extracted_text with all changes
        self.rebuild_masked_text()
        
        # Update preview
        self.masking_preview_area.delete(1.0, tk.END)
        self.masking_preview_area.insert(1.0, self.masked_text)
        
        # Update changes listbox
        self.update_changes_listbox()
        
    
    def rebuild_masked_text(self):
        """Rebuild masked text from extracted_text using all current changes"""
        # Always start from extracted_text to ensure clean rebuild
        self.masked_text = self.extracted_text
        
        # Sort changes by position (reverse order to avoid position shifts)
        sorted_changes = sorted(self.current_changes, key=lambda x: x['position'], reverse=True)
        
        # Verify no overlapping changes and validate positions
        for i, change1 in enumerate(sorted_changes):
            pos1 = change1['position']
            len1 = change1['length']
            
            # Check bounds
            if pos1 < 0 or pos1 + len1 > len(self.extracted_text):
                print(f"Warning: Change position out of bounds: {change1}")
                continue
            
            # Check for overlaps with other changes
            for change2 in sorted_changes[i+1:]:
                pos2 = change2['position']
                len2 = change2['length']
                # Check if ranges overlap
                if not (pos1 + len1 <= pos2 or pos2 + len2 <= pos1):
                    print(f"Warning: Overlapping changes detected: {change1} and {change2}")
        
        # Apply changes in reverse order (from end to beginning)
        for change in sorted_changes:
            pos = change['position']
            length = change['length']
            
            # Skip if out of bounds
            if pos < 0 or pos + length > len(self.masked_text):
                continue
            
            # Verify we're replacing the expected original text
            expected_original = self.extracted_text[pos:pos+length]
            current_text = self.masked_text[pos:pos+length]
            
            # Only replace if we're still at the original text (not already masked)
            # This prevents double-replacement
            if current_text == expected_original:
                self.masked_text = (
                    self.masked_text[:pos] +
                    change['masked'] +
                    self.masked_text[pos + length:]
                )
    
    def update_changes_listbox(self):
        """Update the changes listbox with current changes - one entry per name"""
        self.changes_listbox.delete(0, tk.END)
        
        # Group changes by normalized name
        name_groups = {}
        for change in self.current_changes:
            norm_name = change.get('normalized_name', '')
            if norm_name not in name_groups:
                name_groups[norm_name] = []
            name_groups[norm_name].append(change)
        
        # Display one entry per name with occurrence count
        for norm_name, changes in sorted(name_groups.items(), key=lambda x: self.name_to_id.get(x[0], {}).get('id', 0)):
            if norm_name in self.name_to_id:
                name_info = self.name_to_id[norm_name]
                original_name = name_info['original_name']
                masked = name_info['masked']
                count = len(changes)
                display_text = f"{original_name} → {masked} ({count} occurrence{'s' if count != 1 else ''})"
                self.changes_listbox.insert(tk.END, display_text)
    
    def undo_change(self, event=None):
        """Undo a change when double-clicked"""
        selection = self.changes_listbox.curselection()
        if selection:
            self.undo_selected_change()
    
    def undo_selected_change(self):
        """Undo the selected change - removes all occurrences of the selected name"""
        selection = self.changes_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a change to undo.")
            return
        
        # Get the selected index
        display_index = selection[0]
        
        # Get the normalized name from the listbox entry
        # Build the same grouping as in update_changes_listbox
        name_groups = {}
        for change in self.current_changes:
            norm_name = change.get('normalized_name', '')
            if norm_name not in name_groups:
                name_groups[norm_name] = []
            name_groups[norm_name].append(change)
        
        # Get the normalized name at the selected index
        sorted_names = sorted(name_groups.items(), key=lambda x: self.name_to_id.get(x[0], {}).get('id', 0))
        if display_index >= len(sorted_names):
            return
        
        norm_name_to_undo, changes_to_remove = sorted_names[display_index]
        
        # Remove all occurrences of this name from current_changes
        self.current_changes = [c for c in self.current_changes if c.get('normalized_name') != norm_name_to_undo]
        self.masking_changes = [c for c in self.masking_changes if c.get('normalized_name') != norm_name_to_undo]
        
        # Remove from name mappings
        if norm_name_to_undo in self.name_to_id:
            del self.name_to_id[norm_name_to_undo]
        if norm_name_to_undo in self.name_occurrences:
            del self.name_occurrences[norm_name_to_undo]
        
        # Reassign IDs to be sequential
        sorted_names = sorted(self.name_to_id.items(), key=lambda x: x[1]['id'])
        for idx, (norm_name, name_info) in enumerate(sorted_names, 1):
            old_id = name_info['id']
            new_id = idx
            new_masked = f"[NAME_{new_id}]"
            
            # Update the ID and masked value
            name_info['id'] = new_id
            name_info['masked'] = new_masked
            
            # Update all changes for this name
            for change in self.current_changes:
                if change.get('normalized_name') == norm_name:
                    change['masked'] = new_masked
        
        # Rebuild masked text from scratch with remaining changes
        self.rebuild_masked_text()
        
        # Update preview
        self.masking_preview_area.delete(1.0, tk.END)
        self.masking_preview_area.insert(1.0, self.masked_text)
        
        # Update changes listbox
        self.update_changes_listbox()
        
    
    def send_to_api(self):
        """Send masked text to Claude API (initial request)"""
        if not self.masked_text:
            messagebox.showwarning("Warning", "Please extract and mask text first.")
            return
        
        if not self.selected_model:
            messagebox.showerror("Error", "No model selected. Please select a model from the dropdown.")
            return
        
        # Get instructions from text area
        instructions = self.instructions_text_area.get(1.0, tk.END).strip()
        if not instructions:
            instructions = "Fais un récit chronologique de ce rapport d'expertise medicale. Utilise le discour rapporté. Garde une connotation technique. Fais un récit continu."
        
        # Clear conversation history for new request
        self.conversation_history = []
        self.is_first_message = True
        
        # Prepare the initial prompt
        prompt = f"{instructions}\n\nText:\n{self.masked_text}"
        
        # Send the message
        self._send_api_message(prompt, is_first=True)
    
    def send_chat_message(self):
        """Send a follow-up message in the chat conversation"""
        if not self.masked_text:
            messagebox.showwarning("Warning", "Please extract and mask text first.")
            return
        
        if not self.selected_model:
            messagebox.showerror("Error", "No model selected. Please select a model from the dropdown.")
            return
        
        # Get chat input
        chat_message = self.chat_input.get().strip()
        if not chat_message:
            messagebox.showwarning("Warning", "Please enter a message.")
            return
        
        # Clear the input field
        self.chat_input.delete(0, tk.END)
        
        # Clear final text area before new chat request
        self.final_text_area.delete(1.0, tk.END)
        
        # Send the message
        self._send_api_message(chat_message, is_first=False)
    
    def _send_api_message(self, user_message: str, is_first: bool = False):
        """Internal method to send message to LLM API and handle response"""
        try:
            # Validate model selection
            if not self.selected_model:
                messagebox.showerror("Error", "No model selected. Please select a model from the dropdown.")
                return
            
            # Get provider for selected model
            provider = self.llm_registry.get_provider_for_model(self.selected_model)
            if not provider:
                messagebox.showerror("Error", f"No provider found for model: {self.selected_model}")
                return
            
            # Show processing message
            # Note: Final text area is already cleared in send_to_api or send_chat_message
            self.final_text_area.insert(tk.END, "Processing... Please wait.")
            self.root.update()
            
            # Add user message to conversation history
            self.conversation_history.append({
                "role": "user",
                "content": user_message
            })
            
            # Call LLM API with full conversation history using abstraction layer
            response_text = provider.send_message(
                messages=self.conversation_history,
                model=self.selected_model,
                max_tokens=64000
            )
            
            # Add assistant response to conversation history
            self.conversation_history.append({
                "role": "assistant",
                "content": response_text
            })
            
            # Restore masked names in response
            restored_text = response_text
            for change in reversed(self.current_changes):
                restored_text = restored_text.replace(change['masked'], change['original'])
            
            # Add indentation (1 tab) at the beginning of every paragraph
            paragraphs = restored_text.split('\n')
            indented_paragraphs = []
            for para in paragraphs:
                if para.strip():  # Only indent non-empty paragraphs
                    indented_paragraphs.append('\t' + para)
                else:
                    indented_paragraphs.append(para)  # Keep empty lines as-is
            indented_text = '\n'.join(indented_paragraphs)
            
            # Replace "Processing..." message with response
            content = self.final_text_area.get(1.0, tk.END)
            if "Processing... Please wait." in content:
                self.final_text_area.delete(1.0, tk.END)
                self.final_text_area.insert(1.0, indented_text)
            else:
                # Fallback: just replace all content
                self.final_text_area.delete(1.0, tk.END)
                self.final_text_area.insert(1.0, indented_text)
            
            # Scroll to bottom
            self.final_text_area.see(tk.END)
            
            self.is_first_message = False
            model_display = self.llm_registry.get_model_display_name(self.selected_model)
            
        except Exception as e:
            model_display = self.llm_registry.get_model_display_name(self.selected_model) if self.selected_model else "LLM"
            messagebox.showerror("Error", f"Failed to process message with {model_display}: {str(e)}")
            # Remove processing message and show error
            content = self.final_text_area.get(1.0, tk.END)
            if "Processing... Please wait." in content:
                self.final_text_area.delete(1.0, tk.END)
            self.final_text_area.insert(1.0, f"Error: {str(e)}")
    
    def clear_conversation_history(self):
        """Clear the conversation history"""
        if not self.conversation_history:
            return
        
        if messagebox.askyesno("Confirm", "Clear conversation history? This will reset the chat."):
            self.conversation_history = []
            self.is_first_message = True
            self.final_text_area.delete(1.0, tk.END)
    
    def copy_final_text(self):
        """Copy final text to clipboard"""
        final_text = self.final_text_area.get(1.0, tk.END).strip()
        if not final_text:
            messagebox.showwarning("Warning", "No text to copy.")
            return
        
        self.root.clipboard_clear()
        self.root.clipboard_append(final_text)
    
    def load_chat_messages(self):
        """Load saved chat messages from chat.txt file"""
        try:
            if os.path.exists(self.chat_file):
                with open(self.chat_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    for line in lines:
                        line = line.strip()
                        if not line:
                            continue
                        # Parse format: "label" :: "text"
                        if ' :: ' in line:
                            parts = line.split(' :: ', 1)
                            if len(parts) == 2:
                                label = parts[0].strip().strip('"')
                                text = parts[1].strip().strip('"')
                                # Handle escaped quotes and newlines
                                text = text.replace('\\n', '\n').replace('\\"', '"')
                                self.chat_dict[label] = text
            else:
                # Initialize with default "basic" chat message
                self.chat_dict = {"basic": ""}
                self.save_chat_messages()
        except Exception as e:
            print(f"Error loading chat messages: {e}")
            # Initialize with default "basic" chat message
            self.chat_dict = {"basic": ""}
            self.save_chat_messages()
    
    def save_chat_messages(self):
        """Save chat messages to chat.txt file"""
        try:
            with open(self.chat_file, 'w', encoding='utf-8') as f:
                for label, text in sorted(self.chat_dict.items()):
                    # Escape quotes and newlines for storage
                    escaped_text = text.replace('"', '\\"').replace('\n', '\\n')
                    f.write(f'"{label}" :: "{escaped_text}"\n')
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save chat messages: {str(e)}")
    
    def update_chat_combo(self):
        """Update the chat label combobox with current labels"""
        labels = sorted(self.chat_dict.keys())
        self.chat_label_combo['values'] = labels
        if labels:
            # Set current label if available
            if self.current_chat_label in labels:
                self.chat_label_var.set(self.current_chat_label)
            else:
                self.chat_label_var.set(labels[0])
                self.current_chat_label = labels[0]
    
    def on_chat_label_selected(self, event=None):
        """Handle chat label selection from combobox"""
        selected_label = self.chat_label_var.get()
        if selected_label and selected_label in self.chat_dict:
            chat_text = self.chat_dict[selected_label]
            self.chat_input.delete(0, tk.END)
            self.chat_input.insert(0, chat_text)
            self.current_chat_label = selected_label
    
    def save_chat_message(self):
        """Save current chat input to the selected label"""
        selected_label = self.chat_label_var.get()
        if not selected_label:
            messagebox.showwarning("Warning", "Please select a label.")
            return
        
        # Get current text from chat input
        chat_text = self.chat_input.get().strip()
        
        # Save to dictionary
        self.chat_dict[selected_label] = chat_text
        self.save_chat_messages()
        self.current_chat_label = selected_label
    
    def create_new_chat_message(self):
        """Create a new chat message label with blank text"""
        new_label = simpledialog.askstring("Create New Chat Message", "Enter label name:")
        if not new_label or not new_label.strip():
            return
        
        new_label = new_label.strip()
        
        # Check if label already exists
        if new_label in self.chat_dict:
            messagebox.showwarning("Warning", f"Label '{new_label}' already exists.")
            return
        
        # Create new chat message with blank text
        self.chat_dict[new_label] = ""
        self.save_chat_messages()
        
        # Update combo and select new label
        self.update_chat_combo()
        self.chat_label_var.set(new_label)
        self.on_chat_label_selected()
        
    
    def delete_chat_message(self):
        """Delete the selected chat message label with warning"""
        selected_label = self.chat_label_var.get()
        if not selected_label:
            messagebox.showwarning("Warning", "Please select a label to delete.")
            return
        
        # Prevent deleting "basic" if it's the only one
        if selected_label == "basic" and len(self.chat_dict) == 1:
            messagebox.showwarning("Warning", "Cannot delete the 'basic' chat message. At least one message must exist.")
            return
        
        # Show warning dialog
        if not messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete the chat message label '{selected_label}'?\n\nThis action cannot be undone."):
            return
        
        # Delete the chat message
        del self.chat_dict[selected_label]
        self.save_chat_messages()
        
        # Update combo and select another label
        self.update_chat_combo()
        if self.chat_dict:
            # Select first available label
            first_label = sorted(self.chat_dict.keys())[0]
            self.chat_label_var.set(first_label)
            self.on_chat_label_selected()
        
    
    def load_instructions(self):
        """Load saved instructions from instructions.txt file"""
        try:
            if os.path.exists(self.instructions_file):
                with open(self.instructions_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    for line in lines:
                        line = line.strip()
                        if not line:
                            continue
                        # Parse format: "label" :: "text"
                        if ' :: ' in line:
                            parts = line.split(' :: ', 1)
                            if len(parts) == 2:
                                label = parts[0].strip().strip('"')
                                text = parts[1].strip().strip('"')
                                # Handle escaped quotes and newlines
                                text = text.replace('\\n', '\n').replace('\\"', '"')
                                self.instructions_dict[label] = text
            else:
                # Initialize with default "basic" instruction
                default_text = "Fais un récit chronologique de ce rapport d'expertise medicale. Utilise le discour rapporté. Garde une connotation technique. Fais un récit continu."
                self.instructions_dict = {"basic": default_text}
                self.save_instructions()
        except Exception as e:
            print(f"Error loading instructions: {e}")
            # Initialize with default "basic" instruction
            default_text = "Fais un récit chronologique de ce rapport d'expertise medicale. Utilise le discour rapporté. Garde une connotation technique. Fais un récit continu."
            self.instructions_dict = {"basic": default_text}
            self.save_instructions()
    
    def save_instructions(self):
        """Save instructions to instructions.txt file"""
        try:
            with open(self.instructions_file, 'w', encoding='utf-8') as f:
                for label, text in sorted(self.instructions_dict.items()):
                    # Escape quotes and newlines for storage
                    escaped_text = text.replace('"', '\\"').replace('\n', '\\n')
                    f.write(f'"{label}" :: "{escaped_text}"\n')
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save instructions: {str(e)}")
    
    def update_model_combo(self):
        """Update the model selection dropdown with available models"""
        available_models = self.llm_registry.get_all_models()
        if not available_models:
            self.model_combo['values'] = []
            self.model_var.set("")
            self.selected_model = None
            return
        
        # Create display list with simplified model names only
        display_values = []
        model_display_map = {}  # Map display name to model ID
        for model in available_models:
            display_name = self.llm_registry.get_model_display_name(model)
            display_values.append(display_name)
            model_display_map[display_name] = model
        
        self.model_combo['values'] = display_values
        self.model_display_map = model_display_map  # Store mapping for selection
        
        # Set default selection if not already set
        if self.selected_model and self.selected_model in available_models:
            display_name = self.llm_registry.get_model_display_name(self.selected_model)
            self.model_var.set(display_name)
        else:
            # Select first model
            self.selected_model = available_models[0]
            display_name = self.llm_registry.get_model_display_name(self.selected_model)
            self.model_var.set(display_name)
    
    def on_model_selected(self, event=None):
        """Handle model selection from dropdown"""
        selection = self.model_var.get()
        if not selection:
            return
        
        # Get model ID from display name mapping
        if hasattr(self, 'model_display_map') and selection in self.model_display_map:
            self.selected_model = self.model_display_map[selection]
    
    def update_instruction_combo(self):
        """Update the instruction label combobox with current labels"""
        labels = sorted(self.instructions_dict.keys())
        self.instruction_label_combo['values'] = labels
        if labels:
            # Set current label if available
            if self.current_instruction_label in labels:
                self.instruction_label_var.set(self.current_instruction_label)
            else:
                self.instruction_label_var.set(labels[0])
                self.current_instruction_label = labels[0]
    
    def on_instruction_label_selected(self, event=None):
        """Handle instruction label selection from combobox"""
        selected_label = self.instruction_label_var.get()
        if selected_label and selected_label in self.instructions_dict:
            instruction_text = self.instructions_dict[selected_label]
            self.instructions_text_area.delete(1.0, tk.END)
            self.instructions_text_area.insert(1.0, instruction_text)
            self.current_instruction_label = selected_label
    
    def save_instruction(self):
        """Save current instruction text to the selected label"""
        selected_label = self.instruction_label_var.get()
        if not selected_label:
            messagebox.showwarning("Warning", "Please select a label.")
            return
        
        # Get current text from text area
        instruction_text = self.instructions_text_area.get(1.0, tk.END).strip()
        
        # Save to dictionary
        self.instructions_dict[selected_label] = instruction_text
        self.save_instructions()
        self.current_instruction_label = selected_label
    
    def create_new_instruction(self):
        """Create a new instruction label with blank text"""
        new_label = simpledialog.askstring("Create New Instruction", "Enter label name:")
        if not new_label or not new_label.strip():
            return
        
        new_label = new_label.strip()
        
        # Check if label already exists
        if new_label in self.instructions_dict:
            messagebox.showwarning("Warning", f"Label '{new_label}' already exists.")
            return
        
        # Create new instruction with blank text
        self.instructions_dict[new_label] = ""
        self.save_instructions()
        
        # Update combo and select new label
        self.update_instruction_combo()
        self.instruction_label_var.set(new_label)
        self.on_instruction_label_selected()
        
    
    def delete_instruction(self):
        """Delete the selected instruction label with warning"""
        selected_label = self.instruction_label_var.get()
        if not selected_label:
            messagebox.showwarning("Warning", "Please select a label to delete.")
            return
        
        # Prevent deleting "basic" if it's the only one
        if selected_label == "basic" and len(self.instructions_dict) == 1:
            messagebox.showwarning("Warning", "Cannot delete the 'basic' instruction. At least one instruction must exist.")
            return
        
        # Show warning dialog
        if not messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete the instruction label '{selected_label}'?\n\nThis action cannot be undone."):
            return
        
        # Delete the instruction
        del self.instructions_dict[selected_label]
        self.save_instructions()
        
        # Update combo and select another label
        self.update_instruction_combo()
        if self.instructions_dict:
            # Select first available label
            first_label = sorted(self.instructions_dict.keys())[0]
            self.instruction_label_var.set(first_label)
            self.on_instruction_label_selected()
        


def main():
    # Use TkinterDnD if available, otherwise use regular Tk
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = WordProcessorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

