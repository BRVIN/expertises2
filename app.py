import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import docx
from anthropic import Anthropic
import re
from typing import List, Tuple, Optional
import unicodedata


class WordProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word Document Processor with Claude API")
        self.root.geometry("1200x900")
        
        # API Configuration
        self.api_key = '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'
        self.client = None
        if self.api_key:
            self.client = Anthropic(api_key=self.api_key)
        
        # Development/Testing: Hardcoded file path (set to None to disable, or provide full path)
        # Example: self.hardcoded_file_path = r"C:\Users\bob\Documents\test_document.docx"
        self.hardcoded_file_path = r"C:\Users\bob\Music\test.docx"  # Set this to your test file path for development
        
        # Data storage
        self.full_text = ""
        self.extracted_text = ""
        self.masked_text = ""
        self.masking_changes = []  # List of individual occurrence changes
        self.current_changes = []  # Current list of changes for undo
        self.name_to_id = {}  # Maps normalized name to its unique ID
        self.name_occurrences = {}  # Maps normalized name to list of occurrences
        
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
        main_frame.columnconfigure(1, weight=1)
        
        # File upload section
        ttk.Label(main_frame, text="Word Document:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.file_path_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.file_path_var, width=50, state="readonly").grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_file).grid(row=0, column=2, padx=5)
        
        # Extraction range section
        ttk.Label(main_frame, text="Start Word (wordA):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.start_word_var = tk.StringVar(value="commemoratifs")
        ttk.Entry(main_frame, textvariable=self.start_word_var, width=30).grid(row=1, column=1, sticky=tk.W, padx=5)
        
        ttk.Label(main_frame, text="End Word (wordB):").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.end_word_var = tk.StringVar(value="certificat")
        ttk.Entry(main_frame, textvariable=self.end_word_var, width=30).grid(row=2, column=1, sticky=tk.W, padx=5)
        
        # Button frame for extraction buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=2, padx=5, sticky=tk.W)
        ttk.Button(button_frame, text="Extract Text", command=self.extract_text).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Undo Extraction", command=self.undo_extraction).pack(side=tk.LEFT, padx=2)
        
        # Extracted text display
        ttk.Label(main_frame, text="Extracted Text:").grid(row=3, column=0, sticky=(tk.W, tk.N), pady=5)
        self.extracted_text_area = scrolledtext.ScrolledText(main_frame, height=8, width=80, wrap=tk.WORD)
        self.extracted_text_area.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # Names to mask section
        ttk.Label(main_frame, text="Names/Surnames to Mask (comma-separated):").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.names_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.names_var, width=50).grid(row=4, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(main_frame, text="Apply Masking", command=self.apply_masking).grid(row=4, column=2, padx=5)
        
        # Masking preview
        ttk.Label(main_frame, text="Masking Preview:").grid(row=5, column=0, sticky=(tk.W, tk.N), pady=5)
        self.masking_preview_area = scrolledtext.ScrolledText(main_frame, height=8, width=80, wrap=tk.WORD)
        self.masking_preview_area.grid(row=5, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # Changes list
        ttk.Label(main_frame, text="Changes (click to undo):").grid(row=6, column=0, sticky=(tk.W, tk.N), pady=5)
        self.changes_listbox = tk.Listbox(main_frame, height=5, width=50)
        self.changes_listbox.grid(row=6, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.changes_listbox.bind('<Double-Button-1>', self.undo_change)
        ttk.Button(main_frame, text="Undo Selected", command=self.undo_selected_change).grid(row=6, column=2, padx=5, sticky=tk.N)
        
        # API instructions
        ttk.Label(main_frame, text="Claude API Instructions:").grid(row=7, column=0, sticky=tk.W, pady=5)
        self.instructions_var = tk.StringVar(value="summarize this text")
        ttk.Entry(main_frame, textvariable=self.instructions_var, width=50).grid(row=7, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(main_frame, text="Send to Claude API", command=self.send_to_api).grid(row=7, column=2, padx=5)
        
        # Final text display
        ttk.Label(main_frame, text="Final Text (from Claude):").grid(row=8, column=0, sticky=(tk.W, tk.N), pady=5)
        self.final_text_area = scrolledtext.ScrolledText(main_frame, height=8, width=80, wrap=tk.WORD)
        self.final_text_area.grid(row=8, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        ttk.Button(main_frame, text="Copy Final Text", command=self.copy_final_text).grid(row=8, column=2, padx=5, sticky=tk.S)
        
        # Configure grid weights for resizing
        main_frame.rowconfigure(3, weight=1)
        main_frame.rowconfigure(5, weight=1)
        main_frame.rowconfigure(8, weight=1)
        
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
            messagebox.showinfo("Success", f"Document loaded successfully!\nTotal characters: {len(self.full_text)}")
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
        
        end_pos = start_pos + end_pos_relative + len(end_word)
        
        # Extract text
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
        
        messagebox.showinfo("Success", f"Text extracted successfully!\nCharacters: {len(self.extracted_text)}")
    
    def undo_extraction(self):
        """Undo text extraction - clears extracted text and all masking"""
        if not self.extracted_text:
            messagebox.showinfo("Info", "No extracted text to undo.")
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
        
        messagebox.showinfo("Success", "Extraction undone. All text and masking cleared.")
    
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
            messagebox.showinfo("Info", "No new occurrences found to mask.")
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
        
        messagebox.showinfo("Success", f"Masking applied to {total_occurrences} occurrence(s) of {len([n for n in names if self.normalize_text(n) in self.name_to_id])} name(s).")
    
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
        
        messagebox.showinfo("Success", f"All occurrences of the selected name have been undone.")
    
    def send_to_api(self):
        """Send masked text to Claude API"""
        if not self.masked_text:
            messagebox.showwarning("Warning", "Please extract and mask text first.")
            return
        
        if not self.client:
            messagebox.showerror("Error", "API key not configured. Please set your API key in the code.")
            return
        
        instructions = self.instructions_var.get().strip()
        if not instructions:
            instructions = "summarize this text"
        
        try:
            # Show processing message
            self.final_text_area.delete(1.0, tk.END)
            self.final_text_area.insert(1.0, "Processing... Please wait.")
            self.root.update()
            
            # Prepare the prompt
            prompt = f"{instructions}\n\nText:\n{self.masked_text}"
            
            # Call Claude API (Opus 4.1)
            message = self.client.messages.create(
                model="claude-opus-4-5-20251101",
                max_tokens=4096,
                messages=[
                    {
                        "role": "user",
                        "content": prompt
                    }
                ]
            )
            
            # Get response
            response_text = message.content[0].text if message.content else ""
            
            # Restore masked names
            restored_text = response_text
            for change in reversed(self.current_changes):
                restored_text = restored_text.replace(change['masked'], change['original'])
            
            # Display final text
            self.final_text_area.delete(1.0, tk.END)
            self.final_text_area.insert(1.0, restored_text)
            
            messagebox.showinfo("Success", "Text processed successfully by Claude API!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process text with Claude API: {str(e)}")
            self.final_text_area.delete(1.0, tk.END)
            self.final_text_area.insert(1.0, f"Error: {str(e)}")
    
    def copy_final_text(self):
        """Copy final text to clipboard"""
        final_text = self.final_text_area.get(1.0, tk.END).strip()
        if not final_text:
            messagebox.showwarning("Warning", "No text to copy.")
            return
        
        self.root.clipboard_clear()
        self.root.clipboard_append(final_text)
        messagebox.showinfo("Success", "Text copied to clipboard!")


def main():
    root = tk.Tk()
    app = WordProcessorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

