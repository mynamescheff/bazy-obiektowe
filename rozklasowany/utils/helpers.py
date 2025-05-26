import tkinter as tk
from tkinter import filedialog
import re
import os

class Helper:
    def browse_directory(self, entry_widget):
        directory = filedialog.askdirectory()
        if directory:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, directory)

    def transform_to_swift_accepted_characters(self, text_list):
        """Basic character transformation - simplified from what was referenced"""
        if not text_list:
            return []
            
        result = []
        for text in text_list:
            if text:
                # Remove potentially problematic characters
                clean_text = re.sub(r'[\/:*?"<>|\t]', ' ', str(text))
                result.append(clean_text)
            else:
                result.append("")
        return result
    
    def get_unique_filename(self, base_path, original_filename, extension):
        counter = 2
        new_filename = original_filename
        while os.path.exists(os.path.join(base_path, f"{new_filename}{extension}")):
            new_filename = f"{original_filename} {counter}"
            counter += 1
        return new_filename