import pandas as pd
import os
import time
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread

# Check if deep_translator is installed, if not guide the user to install it
try:
    from deep_translator import GoogleTranslator
except ImportError:
    print("The deep_translator package is not installed.")
    print("Please install it by running: pip install deep_translator")
    input("Press Enter to exit...")
    sys.exit(1)

# Dictionary mapping user-friendly language names to language codes
LANGUAGE_MAP = {
    "English": "en",
    "French": "fr",
    "Spanish": "es",
    "German": "de",
    "Italian": "it",
    "Dutch": "nl",
    "Portuguese": "pt",
    "Russian": "ru",
    "Chinese": "zh-CN",
    "Japanese": "ja",
    "Korean": "ko",
    "Arabic": "ar",
    "Hindi": "hi",
    "Turkish": "tr",
    "Greek": "el",
    "Polish": "pl",
    "Vietnamese": "vi",
    "Thai": "th",
    "Swedish": "sv",
    "Danish": "da",
    "Finnish": "fi",
    "Norwegian": "no",
    "Czech": "cs",
    "Romanian": "ro",
    "Hungarian": "hu",
    "Bulgarian": "bg",
    "Ukrainian": "uk",
    "Croatian": "hr",
    "Slovak": "sk",
    "Indonesian": "id",
    "Malay": "ms",
    "Hebrew": "he"
}

class ExcelTranslatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Translator")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # Set minimum window size
        self.root.minsize(500, 400)
        
        # Variables
        self.input_file_var = tk.StringVar()
        self.output_location_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        self.progress_var = tk.DoubleVar(value=0)
        self.languages = list(LANGUAGE_MAP.keys())
        self.selected_languages = []
        
        # Create the GUI
        self.create_widgets()
        
        # Center the window
        self.center_window()
    
    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_widgets(self):
        # Create main frame with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Input file row
        input_file_frame = ttk.Frame(file_frame)
        input_file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_file_frame, text="Input Excel File:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(input_file_frame, textvariable=self.input_file_var, width=40).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(input_file_frame, text="Browse...", command=self.browse_input_file).pack(side=tk.LEFT, padx=5)
        
        # Output location row
        output_frame = ttk.Frame(file_frame)
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_frame, text="Output Location:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(output_frame, textvariable=self.output_location_var, width=40).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(output_frame, text="Browse...", command=self.browse_output_location).pack(side=tk.LEFT, padx=5)
        
        # Language selection section
        lang_frame = ttk.LabelFrame(main_frame, text="Languages", padding="10")
        lang_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Split into available and selected languages
        lang_paned = ttk.PanedWindow(lang_frame, orient=tk.HORIZONTAL)
        lang_paned.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Available languages frame
        available_frame = ttk.Frame(lang_paned)
        lang_paned.add(available_frame, weight=1)
        
        ttk.Label(available_frame, text="Available Languages:").pack(anchor=tk.W, padx=5, pady=2)
        
        # Listbox with scrollbar for available languages
        avail_scroll = ttk.Scrollbar(available_frame)
        avail_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.available_listbox = tk.Listbox(available_frame, yscrollcommand=avail_scroll.set, selectmode=tk.EXTENDED)
        self.available_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        avail_scroll.config(command=self.available_listbox.yview)
        
        # Populate available languages
        for lang in sorted(self.languages):
            self.available_listbox.insert(tk.END, lang)
        
        # Buttons frame
        btn_frame = ttk.Frame(lang_paned)
        lang_paned.add(btn_frame, weight=0)
        
        ttk.Button(btn_frame, text="→", command=self.add_languages).pack(pady=10)
        ttk.Button(btn_frame, text="←", command=self.remove_languages).pack(pady=10)
        
        # Selected languages frame
        selected_frame = ttk.Frame(lang_paned)
        lang_paned.add(selected_frame, weight=1)
        
        ttk.Label(selected_frame, text="Selected Languages:").pack(anchor=tk.W, padx=5, pady=2)
        
        # Listbox with scrollbar for selected languages
        sel_scroll = ttk.Scrollbar(selected_frame)
        sel_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.selected_listbox = tk.Listbox(selected_frame, yscrollcommand=sel_scroll.set, selectmode=tk.EXTENDED)
        self.selected_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        sel_scroll.config(command=self.selected_listbox.yview)
        
        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(button_frame, text="Translate", command=self.start_translation).pack(side=tk.RIGHT, padx=5)
        
        # Status bar and progress bar
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(status_frame, textvariable=self.status_var).pack(anchor=tk.W, padx=5)
    
    def browse_input_file(self):
        filetypes = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        filename = filedialog.askopenfilename(title="Select Excel File", filetypes=filetypes)
        if filename:
            self.input_file_var.set(filename)
            
            # Auto-set output location to same directory
            output_dir = os.path.dirname(filename)
            self.output_location_var.set(output_dir)
    
    def browse_output_location(self):
        directory = filedialog.askdirectory(title="Select Output Location")
        if directory:
            self.output_location_var.set(directory)
    
    def add_languages(self):
        # Get selected indices from available list
        selected_indices = self.available_listbox.curselection()
        if not selected_indices:
            return
        
        # Get selected languages and add to selected list
        for index in selected_indices:
            lang = self.available_listbox.get(index)
            if lang not in self.selected_languages:
                self.selected_languages.append(lang)
                self.selected_listbox.insert(tk.END, lang)
        
        # Update listboxes
        self.update_listboxes()
    
    def remove_languages(self):
        # Get selected indices from selected list
        selected_indices = self.selected_listbox.curselection()
        if not selected_indices:
            return
        
        # Get selected languages and remove from selected list
        langs_to_remove = [self.selected_listbox.get(index) for index in selected_indices]
        for lang in langs_to_remove:
            if lang in self.selected_languages:
                self.selected_languages.remove(lang)
        
        # Update listboxes
        self.update_listboxes()
    
    def update_listboxes(self):
        # Clear and repopulate selected listbox
        self.selected_listbox.delete(0, tk.END)
        for lang in self.selected_languages:
            self.selected_listbox.insert(tk.END, lang)
        
        # Clear and repopulate available listbox
        self.available_listbox.delete(0, tk.END)
        for lang in sorted(self.languages):
            if lang not in self.selected_languages:
                self.available_listbox.insert(tk.END, lang)
    
    def start_translation(self):
        # Get input values
        input_file = self.input_file_var.get().strip()
        output_location = self.output_location_var.get().strip()
        selected_languages = self.selected_languages
        
        # Validate inputs
        if not input_file:
            messagebox.showerror("Error", "Please select an input Excel file.")
            return
        
        if not os.path.isfile(input_file):
            messagebox.showerror("Error", "Input file does not exist.")
            return
        
        if not selected_languages:
            messagebox.showerror("Error", "Please select at least one language for translation.")
            return
        
        # Start translation in a separate thread
        self.progress_var.set(0)
        self.status_var.set("Initializing translation...")
        
        # Convert language names to language codes
        language_codes = [LANGUAGE_MAP[lang] for lang in selected_languages]
        
        # Start translation in a separate thread
        translation_thread = Thread(target=self.translate_excel, 
                                   args=(input_file, language_codes, output_location))
        translation_thread.daemon = True
        translation_thread.start()
    
    def translate_excel(self, input_file, target_languages, output_location):
        try:
            # Read the Excel file
            self.update_status(f"Reading Excel file: {input_file}")
            try:
                df = pd.read_excel(input_file)
            except Exception as e:
                self.show_error(f"Failed to read Excel file: {e}")
                return
            
            # Get the base filename without extension
            base_filename = os.path.splitext(os.path.basename(input_file))[0]
            
            # Create output folder if it doesn't exist
            try:
                if not os.path.exists(output_location):
                    os.makedirs(output_location)
                    self.update_status(f"Created output directory: {output_location}")
            except Exception as e:
                self.show_error(f"Could not create output directory: {e}")
                return
            
            # Process each target language
            total_languages = len(target_languages)
            
            for lang_index, lang_code in enumerate(target_languages):
                # Get language name from code
                lang_name = next((name for name, code in LANGUAGE_MAP.items() if code == lang_code), lang_code)
                
                self.update_status(f"Translating to {lang_name} ({lang_code}) - Language {lang_index+1}/{total_languages}")
                
                # Create a copy of the original dataframe
                translated_df = df.copy()
                
                try:
                    # Initialize translator
                    translator = GoogleTranslator(source='auto', target=lang_code)
                    
                    # Track progress
                    total_cells = 0
                    translated_cells = 0
                    
                    # Identify translatable columns first
                    translatable_columns = []
                    for col in translated_df.columns:
                        if translated_df[col].dtype == 'object':
                            translatable_columns.append(col)
                    
                    if not translatable_columns:
                        self.update_status(f"No text columns found to translate for {lang_name}.")
                        # Move to next language
                        continue
                    
                    # Count total cells to translate for progress reporting
                    for col in translatable_columns:
                        for idx in range(len(translated_df)):
                            cell_value = translated_df.at[idx, col]
                            if isinstance(cell_value, str) and cell_value.strip():
                                total_cells += 1
                    
                    # Translate each string column in the dataframe
                    for col_index, col in enumerate(translatable_columns):
                        col_progress = (col_index / len(translatable_columns)) * 100
                        self.update_status(f"Translating column: {col} ({col_index+1}/{len(translatable_columns)})")
                        
                        # Translate each cell in the column
                        for idx in range(len(translated_df)):
                            cell_value = translated_df.at[idx, col]
                            
                            # Only translate string values that are not empty
                            if isinstance(cell_value, str) and cell_value.strip():
                                try:
                                    # Translate the text
                                    translated_text = translator.translate(cell_value)
                                    translated_df.at[idx, col] = translated_text
                                    translated_cells += 1
                                    
                                    # Update progress
                                    if total_cells > 0:
                                        overall_progress = (
                                            (lang_index / total_languages) * 100 + 
                                            (1 / total_languages) * (translated_cells / total_cells) * 100
                                        )
                                        self.progress_var.set(overall_progress)
                                        
                                        cell_progress = (translated_cells / total_cells) * 100
                                        self.update_status(
                                            f"Translating to {lang_name}: {translated_cells}/{total_cells} cells ({cell_progress:.1f}%)"
                                        )
                                    
                                    # Add a small delay to avoid hitting rate limits
                                    time.sleep(0.2)
                                except Exception as e:
                                    self.update_status(f"Error translating '{cell_value}': {str(e)[:100]}...")
                    
                    # Save the translated dataframe to a new Excel file
                    output_file = os.path.join(output_location, f"{base_filename}_{lang_code}.xlsx")
                    try:
                        translated_df.to_excel(output_file, index=False)
                        self.update_status(f"Saved translated file: {output_file}")
                    except Exception as e:
                        self.show_error(f"Error saving file {output_file}: {e}")
                
                except Exception as e:
                    self.show_error(f"Error during translation to {lang_name}: {e}")
            
            # Complete
            self.progress_var.set(100)
            self.update_status("Translation completed!")
            messagebox.showinfo("Success", "Translation completed successfully!")
        
        except Exception as e:
            self.show_error(f"Unexpected error: {e}\n{traceback.format_exc()}")
    
    def update_status(self, message):
        # Update status message in the main thread
        self.root.after(0, lambda: self.status_var.set(message))
        print(message)
    
    def show_error(self, message):
        # Show error message in the main thread
        print(f"ERROR: {message}")
        self.root.after(0, lambda: self.status_var.set(f"Error: {message}"))
        self.root.after(0, lambda: messagebox.showerror("Error", message))

def check_dependencies():
    """Check if required packages are installed"""
    missing_packages = []
    
    try:
        import pandas
    except ImportError:
        missing_packages.append("pandas")
    
    try:
        import deep_translator
    except ImportError:
        missing_packages.append("deep_translator")
    
    try:
        import openpyxl
    except ImportError:
        missing_packages.append("openpyxl")
    
    if missing_packages:
        message = "The following required packages are missing:\n"
        message += "\n".join(missing_packages)
        message += "\n\nWould you like to install them now?"
        
        # Use raw tkinter for this as we need it before the main app starts
        root = tk.Tk()
        root.withdraw()
        install = messagebox.askyesno("Missing Dependencies", message)
        root.destroy()
        
        if install:
            print("Installing missing packages...")
            try:
                import subprocess
                for package in missing_packages:
                    print(f"Installing {package}...")
                    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                print("All packages installed successfully.")
                return True
            except Exception as e:
                print(f"Error installing packages: {e}")
                message = f"Error installing packages. Please install manually:\n"
                message += "\n".join([f"pip install {pkg}" for pkg in missing_packages])
                
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror("Installation Error", message)
                root.destroy()
                return False
        else:
            return False
    
    return True

def main():
    # Check dependencies first
    if not check_dependencies():
        sys.exit(1)
    
    # Create and run the application
    root = tk.Tk()
    app = ExcelTranslatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
