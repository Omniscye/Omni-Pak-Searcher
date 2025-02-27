import os
import shutil
import subprocess
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import ttkbootstrap as tb
import zipfile
import random
import psutil
import threading
import pygame.mixer
import sys

# Program name and theme
PROGRAM_NAME = "Omni Pak Searcher"
THEME = "darkly"

# Mocking messages
STAGE_1_INSULTS = [
    "Oh, you actually picked a file? How original.",
    "Wow, you managed to click 'Browse'. Impressive.",
    "File selected. Let's see if you can handle the rest.",
    "You found a file? I'm shocked.",
    "File chosen. Don't get too excited, it's not a trophy.",
]

STAGE_2_SEARCH_INSULTS = [
    "Searching... or am I just pretending?",
    "Looking for your files... like looking for a brain cell.",
    "Searching... because you clearly can't do it yourself.",
    "Hold on, I'm searching... unlike you, who's just lost.",
    "Searching... this might take a while. Go grab a coffee.",
]

STAGE_2_EXTRACT_INSULTS = [
    "Extracting... in reverse, because why not?",
    "Extracting... just like I extract joy from your life.",
    "Extracting... this is gonna be fun. For me, not you.",
    "Extracting... because you can't handle it the normal way.",
    "Extracting... prepare for chaos.",
]

# Global variables
INSULTS_ENABLED = True
search_results = []  # Store search results for extraction

# Function to mock the user
def mock_user(messages):
    if not INSULTS_ENABLED:
        return
    messagebox.showinfo("Roasting You", random.choice(messages))

# Function to search inside .pak files (or folder of .paks) and store results
def search_in_pak(path, search_term, result_text):
    global search_results
    search_results = []  # Reset search results
    pak_files = []

    if os.path.isfile(path) and path.lower().endswith(".pak"):
        # Single .pak file
        pak_files = [path]
    elif os.path.isdir(path):
        # Folder containing .pak files
        for item in os.listdir(path):
            if item.lower().endswith(".pak"):
                pak_files.append(os.path.join(path, item))
        if not pak_files:
            result_text.insert(tk.END, "No .pak files found in the selected folder. Pathetic.\n")
            return
    else:
        result_text.insert(tk.END, "Error: Invalid path. Pick a .pak file or folder, genius.\n")
        return

    mock_user(STAGE_2_SEARCH_INSULTS)
    result_text.insert(tk.END, f"Searching for '{search_term}' in {len(pak_files)} .pak file(s)...\n")
    time.sleep(1)

    all_found_files = []
    for pak_file in pak_files:
        try:
            with zipfile.ZipFile(pak_file, 'r') as pak:
                found_files = [file for file in pak.namelist() if search_term.lower() in file.lower()]
                if found_files:
                    result_text.insert(tk.END, f"\nResults in '{os.path.basename(pak_file)}':\n")
                    for file in found_files:
                        result_text.insert(tk.END, f"- {file}\n")
                    all_found_files.extend([(pak_file, f) for f in found_files])  # Store with .pak path
        except Exception as e:
            result_text.insert(tk.END, f"Error searching {pak_file}: {e}. Maybe it’s laughing at you.\n")

    search_results = all_found_files  # Store as (pak_path, file_path) tuples
    if all_found_files:
        result_text.insert(tk.END, f"\nFound {len(all_found_files)} matching files across {len(pak_files)} .pak(s).\n")
        if messagebox.askyesno("Extract Found Files?", f"Do you want to extract these {len(all_found_files)} found files, preserving their directory structure?"):
            extract_search_results(result_text)
    else:
        result_text.insert(tk.END, "No files found. Maybe the files are mocking you too.\n")

# Function to extract all files from .pak with progress
def extract_from_pak(file_path, output_dir, result_text, progress_window, progress_bar, total_files):
    def do_extraction():
        try:
            # Initialize pygame mixer
            pygame.mixer.init()
            
            # Determine the correct path for loading_music.mp3
            if getattr(sys, 'frozen', False):  # Running as EXE
                base_path = sys._MEIPASS
            else:  # Running as script
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            mp3_path = os.path.join(base_path, "loading_music.mp3")
            
            # Debug the path and file existence
            result_text.insert(tk.END, f"Trying to load music from: {mp3_path}\n")
            result_text.see(tk.END)
            
            try:
                if os.path.exists(mp3_path):
                    pygame.mixer.music.load(mp3_path)
                    result_text.insert(tk.END, "Music loaded successfully, starting playback...\n")
                    result_text.see(tk.END)
                    pygame.mixer.music.play(-1)  # Loop indefinitely
                else:
                    result_text.insert(tk.END, f"Error: Music file 'loading_music.mp3' not found at {mp3_path}. No tunes for you.\n")
                    result_text.see(tk.END)
            except Exception as e:
                result_text.insert(tk.END, f"Error: Couldn’t play music - {e}. No tunes for you.\n")
                result_text.see(tk.END)

            with zipfile.ZipFile(file_path, 'r') as pak_file:
                files = pak_file.namelist()
                if not files:
                    result_text.insert(tk.END, "Error: The .pak file is empty. Shocking.\n")
                    progress_window.destroy()
                    return
                result_text.insert(tk.END, f"Found {len(files)} files to extract.\n")
                result_text.see(tk.END)
                total_files.set(len(files))
                
                for i, file in enumerate(files):
                    try:
                        dest = os.path.join(output_dir, file)
                        os.makedirs(os.path.dirname(dest), exist_ok=True)
                        result_text.insert(tk.END, f"Extracting: {file}\n")
                        result_text.see(tk.END)
                        with pak_file.open(file) as asset_file, open(dest, 'wb') as out_file:
                            shutil.copyfileobj(asset_file, out_file)
                        progress_bar['value'] = (i + 1) / total_files.get() * 100
                        progress_window.update_idletasks()
                        time.sleep(0.01)  # Small delay to ensure GUI updates
                    except Exception as e:
                        result_text.insert(tk.END, f"Error extracting {file}: {e}\n")
                        result_text.see(tk.END)
                        continue  # Skip this file and move on

                result_text.insert(tk.END, "\nExtraction complete. Check the folder, genius.\n")
                subprocess.Popen(f'explorer "{output_dir}"')
        except Exception as e:
            result_text.insert(tk.END, f"Critical error: {e}. You’re truly cursed.\n")
        finally:
            if pygame.mixer.get_init():
                pygame.mixer.music.stop()
            progress_window.destroy()

    # Validation checks
    if not file_path or not os.path.exists(file_path):
        result_text.insert(tk.END, "Error: No valid .pak file selected. What a surprise.\n")
        return
    if not output_dir:
        result_text.insert(tk.END, "Error: No output directory specified. Typical.\n")
        return
    if not os.path.isdir(output_dir):
        try:
            os.makedirs(output_dir)
            result_text.insert(tk.END, f"Created output directory '{output_dir}'.\n")
        except OSError as e:
            result_text.insert(tk.END, f"Error: Couldn’t create directory '{output_dir}': {e}.\n")
            return

    mock_user(STAGE_2_EXTRACT_INSULTS)
    result_text.insert(tk.END, f"Extracting '{file_path}' to '{output_dir}'...\n")
    result_text.see(tk.END)
    threading.Thread(target=do_extraction, daemon=True).start()

# Function to extract only search results, preserving directory structure
def extract_search_results(result_text):
    global search_results
    if not search_results:
        result_text.insert(tk.END, "No search results to extract. What are you even doing?\n")
        return

    output_dir = filedialog.askdirectory(title="Select Output Directory for Search Results")
    if not output_dir:
        result_text.insert(tk.END, "Extraction cancelled. Too scared to choose a folder?\n")
        return

    # Create progress window for search results extraction
    progress_window = tb.Toplevel()
    progress_window.title("Extracting Search Results...")
    progress_window.geometry("300x100")
    progress_window.transient(result_text.master)  # Tie to main window
    progress_window.grab_set()  # Make it modal
    
    progress_label = tk.Label(progress_window, text="Extracting search results, hold tight...")
    progress_label.pack(pady=10)
    
    progress_bar = ttk.Progressbar(progress_window, mode="determinate", maximum=100)
    progress_bar.pack(pady=10, padx=20, fill="x")
    
    total_files = tk.DoubleVar(value=0)

    def do_extract_search():
        try:
            pygame.mixer.init()
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            mp3_path = os.path.join(base_path, "loading_music.mp3")
            
            result_text.insert(tk.END, f"Trying to load music from: {mp3_path}\n")
            result_text.see(tk.END)
            
            try:
                if os.path.exists(mp3_path):
                    pygame.mixer.music.load(mp3_path)
                    result_text.insert(tk.END, "Music loaded successfully, starting playback...\n")
                    result_text.see(tk.END)
                    pygame.mixer.music.play(-1)  # Loop indefinitely
                else:
                    result_text.insert(tk.END, f"Error: Music file 'loading_music.mp3' not found at {mp3_path}. No tunes for you.\n")
                    result_text.see(tk.END)
            except Exception as e:
                result_text.insert(tk.END, f"Error: Couldn’t play music - {e}. No tunes for you.\n")
                result_text.see(tk.END)

            total_extracted = 0
            for pak_path, file in search_results:
                try:
                    with zipfile.ZipFile(pak_path, 'r') as pak_file:
                        dest = os.path.join(output_dir, file)
                        os.makedirs(os.path.dirname(dest), exist_ok=True)
                        result_text.insert(tk.END, f"Extracting search result from {os.path.basename(pak_path)}: {file}\n")
                        result_text.see(tk.END)
                        with pak_file.open(file) as asset_file, open(dest, 'wb') as out_file:
                            shutil.copyfileobj(asset_file, out_file)
                        total_extracted += 1
                        progress_bar['value'] = (total_extracted / len(search_results)) * 100
                        progress_window.update_idletasks()
                        time.sleep(0.01)  # Small delay to ensure GUI updates
                except Exception as e:
                    result_text.insert(tk.END, f"Error extracting {file} from {os.path.basename(pak_path)}: {e}\n")
                    result_text.see(tk.END)
                    continue

            result_text.insert(tk.END, "\nSearch results extraction complete. Check the folder, genius.\n")
            subprocess.Popen(f'explorer "{output_dir}"')
        except Exception as e:
            result_text.insert(tk.END, f"Critical error: {e}. You’re truly cursed.\n")
        finally:
            if pygame.mixer.get_init():
                pygame.mixer.music.stop()
            progress_window.destroy()

    mock_user(STAGE_2_EXTRACT_INSULTS)
    result_text.insert(tk.END, f"Extracting search results to '{output_dir}'...\n")
    result_text.see(tk.END)
    threading.Thread(target=do_extract_search, daemon=True).start()

# Function to find KCD2 path
def find_kcd2_path():
    path = None
    kcd2_rel_path = os.path.join("steamapps", "common", "KingdomComeDeliverance2")
    for disk in psutil.disk_partitions():
        kcd2_abs_path = None
        if disk.mountpoint == "C:\\":
            kcd2_abs_path = os.path.join(disk.mountpoint, "Program Files (x86)", "Steam", kcd2_rel_path)
        else:
            kcd2_abs_path = os.path.join(disk.mountpoint, "SteamLibrary", kcd2_rel_path)
        if os.path.isdir(kcd2_abs_path):
            path = kcd2_abs_path
            break
    return path

# Main GUI Application
class OmniPakSearcherApp(tb.Window):
    def __init__(self):
        super().__init__(themename=THEME)  # Dark modern theme
        self.title(PROGRAM_NAME)
        self.geometry("600x450")

        # File path entry
        self.file_path = tk.StringVar()
        tk.Label(self, text=".pak File Path or Folder:").grid(row=0, column=0, padx=10, pady=10)
        tk.Entry(self, textvariable=self.file_path, width=50).grid(row=0, column=1, padx=10, pady=10)
        tb.Button(self, text="Browse", command=self.browse_file, bootstyle="primary").grid(row=0, column=2, padx=10, pady=10)

        # Search term entry
        self.search_term = tk.StringVar()
        tk.Label(self, text="Search Term:").grid(row=1, column=0, padx=10, pady=10)
        tk.Entry(self, textvariable=self.search_term, width=50).grid(row=1, column=1, padx=10, pady=10)

        # Output directory entry
        self.output_dir = tk.StringVar()
        tk.Label(self, text="Output Directory:").grid(row=2, column=0, padx=10, pady=10)
        tk.Entry(self, textvariable=self.output_dir, width=50).grid(row=2, column=1, padx=10, pady=10)
        tb.Button(self, text="Browse", command=self.browse_output_dir, bootstyle="primary").grid(row=2, column=2, padx=10, pady=10)

        # Action buttons
        tb.Button(self, text="Search", command=self.search, bootstyle="success").grid(row=3, column=1, pady=10)
        tb.Button(self, text="Extract", command=self.extract, bootstyle="danger").grid(row=4, column=1, pady=10)

        # Turn Off Insults button
        self.insults_enabled = tk.BooleanVar(value=True)
        tb.Checkbutton(self, text="Turn Off Insults", variable=self.insults_enabled, command=self.toggle_insults).grid(row=5, column=1, pady=10)

        # Result text box
        self.result_text = tk.Text(self, height=10, width=70)
        self.result_text.grid(row=6, column=0, columnspan=3, padx=10, pady=10)

    def browse_file(self):
        kcd2_path = find_kcd2_path()
        initial_dir = kcd2_path if kcd2_path else os.path.expanduser("~")
        
        # Offer choice between file and folder
        choice = messagebox.askyesno("Browse Option", "Browse for a single .pak file, or a folder containing .pak files?")
        if choice:  # Yes = Folder
            path = filedialog.askdirectory(initialdir=initial_dir, title="Select Folder Containing .pak Files")
        else:  # No = Single File
            path = filedialog.askopenfilename(initialdir=initial_dir, filetypes=[("PAK Files", "*.pak")], title="Select a .pak File")
        
        if path:
            self.file_path.set(path)
            mock_user(STAGE_1_INSULTS)

    def browse_output_dir(self):
        output_dir = filedialog.askdirectory()
        self.output_dir.set(output_dir)

    def search(self):
        global INSULTS_ENABLED, search_results
        search_results = []  # Reset search results
        INSULTS_ENABLED = self.insults_enabled.get()
        self.result_text.delete(1.0, tk.END)
        search_in_pak(self.file_path.get(), self.search_term.get(), self.result_text)

    def extract(self):
        global INSULTS_ENABLED
        INSULTS_ENABLED = self.insults_enabled.get()
        self.result_text.delete(1.0, tk.END)
        
        # Create progress window
        progress_window = tb.Toplevel(self)
        progress_window.title("Extracting...")
        progress_window.geometry("300x100")
        progress_window.transient(self)  # Tie it to the main window
        progress_window.grab_set()  # Make it modal
        
        progress_label = tk.Label(progress_window, text="Extracting files, hold tight...")
        progress_label.pack(pady=10)
        
        progress_bar = ttk.Progressbar(progress_window, mode="determinate", maximum=100)
        progress_bar.pack(pady=10, padx=20, fill="x")
        
        total_files = tk.DoubleVar(value=0)
        extract_from_pak(self.file_path.get(), self.output_dir.get(), self.result_text, progress_window, progress_bar, total_files)

    def toggle_insults(self):
        global INSULTS_ENABLED
        INSULTS_ENABLED = self.insults_enabled.get()
        if INSULTS_ENABLED:
            messagebox.showinfo("Roasting You", "Insults are back on. Brace yourself!")
        else:
            messagebox.showinfo("Roasting You", "Insults are off. How boring.")

# Run the application
if __name__ == "__main__":
    app = OmniPakSearcherApp()
    app.mainloop()