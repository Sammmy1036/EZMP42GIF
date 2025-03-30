import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import threading
import os
import sys
import msvcrt
import tkinter.messagebox as tkmessagebox
import ctypes
import atexit
import psutil  # For checking if a process is running

# Global variables for progress bar and label
progress_bar = None
progress_label = None
convert_button = None

# Lock file for single-instance check
LOCK_FILE = "mp4_to_gif.lock"

# Dictionary to map display labels to numeric width values
width_mapping = {
    "720 (SD)": 720,
    "1080 (HD)": 1080,
    "2560 (2K)": 2560,
    "3840 (4K)": 3840
}

def hide_file(file_path):
    """Set the hidden attribute on a file (Windows-specific)."""
    try:
        # FILE_ATTRIBUTE_HIDDEN = 0x2
        ctypes.windll.kernel32.SetFileAttributesW(file_path, 0x2)
        print(f"Set hidden attribute on {file_path}")
    except Exception as e:
        print(f"Error setting hidden attribute on {file_path}: {e}")

def is_process_running(pid):
    """Check if a process with the given PID is running."""
    try:
        return psutil.pid_exists(pid) and psutil.Process(pid).status() != psutil.STATUS_ZOMBIE
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return False

def check_single_instance():
    try:
        # Check if the lock file exists and read the PID
        if os.path.exists(LOCK_FILE):
            with open(LOCK_FILE, 'rb') as f:
                pid_str = f.read().decode('utf-8').strip()
                try:
                    pid = int(pid_str)
                    if is_process_running(pid):
                        print("Another instance is already running (PID: {}).".format(pid))
                        tkmessagebox.showerror("Error", "Another instance of MP4 to GIF Converter is already running.")
                        return False
                    else:
                        print("Found stale lock file with PID {}. Removing it.".format(pid))
                        os.remove(LOCK_FILE)
                except ValueError:
                    print("Invalid PID in lock file. Removing it.")
                    os.remove(LOCK_FILE)

        # Create a new lock file
        lock_file_handle = open(LOCK_FILE, 'wb+')
        msvcrt.locking(lock_file_handle.fileno(), msvcrt.LK_NBLCK, 1)
        sys.lock_file_handle = lock_file_handle  # Store the handle to keep the file open
        lock_file_handle.write(str(os.getpid()).encode('utf-8'))
        lock_file_handle.flush()
        # Hide the lock file
        hide_file(LOCK_FILE)
        print("Single instance check passed")
        return True
    except IOError as e:
        if e.errno == 13:  # Permission denied (file is locked by another instance)
            print("Another instance is already running.")
            tkmessagebox.showerror("Error", "Another instance of MP4 to GIF Converter is already running.")
            return False
        else:
            print(f"Unexpected error checking lock file: {e}")
            tkmessagebox.showerror("Error", f"Unexpected error checking lock file: {e}")
            return False
    except Exception as e:
        print(f"Error in single instance check: {e}")
        tkmessagebox.showerror("Error", f"Error in single instance check: {e}")
        return False

def cleanup_lock_file():
    """Clean up the lock file when the application exits."""
    if hasattr(sys, 'lock_file_handle'):
        try:
            lock_file_handle = sys.lock_file_handle
            # Unlock the file
            msvcrt.locking(lock_file_handle.fileno(), msvcrt.LK_UNLCK, 1)
            # Close the file handle
            lock_file_handle.close()
            # Delete the file
            if os.path.exists(LOCK_FILE):
                os.remove(LOCK_FILE)
            print("Lock file cleaned up")
        except Exception as e:
            print(f"Error cleaning up lock file: {e}")
        finally:
            # Ensure the handle is removed from sys to prevent further attempts
            delattr(sys, 'lock_file_handle')

# Register cleanup with atexit to handle abnormal exits
atexit.register(cleanup_lock_file)

# Function to get the path to bundled resources (for PyInstaller)
def resource_path(relative_path):
    """Get the absolute path to a resource, works for dev and for PyInstaller."""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # If not running as a PyInstaller bundle, use the script's directory
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def mp4_to_gif(input_path, output_path, fps, width):
    global progress_bar, progress_label, convert_button
    try:
        # Ensure FFmpeg is available
        try:
            subprocess.run(["ffmpeg", "-version"], capture_output=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        except (subprocess.CalledProcessError, FileNotFoundError):
            raise Exception("FFmpeg is not installed or not found in PATH.")

        # Step 1: Generate a palette for better GIF quality
        palette_path = "palette.png"
        palette_cmd = [
            "ffmpeg", "-y", "-i", input_path,
            "-vf", f"fps={fps},scale={width}:-1:flags=lanczos,palettegen=stats_mode=full",
            palette_path
        ]
        subprocess.run(palette_cmd, check=True, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)

        # Step 2: Convert the video to GIF using the palette with improved dithering
        gif_cmd = [
            "ffmpeg", "-y", "-i", input_path, "-i", palette_path,
            "-filter_complex", f"fps={fps},scale={width}:-1:flags=lanczos[x];[x][1:v]paletteuse=dither=floyd_steinberg",
            output_path
        ]
        subprocess.run(gif_cmd, check=True, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)

        # Clean up the palette file
        if os.path.exists(palette_path):
            os.remove(palette_path)

        # Show success message (must be done in the main thread)
        root.after(0, lambda: messagebox.showinfo("Success", f"GIF saved as {output_path}"))
        
    except Exception as e:
        # Capture the error message in a local variable
        error_message = f"An error occurred: {str(e)}"
        # Show error message in the main thread
        root.after(0, lambda msg=error_message: messagebox.showerror("Error", msg))
    
    finally:
        # Hide the progress bar and label, re-enable the button (in the main thread)
        root.after(0, lambda: progress_label.pack_forget())
        root.after(0, lambda: progress_bar.pack_forget())
        root.after(0, lambda: convert_button.config(state="normal"))

def select_input_file():
    file_path = filedialog.askopenfilename(
        title="Select MP4 File",
        filetypes=[("MP4 files", "*.mp4"), ("All files", "*.*")]
    )
    if file_path:
        input_entry.delete(0, tk.END)
        input_entry.insert(0, file_path)

def select_output_directory():
    directory = filedialog.askdirectory(
        title="Select Output Directory"
    )
    if directory:
        output_dir_entry.delete(0, tk.END)
        output_dir_entry.insert(0, directory)

def convert_file():
    input_path = input_entry.get()
    output_dir = output_dir_entry.get().strip()
    output_name = output_filename_entry.get().strip()
    fps_value = fps_entry.get()
    width_label = width_var.get()  # Get the selected width label from the dropdown
    width = width_mapping[width_label]  # Map the label to the numeric value

    if not input_path:
        messagebox.showwarning("Input Error", "Please select an MP4 file.")
        return
    if not output_dir:
        messagebox.showwarning("Input Error", "Please select an output directory.")
        return
    if not output_name:
        messagebox.showwarning("Input Error", "Please enter an output filename.")
        return
    if not fps_value.isdigit() or int(fps_value) <= 0:
        messagebox.showwarning("Input Error", "Please enter a valid FPS (positive number).")
        return

    # Ensure the output has .gif extension
    if not output_name.endswith(".gif"):
        output_name += ".gif"
    
    # Combine the output directory and filename
    output_path = os.path.join(output_dir, output_name)
    
    fps = int(fps_value)

    # Show the progress bar and label, disable the button
    progress_label.pack(pady=5)
    progress_bar.pack(pady=5)
    convert_button.config(state="disabled")

    # Start the progress bar in indeterminate mode
    progress_bar.start()

    # Run the conversion in a separate thread
    thread = threading.Thread(target=mp4_to_gif, args=(input_path, output_path, fps, width))
    thread.start()

# Check for single instance before starting the GUI
if not check_single_instance():
    sys.exit(1)

# Set up the GUI
root = tk.Tk()
root.title("MP4 to GIF Converter")
root.geometry("500x500")  # Increased height to accommodate new fields

# Set the window icon dynamically
try:
    icon_path = resource_path("EZ MP4 2 GIF.ico")
    root.iconbitmap(icon_path)
except tk.TclError as e:
    print(f"Error setting icon: {e}")

# Register cleanup on window close
root.protocol("WM_DELETE_WINDOW", lambda: [cleanup_lock_file(), root.destroy()])

# Input file selection
tk.Label(root, text="Input MP4 File:").pack(pady=5)
input_entry = tk.Entry(root, width=50)
input_entry.pack(pady=5)
tk.Button(root, text="Browse", command=select_input_file, width=10).pack(pady=5)

# Output directory selection
tk.Label(root, text="Output Directory:").pack(pady=5)
output_dir_entry = tk.Entry(root, width=50)
output_dir_entry.pack(pady=5)
tk.Button(root, text="Browse", command=select_output_directory, width=10).pack(pady=5)

# Output filename
tk.Label(root, text="Output GIF Filename:").pack(pady=5)
output_filename_entry = tk.Entry(root, width=50)
output_filename_entry.pack(pady=5)

# FPS selection
tk.Label(root, text="FPS (Frames Per Second):").pack(pady=5)
fps_entry = tk.Entry(root, width=10)
fps_entry.insert(0, "15")  # Default FPS value
fps_entry.pack(pady=5)

# Width selection (dropdown)
tk.Label(root, text="Export Width:").pack(pady=5)
width_var = tk.StringVar(value="3840 (4K)")  # Default width
width_options = ["720 (SD)", "1080 (HD)", "2560 (2K)", "3840 (4K)"]
width_menu = tk.OptionMenu(root, width_var, *width_options)
width_menu.config(width=10)
width_menu.pack(pady=5)

# Progress bar (hidden by default, using indeterminate mode)
progress_label = tk.Label(root, text="Converting...")
progress_bar = ttk.Progressbar(root, mode="indeterminate", length=300)

# Convert button
convert_button = tk.Button(root, text="Convert to GIF", command=convert_file, width=15, bg="green", fg="white")
convert_button.pack(pady=20)

# Start the GUI loop
root.mainloop()