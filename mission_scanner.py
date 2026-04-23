import os
import tkinter as tk
from tkinter import messagebox
import win32com.client
from PIL import Image, ImageTk

def scan_image(side):
    """Triggers the Windows scanner and loads the image into the UI."""
    try:
        status_label.config(text="Connecting to scanner... Check for Windows prompt!")
        root.update()

        # Connect to Windows Image Acquisition (WIA)
        wia_dialog = win32com.client.Dispatch("WIA.CommonDialog")
        
        # Show the native scan prompt
        image = wia_dialog.ShowAcquireImage()
        
        if image is not None:
            # Create a temporary file path
            file_path = os.path.abspath(f"temp_scan_{side}.jpg")
            
            # Delete old temp file if it exists
            if os.path.exists(file_path):
                os.remove(file_path)
                
            # Save the new scan locally
            image.SaveFile(file_path)
            
            # Open the image with Pillow and resize it to fit the UI
            img = Image.open(file_path)
            img.thumbnail((350, 250))  # Resize to fit nicely in the window
            tk_img = ImageTk.PhotoImage(img)
            
            # Update the appropriate label based on the 'side' scanned
            if side == "front":
                front_image_label.config(image=tk_img, text="")
                front_image_label.image = tk_img  # Keep a reference so it doesn't get garbage collected
            elif side == "back":
                back_image_label.config(image=tk_img, text="")
                back_image_label.image = tk_img
            
            status_label.config(text=f"Successfully scanned {side}!")
        else:
            status_label.config(text="Scan cancelled.")
            
    except Exception as e:
        messagebox.showerror("Scanner Error", f"Failed to scan: {str(e)}")
        status_label.config(text="Ready")

# ==========================================
# Build the UI
# ==========================================
root = tk.Tk()
root.title("Mission Office - ID Card Scanner")
root.geometry("800x550")
root.configure(bg="#f0f0f0")

# Header
header_label = tk.Label(root, text="Mission Insurance Card Scanner", font=("Arial", 16, "bold"), bg="#f0f0f0")
header_label.pack(pady=10)

status_label = tk.Label(root, text="Ready. Place card on scanner and click a button.", font=("Arial", 10, "italic"), bg="#f0f0f0", fg="#333")
status_label.pack(pady=5)

# Buttons Frame
btn_frame = tk.Frame(root, bg="#f0f0f0")
btn_frame.pack(pady=10)

btn_front = tk.Button(btn_frame, text="1. Scan Front", font=("Arial", 12, "bold"), bg="#0056b3", fg="white", width=15, command=lambda: scan_image("front"))
btn_front.grid(row=0, column=0, padx=20)

btn_back = tk.Button(btn_frame, text="2. Scan Back", font=("Arial", 12, "bold"), bg="#0056b3", fg="white", width=15, command=lambda: scan_image("back"))
btn_back.grid(row=0, column=1, padx=20)

# Image Display Frame
img_frame = tk.Frame(root, bg="#f0f0f0")
img_frame.pack(pady=20)

# Front Image Display
front_label = tk.Label(img_frame, text="Front of Card", font=("Arial", 12, "bold"), bg="#f0f0f0")
front_label.grid(row=0, column=0, padx=10)

front_image_label = tk.Label(img_frame, text="[ Image will appear here ]", width=40, height=15, bg="#ddd", relief="sunken")
front_image_label.grid(row=1, column=0, padx=10)

# Back Image Display
back_label = tk.Label(img_frame, text="Back of Card", font=("Arial", 12, "bold"), bg="#f0f0f0")
back_label.grid(row=0, column=1, padx=10)

back_image_label = tk.Label(img_frame, text="[ Image will appear here ]", width=40, height=15, bg="#ddd", relief="sunken")
back_image_label.grid(row=1, column=1, padx=10)

# Run the App
root.mainloop()