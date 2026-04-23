import os
import tkinter as tk
from tkinter import messagebox, filedialog
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
                front_image_label.image = tk_img  
            elif side == "back":
                back_image_label.config(image=tk_img, text="")
                back_image_label.image = tk_img
            
            status_label.config(text=f"Successfully scanned {side}!")
        else:
            status_label.config(text="Scan cancelled.")
            
    except Exception as e:
        messagebox.showerror("Scanner Error", f"Failed to scan: {str(e)}")
        status_label.config(text="Ready")

def save_to_pdf():
    """Combines the front and back temporary images into a single PDF."""
    try:
        front_path = os.path.abspath("temp_scan_front.jpg")
        back_path = os.path.abspath("temp_scan_back.jpg")
        
        images_to_save = []
        
        # Check which scans exist
        if os.path.exists(front_path):
            images_to_save.append(Image.open(front_path).convert('RGB'))
        if os.path.exists(back_path):
            images_to_save.append(Image.open(back_path).convert('RGB'))
            
        if not images_to_save:
            messagebox.showwarning("No Scans Found", "Please scan at least one side of the card before saving.")
            return

        # Prompt user for where to save the file
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Documents", "*.pdf")],
            title="Save ID Card as PDF"
        )
        
        if not save_path:
            return  # User cancelled the save dialog

        # Create a blank Letter-sized canvas (8.5x11 inches at 150 DPI = 1275x1650 pixels)
        page_width, page_height = 1275, 1650
        pdf_page = Image.new('RGB', (page_width, page_height), 'white')
        
        y_offset = 150  # Starting vertical margin
        
        for img in images_to_save:
            # Scale the image so it fits nicely on the page
            img.thumbnail((1000, 1000)) 
            
            # Center the image horizontally
            x_offset = (page_width - img.width) // 2
            
            # Paste the card onto our blank page
            pdf_page.paste(img, (x_offset, y_offset))
            
            # Move the vertical offset down for the next card
            y_offset += img.height + 150 

        # Save the stitched page as a PDF!
        pdf_page.save(save_path, "PDF", resolution=150.0)
        
        status_label.config(text=f"Saved PDF to: {os.path.basename(save_path)}")
        messagebox.showinfo("Success!", "The ID card has been saved as a PDF!")

    except Exception as e:
        messagebox.showerror("Save Error", f"Failed to save PDF: {str(e)}")

# ==========================================
# Build the UI
# ==========================================
root = tk.Tk()
root.title("Mission Office - ID Card Scanner")
root.geometry("800x650") # Made the window slightly taller
root.configure(bg="#f0f0f0")

# Header
header_label = tk.Label(root, text="Mission ID/Insurance Card Scanner", font=("Arial", 16, "bold"), bg="#f0f0f0")
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

# Save to PDF Button
save_btn = tk.Button(root, text="3. Save as PDF", font=("Arial", 14, "bold"), bg="#28a745", fg="white", width=25, command=save_to_pdf)
save_btn.pack(pady=20)

# Run the App
root.mainloop()