import os
import tkinter as tk
from tkinter import messagebox, filedialog
import win32com.client
from PIL import Image, ImageTk

# Store photo references globally so Tkinter doesn't accidentally delete them
scanned_tk_images = {"front": None, "back": None}

def auto_crop(img):
    """Cuts the image in half and trims away the white scanner background."""
    # Step 1: Cut the image to the top half (since the card is usually placed at the top)
    width, height = img.size
    top_half = img.crop((0, 0, width, height // 2))
    
    # Step 2: Auto-trim the remaining white space
    gray = top_half.convert("L")
    # Anything darker than very light gray is considered part of the card
    mask = gray.point(lambda p: 255 if p < 240 else 0)
    bbox = mask.getbbox()
    
    if bbox:
        # Add a tight 20px padding around the card
        pad = 20
        final_bbox = (
            max(0, bbox[0] - pad),
            max(0, bbox[1] - pad),
            min(top_half.width, bbox[2] + pad),
            min(top_half.height, bbox[3] + pad)
        )
        return top_half.crop(final_bbox)
        
    return top_half # Fallback just in case

def scan_image(side):
    """Triggers the Windows scanner, crops the card, and loads it into the UI."""
    try:
        status_label.config(text="Connecting to scanner... Check for Windows prompt!")
        root.update()

        # Connect to Windows Image Acquisition (WIA)
        wia_dialog = win32com.client.Dispatch("WIA.CommonDialog")
        
        # Show the native scan prompt
        image = wia_dialog.ShowAcquireImage()
        
        if image is not None:
            file_path = os.path.abspath(f"temp_scan_{side}.jpg")
            
            # Delete old temp file if it exists
            if os.path.exists(file_path):
                os.remove(file_path)
                
            # Save the raw full-page scan locally first
            image.SaveFile(file_path)
            
            # Open the raw scan, CROP it, and save it back
            raw_img = Image.open(file_path)
            cropped_img = auto_crop(raw_img)
            cropped_img.save(file_path)
            
            # Create a preview of the *cropped* image for the UI
            preview_img = cropped_img.copy()
            preview_img.thumbnail((350, 250)) 
            tk_img = ImageTk.PhotoImage(preview_img)
            
            # Save the reference globally so it doesn't disappear
            scanned_tk_images[side] = tk_img
            
            # Update the appropriate label
            if side == "front":
                front_image_label.config(image=scanned_tk_images["front"], text="")
            elif side == "back":
                back_image_label.config(image=scanned_tk_images["back"], text="")
            
            status_label.config(text=f"Successfully scanned {side}!")
        else:
            status_label.config(text="Scan cancelled.")
            
    except Exception as e:
        messagebox.showerror("Scanner Error", f"Failed to scan: {str(e)}")
        status_label.config(text="Ready")

def save_to_pdf():
    """Combines the front and back cropped images into a single PDF with a small margin."""
    try:
        front_path = os.path.abspath("temp_scan_front.jpg")
        back_path = os.path.abspath("temp_scan_back.jpg")
        
        images_to_save = []
        
        if os.path.exists(front_path):
            images_to_save.append(Image.open(front_path).convert('RGB'))
        if os.path.exists(back_path):
            images_to_save.append(Image.open(back_path).convert('RGB'))
            
        if not images_to_save:
            messagebox.showwarning("No Scans Found", "Please scan at least one side of the card.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Documents", "*.pdf")],
            title="Save ID Card as PDF"
        )
        
        if not save_path:
            return 

        # Create a blank Letter-sized canvas (8.5x11 inches at 150 DPI)
        page_width, page_height = 1275, 1650
        pdf_page = Image.new('RGB', (page_width, page_height), 'white')
        
        # Calculate starting height to center the cards vertically on the page
        total_height = sum(img.height for img in images_to_save) + 50 # 50px is the small margin
        y_offset = (page_height - total_height) // 2
        
        if y_offset < 100: 
            y_offset = 100 # Ensure it doesn't hit the very top edge

        for img in images_to_save:
            # Center the card horizontally
            x_offset = (page_width - img.width) // 2
            
            # Paste the card onto our blank page
            pdf_page.paste(img, (x_offset, y_offset))
            
            # Move the vertical offset down exactly the height of the card PLUS a very small 50px margin
            y_offset += img.height + 50 

        # Save the beautifully stitched page!
        pdf_page.save(save_path, "PDF", resolution=150.0)
        
        status_label.config(text=f"Saved PDF to: {os.path.basename(save_path)}")
        messagebox.showinfo("Success!", "The ID card has been saved perfectly!")

    except Exception as e:
        messagebox.showerror("Save Error", f"Failed to save PDF: {str(e)}")

# ==========================================
# Build the UI
# ==========================================
root = tk.Tk()
root.title("Mission Office - ID Card Scanner")
root.geometry("800x650")
root.configure(bg="#f0f0f0")

header_label = tk.Label(root, text="Mission Insurance Card Scanner", font=("Arial", 16, "bold"), bg="#f0f0f0")
header_label.pack(pady=10)

status_label = tk.Label(root, text="Ready. Place card near the TOP of the scanner.", font=("Arial", 10, "italic"), bg="#f0f0f0", fg="#333")
status_label.pack(pady=5)

btn_frame = tk.Frame(root, bg="#f0f0f0")
btn_frame.pack(pady=10)

btn_front = tk.Button(btn_frame, text="1. Scan Front", font=("Arial", 12, "bold"), bg="#0056b3", fg="white", width=15, command=lambda: scan_image("front"))
btn_front.grid(row=0, column=0, padx=20)

btn_back = tk.Button(btn_frame, text="2. Scan Back", font=("Arial", 12, "bold"), bg="#0056b3", fg="white", width=15, command=lambda: scan_image("back"))
btn_back.grid(row=0, column=1, padx=20)

img_frame = tk.Frame(root, bg="#f0f0f0")
img_frame.pack(pady=20)

front_label = tk.Label(img_frame, text="Front of Card", font=("Arial", 12, "bold"), bg="#f0f0f0")
front_label.grid(row=0, column=0, padx=10)

front_image_label = tk.Label(img_frame, text="[ Image will appear here ]", width=40, height=15, bg="#ddd", relief="sunken")
front_image_label.grid(row=1, column=0, padx=10)

back_label = tk.Label(img_frame, text="Back of Card", font=("Arial", 12, "bold"), bg="#f0f0f0")
back_label.grid(row=0, column=1, padx=10)

back_image_label = tk.Label(img_frame, text="[ Image will appear here ]", width=40, height=15, bg="#ddd", relief="sunken")
back_image_label.grid(row=1, column=1, padx=10)

save_btn = tk.Button(root, text="3. Save as PDF", font=("Arial", 14, "bold"), bg="#28a745", fg="white", width=25, command=save_to_pdf)
save_btn.pack(pady=20)

root.mainloop()