import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import os
import fitz  # PyMuPDF for PDF processing
import pikepdf  # Unlock secured PDFs
from PIL import Image  # Image compression
import subprocess  # For FFmpeg (video compression)

# --- PDF Page Remover ---
def remove_pages():
    file_path = filedialog.askopenfilename(title="Select PDF", filetypes=[("PDF Files", "*.pdf")])
    if not file_path:
        return

    output_pdf = os.path.splitext(file_path)[0] + "_modified.pdf"
    pages_to_remove_str = simpledialog.askstring("Remove Pages", f"Enter page numbers to remove from {os.path.basename(file_path)} (comma-separated):")

    if not pages_to_remove_str:
        messagebox.showwarning("Skipped", "No pages selected. Operation canceled.")
        return

    try:
        pages_to_remove = sorted(set(int(p.strip()) for p in pages_to_remove_str.split(",")))
    except ValueError:
        messagebox.showerror("Error", "Invalid input. Please enter numbers separated by commas.")
        return

    doc = fitz.open(file_path)
    pages_to_remove = [p - 1 for p in pages_to_remove if 0 < p <= len(doc)]

    for page in reversed(pages_to_remove):
        doc.delete_page(page)

    doc.save(output_pdf)
    doc.close()
    messagebox.showinfo("Success", f"Modified PDF saved as:\n{output_pdf}")

# --- File Compression Functions ---
def compress_pdf():
    """Compress a selected PDF file."""
    file_path = filedialog.askopenfilename(title="Select PDF", filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        output_pdf = os.path.splitext(file_path)[0] + "_compressed.pdf"
        try:
            pdf = pikepdf.open(file_path)
            pdf.save(output_pdf, linearize=True)
            messagebox.showinfo("PDF Compression", f"Compressed PDF saved as:\n{output_pdf}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to compress PDF: {e}")

def compress_image():
    """Compress a selected image file."""
    file_path = filedialog.askopenfilename(title="Select Image", filetypes=[("Images", "*.jpg *.png *.jpeg *.gif *.bmp")])
    if file_path:
        output_image = os.path.splitext(file_path)[0] + "_compressed.jpg"
        try:
            quality = 50  # Compression level
            img = Image.open(file_path)
            img.save(output_image, "JPEG", quality=quality)
            messagebox.showinfo("Image Compression", f"Compressed image saved as:\n{output_image}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to compress image: {e}")

def compress_video():
    """Compress a selected video file."""
    file_path = filedialog.askopenfilename(title="Select Video", filetypes=[("Videos", "*.mp4 *.avi *.mov *.mkv")])
    if file_path:
        output_video = os.path.splitext(file_path)[0] + "_compressed.mp4"
        try:
            command = [
                "ffmpeg", "-i", file_path, "-b:v", "800k", output_video
            ]
            subprocess.run(command, check=True)
            messagebox.showinfo("Video Compression", f"Compressed video saved as:\n{output_video}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to compress video: {e}")

# --- File Compressor UI ---
def open_compressor_gui():
    """Opens the File Compressor menu."""
    compressor_window = tk.Toplevel(root)
    compressor_window.title("File Compressor")
    compressor_window.geometry("600x400")
    compressor_window.configure(bg="#D3D3D3")  

    frame = tk.Frame(compressor_window, bg="#D3D3D3", width=550, height=350)
    frame.place(relx=0.5, rely=0.5, anchor="center")

    title_label = tk.Label(frame, text="File Compressor", font=("Arial", 40, "bold"), bg="#D3D3D3")
    title_label.pack(pady=20)

    buttons_frame = tk.Frame(frame, bg="#D3D3D3")
    buttons_frame.pack()

    btn_style = {
        "font": ("Arial", 16, "bold"),
        "width": 10,  # Approx. 95px
        "height": 1
    }

    # Compress PDF Button
    tk.Label(buttons_frame, text="Compress PDF", font=("Arial", 14), bg="#D3D3D3").grid(row=0, column=0, padx=10, pady=10)
    tk.Button(buttons_frame, text="Start", bg="#00FF00", fg="black", **btn_style, command=compress_pdf).grid(row=0, column=1, padx=10, pady=10)

    # Compress Image Button
    tk.Label(buttons_frame, text="Compress Image", font=("Arial", 14), bg="#D3D3D3").grid(row=1, column=0, padx=10, pady=10)
    tk.Button(buttons_frame, text="Start", bg="#00FF00", fg="black", **btn_style, command=compress_image).grid(row=1, column=1, padx=10, pady=10)

    # Compress Video Button
    tk.Label(buttons_frame, text="Compress Video", font=("Arial", 14), bg="#D3D3D3").grid(row=2, column=0, padx=10, pady=10)
    tk.Button(buttons_frame, text="In progress", bg="#00FF00", fg="black", **btn_style, state="disabled").grid(row=2, column=1, padx=10, pady=10)

    footer_label = tk.Label(frame, text="© JJ | Officium.Inc", font=("Arial", 10), bg="#D3D3D3")
    footer_label.pack(side="bottom", pady=10)

# --- Main Officium Toolkit UI ---
root = tk.Tk()
root.title("Officium Toolkit")
root.geometry("600x400")
root.configure(bg="#D3D3D3")  

# Inner Frame
frame = tk.Frame(root, bg="#D3D3D3", width=550, height=350)
frame.place(relx=0.5, rely=0.5, anchor="center")

title_label = tk.Label(frame, text="Officium Toolkit", font=("Arial", 40, "bold"), bg="#D3D3D3")
title_label.pack(pady=20)

buttons_frame = tk.Frame(frame, bg="#D3D3D3")
buttons_frame.pack()

btn_style = {
    "font": ("Arial", 16, "bold"),
    "width": 10,  # Approx. 95px
    "height": 1
}

# PDF Page Remover Button
tk.Label(buttons_frame, text="PDF Page remover", font=("Arial", 14), bg="#D3D3D3").grid(row=0, column=0, padx=10, pady=10)
tk.Button(buttons_frame, text="Start", bg="#00FF00", fg="black", **btn_style, command=remove_pages).grid(row=0, column=1, padx=10, pady=10)

# File Size Compressor Button
tk.Label(buttons_frame, text="File size compressor", font=("Arial", 14), bg="#D3D3D3").grid(row=1, column=0, padx=10, pady=10)
tk.Button(buttons_frame, text="Start", bg="#00FF00", fg="black", **btn_style, command=open_compressor_gui).grid(row=1, column=1, padx=10, pady=10)

# Coming Soon Button (Disabled)
tk.Label(buttons_frame, text="Docx to PDF", font=("Arial", 14), bg="#D3D3D3").grid(row=2, column=0, padx=10, pady=10)
tk.Button(buttons_frame, text="Coming soon", bg="#00FF00", fg="black", **btn_style, state="disabled").grid(row=2, column=1, padx=10, pady=10)

footer_label = tk.Label(frame, text="© JJ | Officium.Inc", font=("Arial", 10), bg="#D3D3D3")
footer_label.pack(side="bottom", pady=10)

# Run the GUI
root.mainloop()
