import os
import fitz  # pymupdf for fast PDF processing
import pikepdf  # To unlock secured PDFs
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

def unlock_pdf(input_path, output_path, password):
    """Unlock a secured PDF using pikepdf."""
    try:
        with pikepdf.open(input_path, password=password) as pdf:
            pdf.save(output_path)
        return output_path
    except Exception as e:
        messagebox.showerror("Error", f"Failed to unlock PDF: {e}")
        return None

def remove_pages(input_pdf, output_pdf, pages_to_remove):
    """Remove multiple specified pages from a PDF."""
    doc = fitz.open(input_pdf)

    # Convert page numbers to zero-based index
    pages_to_remove = sorted(set([p - 1 for p in pages_to_remove if 0 < p <= len(doc)]), reverse=True)

    if not pages_to_remove:
        messagebox.showwarning("Warning", f"No valid pages to remove from {input_pdf}. Skipping.")
        return None
    
    for page in pages_to_remove:
        doc.delete_page(page)  # Delete the specified pages
    
    doc.save(output_pdf)
    doc.close()
    return output_pdf

def ask_remove_another_pdf():
    """GUI to ask if the user wants to remove pages from another PDF."""
    remove_another_window = tk.Toplevel(root)
    remove_another_window.title("Process Another PDF?")
    remove_another_window.geometry("400x200")
    remove_another_window.configure(bg="#D3D3D3")  # Light gray background

    # Title label
    label = tk.Label(remove_another_window, text="Process another PDF?", font=("Arial", 18, "bold"), bg="#D3D3D3")
    label.pack(pady=20)

    button_frame = tk.Frame(remove_another_window, bg="#D3D3D3")
    button_frame.pack(pady=20)

    def yes_action():
        remove_another_window.destroy()
        process_pdfs()  # Restart the process for another file

    def no_action():
        remove_another_window.destroy()
        messagebox.showinfo("Process Completed", "All PDFs processed successfully!")
        root.quit()  # End the program

    # Yes Button
    yes_button = tk.Button(button_frame, text="YES", font=("Arial", 14, "bold"), bg="#00FF00", fg="black", width=15, height=2, command=yes_action)
    yes_button.pack(side="left", padx=10)

    # No Button
    no_button = tk.Button(button_frame, text="NO", font=("Arial", 14, "bold"), bg="#FF0000", fg="black", width=15, height=2, command=no_action)
    no_button.pack(side="right", padx=10)

    remove_another_window.mainloop()

def process_pdfs():
    """Main function to handle PDF selection and processing."""
    root.withdraw()  # Hide the main window

    # Select PDFs
    input_pdfs = filedialog.askopenfilenames(title="Select PDFs", filetypes=[("PDF Files", "*.pdf")])
    
    if not input_pdfs:
        return
    
    output_dir = filedialog.askdirectory(title="Select Output Directory")
    
    if not output_dir:
        return

    for pdf_path in input_pdfs:
        file_name = os.path.basename(pdf_path)
        output_pdf_path = os.path.join(output_dir, file_name)

        # Ask the user which pages to remove
        pages_to_remove_str = simpledialog.askstring("Remove Pages", f"Enter the page numbers to remove from {file_name} (comma-separated):")
        
        if not pages_to_remove_str:
            messagebox.showwarning("Skipped", f"Skipping {file_name}. No pages selected.")
            continue
        
        try:
            pages_to_remove = [int(p.strip()) for p in pages_to_remove_str.split(",")]
        except ValueError:
            messagebox.showerror("Error", f"Invalid input for {file_name}. Please enter numbers separated by commas.")
            continue

        # Check if PDF is secured
        try:
            with pikepdf.open(pdf_path) as test_pdf:
                pass  # No password required
            unlocked_pdf = pdf_path  # No need to unlock
        except pikepdf._qpdf.PasswordError:
            password = simpledialog.askstring("Password Required", f"Enter password for {file_name}:")
            if not password:
                messagebox.showwarning("Skipped", f"Skipping {file_name} due to missing password.")
                continue
            unlocked_pdf = os.path.join(output_dir, f"unlocked_{file_name}")
            unlocked_pdf = unlock_pdf(pdf_path, unlocked_pdf, password)
            if not unlocked_pdf:
                continue
        
        # Process PDF and remove selected pages
        modified_pdf = remove_pages(unlocked_pdf, output_pdf_path, pages_to_remove)

        if modified_pdf and unlocked_pdf.startswith("unlocked_"):
            os.remove(unlocked_pdf)  # Remove temporary unlocked file

    # Ask if the user wants to process another PDF
    ask_remove_another_pdf()

def start_gui():
    """Creates the main GUI with a 'Start' button."""
    global root

    root = tk.Tk()
    root.title("PDF Page Remover")
    root.geometry("400x250")
    root.configure(bg="#D3D3D3")  # Light gray background

    # Title label
    label = tk.Label(root, text="PDF Page Remover", font=("Arial", 20, "bold"), bg="#D3D3D3")
    label.pack(pady=20)

    # Start Button
    start_button = tk.Button(root, text="Start", font=("Arial", 16, "bold"), bg="#B6FF6D", fg="black", width=15, height=2, command=process_pdfs)
    start_button.pack(pady=20)

    # Footer Label
    footer_label = tk.Label(root, text="© JJ | Officium.Inc", font=("Arial", 10), bg="#D3D3D3")
    footer_label.pack(side="bottom", pady=5)

    root.mainloop()

if __name__ == "__main__":
    start_gui()
