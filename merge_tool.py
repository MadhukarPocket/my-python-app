import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document

def delete_paragraph(paragraph):
    """Safely removes a paragraph XML element to retain all document formatting."""
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

def run_merger():
    # Hide the main background window for a cleaner look
    root = tk.Tk()
    root.withdraw()

    # 1. Pop-up to select the input file
    input_path = filedialog.askopenfilename(
        title="Select the 100-Chapter DOCX file",
        filetypes=[("Word Documents", "*.docx")]
    )
    if not input_path:
        return # User cancelled

    # 2. Pop-up to select where to save the output file
    output_path = filedialog.asksaveasfilename(
        title="Save the 50-Chapter DOCX as",
        defaultextension=".docx",
        filetypes=[("Word Documents", "*.docx")]
    )
    if not output_path:
        return # User cancelled

    try:
        # Load the document
        doc = Document(input_path)
        current_chapter_count = 0

        # Iterate through paragraphs
        for para in doc.paragraphs:
            # Check if paragraph is formatted as a Heading
            if para.style.name.startswith('Heading'):
                current_chapter_count += 1
                
                # If it's an EVEN chapter (2, 4, 6...), delete the heading.
                # The text remains, seamlessly merging into the ODD chapter above it.
                if current_chapter_count % 2 == 0:
                    delete_paragraph(para)
        
        # Save the finalized document
        doc.save(output_path)
        
        # Calculate final chapter count
        final_count = current_chapter_count // 2 + current_chapter_count % 2
        
        # Show success message to the user
        messagebox.showinfo("Success!", f"Found {current_chapter_count} headings.\nSuccessfully condensed into {final_count} chapters!\n\nSaved to: {output_path}")

    except Exception as e:
        # Catch errors (like if the file is currently open in Word)
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}\n\nMake sure the document is closed in Word before running.")

if __name__ == "__main__":
    run_merger()
