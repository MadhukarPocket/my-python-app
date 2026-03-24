import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
import re  # Added for smart text replacement

def delete_paragraph(paragraph):
    """Safely removes a paragraph XML element to retain all document formatting."""
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

def run_merger():
    # Hide the main background window for a cleaner look
    root = tk.Tk()
    root.withdraw()

    # 1. Pop-up to select the input file (Now generic)
    input_path = filedialog.askopenfilename(
        title="Select the Word Document to Merge",
        filetypes=[("Word Documents", "*.docx")]
    )
    if not input_path:
        return # User cancelled

    # 2. Pop-up to select where to save the output file (Now generic)
    output_path = filedialog.asksaveasfilename(
        title="Save the Merged Document as",
        defaultextension=".docx",
        filetypes=[("Word Documents", "*.docx")]
    )
    if not output_path:
        return # User cancelled

    try:
        # Load the document
        doc = Document(input_path)
        current_chapter_count = 0
        new_chapter_number = 1

        # Iterate through paragraphs
        for para in doc.paragraphs:
            # Check if paragraph is formatted as a Heading
            if para.style.name.startswith('Heading'):
                current_chapter_count += 1
                
                # If it's an EVEN chapter, delete the heading.
                if current_chapter_count % 2 == 0:
                    delete_paragraph(para)
                
                # If it's an ODD chapter, keep it and rename it sequentially.
                else:
                    # Find "Chapter X" (case-insensitive) and replace with the new sequential number.
                    updated_text = re.sub(r'(?i)Chapter\s*\d+', f'Chapter {new_chapter_number}', para.text)
                    
                    # Update the text only if "Chapter" was actually in the heading
                    if updated_text != para.text:
                        para.text = updated_text
                    
                    new_chapter_number += 1
        
        # Save the finalized document
        doc.save(output_path)
        
        # Calculate final chapter count
        final_count = new_chapter_number - 1
        
        # Show success message to the user
        messagebox.showinfo("Success!", f"Found {current_chapter_count} headings.\nSuccessfully condensed into {final_count} chapters!\n\nSaved to: {output_path}")

    except Exception as e:
        # Catch errors (like if the file is currently open in Word)
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}\n\nMake sure the document is closed in Word before running.")

if __name__ == "__main__":
    run_merger()
