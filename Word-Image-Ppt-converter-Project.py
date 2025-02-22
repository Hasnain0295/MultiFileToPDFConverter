import os
import win32com.client
from tkinter import Tk, filedialog
from PIL import Image


def convert_word_to_pdf(word_file, pdf_output):
    """
    Convert a Word document (.docx) to a PDF.
    """
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(word_file)
    doc.SaveAs(pdf_output, FileFormat=17)  # 17 is the code for PDF format
    doc.Close()
    word.Quit()
    print(f"Word document converted to PDF: {pdf_output}")


def convert_ppt_to_pdf(ppt_file, pdf_output):
    """
    Convert a PowerPoint presentation (.pptx) to a PDF.
    """
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    presentation = powerpoint.Presentations.Open(ppt_file)
    presentation.SaveAs(pdf_output, FileFormat=32)  # 32 is the code for PDF format
    presentation.Close()
    powerpoint.Quit()
    print(f"PowerPoint presentation converted to PDF: {pdf_output}")


def convert_images_to_pdf(image_paths, pdf_output):
    """
    Converts multiple images into a single PDF.
    """
    images = []
    for img_path in image_paths:
        try:
            print(f"Opening image: {img_path}")  # Debugging: Show image path
            image = Image.open(img_path)
            images.append(image.convert('RGB'))
        except Exception as e:
            print(f"Error opening {img_path}: {e}")

    if images:
        images[0].save(pdf_output, save_all=True, append_images=images[1:])
        print(f"Images converted to PDF: {pdf_output}")
    else:
        print("No valid images to convert.")


def select_file():
    """
    Opens a file dialog to select a Word, PowerPoint, or Image file.
    """
    root = Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(
        title="Select File (Word, PowerPoint, or Image)",
        filetypes=[("Word Files", "*.docx"), ("PowerPoint Files", "*.pptx"),
                   ("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif")]
    )
    return file_path


def main():
    print("Select the file you want to convert to PDF.")
    file_path = select_file()  # Select the file using a dialog box
    if file_path:
        # Ask for the output PDF file location
        output_pdf_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")], title="Save PDF As"
        )

        if output_pdf_path:
            if file_path.lower().endswith('.docx'):
                convert_word_to_pdf(file_path, output_pdf_path)
            elif file_path.lower().endswith('.pptx'):
                convert_ppt_to_pdf(file_path, output_pdf_path)
            elif file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
                convert_images_to_pdf([file_path], output_pdf_path)
            else:
                print("Invalid file format. Please select a Word, PowerPoint, or Image file.")
        else:
            print("No output file path selected.")
    else:
        print("No file selected.")


if __name__ == "__main__":
    main()
