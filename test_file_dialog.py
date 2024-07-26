from tkinter import Tk, filedialog

def select_file():
    root = Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
    )
    if file_path:
        print(f"Selected file: {file_path}")
    else:
        print("No file selected.")

if __name__ == '__main__':
    select_file()
