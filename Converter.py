import os
import shutil
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class PublisherConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MS Publisher to PDF Mass Converter")
        self.root.geometry("600x350")

        self.lbl_info = tk.Label(root, text="Select root folder\n(All subfolders will be also scanned for .pub files)", font=("Arial", 10))
        self.lbl_info.pack(pady=15)

        self.btn_browse = tk.Button(root, text="Select Folder & Start", command=self.start_conversion, bg="#4CAF50", fg="white", font=("Arial", 11, "bold"))
        self.btn_browse.pack(pady=10)

        self.progress = ttk.Progressbar(root, orient="horizontal", length=500, mode="determinate")
        self.progress.pack(pady=15)

        self.lbl_status = tk.Label(root, text="In Progress...", fg="gray")

    def start_conversion(self):
        target_folder = filedialog.askdirectory(title="Select the folder with .pub files")
        if not target_folder:
            return

        self.btn_browse.config(state="disabled")
        self.lbl_status.config(text="Searching .pub files...", fg="black")

        thread = threading.Thread(target=self.convert_files, args=(target_folder,))
        thread.start()

    def convert_files(self, target_folder):
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()

        try:
            pub_files = []
            for root_dir, dirs, files in os.walk(target_folder):
                for file in files:
                    if file.lower().endswith('.pub'):
                        pub_files.append(os.path.join(root_dir, file))

            total_files = len(pub_files)
            if total_files == 0:
                self.root.after(0, lambda: messagebox.showinfo("Done, No .pub files detected on this folder"))
                self.root.after(0, self.reset_ui)
                return

            self.root.after(0, lambda: self.progress.config(maximum=total_files, value=0))

            pub_app = win32com.client.Dispatch("Publisher.Application")
            pbFixedFormatTypePDF = 2

            temp_dir = tempfile.gettempdir()

            for i, pub_path in enumerate(pub_files):
                filename = os.path.basename(pub_path)
                temp_pub_path = os.path.join(temp_dir, filename)

                pdf_filename = os.path.splitext(filename)[0] + ".pdf"
                temp_pdf_path = os.path.join(temp_dir, pdf_filename)

                final_pdf_path = os.path.join(os.path.dirname(pub_path), pdf_filename)

                self.root.after(0, lambda f=filename: self.lbl_status.config(text=f"Processing: {f}"))

                try:
                    shutil.copy2(pub_path, temp_pub_path)

                    doc = pub_app.Open(temp_pub_path)
                    doc.ExportAsFixedFormat(pbFixedFormatTypePDF, temp_pdf_path)
                    doc.Close()

                    if os.path.exists(temp_pdf_path):
                        shutil.move(temp_pdf_path, final_pdf_path)

                except Exception as e:
                    print(f"Error on file {filename}: {str(e)}")
                finally:
                    if os.path.exists(temp_pub_path):
                        os.remove(temp_pub_path)

                self.root.after(0, lambda val=i+1: self.progress.config(value=val))

            pub_app.Quit()

            self.root.after(0, lambda: messagebox.showinfo("Success", f"The conversion completed successfuly!\nTotal {total_files} PDF files created."))

        except Exception as e:
            self.root.after(0, lambda err=str(e): messagebox.showerror("Error", f"Unexpected Error:\n{err}"))
        finally:
            pythoncom.CoUninitialize()
            self.root.after(0, self.reset_ui)

    def reset_ui(self):
        self.btn_browse.config(state="normal")
        self.lbl_status.config(text="Standing by...", fg="gray")
        self.progress.config(value=0)

if __name__ == "__main__":
    root = tk.Tk()
    app = PublisherConverterApp(root)
    root.mainloop()