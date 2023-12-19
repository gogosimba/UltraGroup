import win32com.client
from pathlib import Path
import re
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
from tkinter import scrolledtext
from tkinter import Listbox

class EmailViewer(tk.Toplevel):
    def __init__(self, subject, original_body, removed_text_indices):
        super().__init__()

        self.title(subject)
        self.geometry("800x600")

        self.original_text = scrolledtext.ScrolledText(self, height=30, width=100, wrap=tk.WORD)
        self.original_text.pack(padx=10, pady=10)
        self.original_text.insert(tk.END, original_body)

        start = 1.0
        for removed_start, removed_end in removed_text_indices:
            self.original_text.tag_add("removed", f"{start + removed_start}", f"{start + removed_end}")
            self.original_text.tag_config("removed", background="red")

        self.original_text.config(state=tk.DISABLED)

def remove_specific_lines(email_body):
    # Define the lines to be removed
    lines_to_remove = ["Från:", "Skickat:", "Till:", "Ämne:"]

    # Remove lines starting with the specified phrases
    lines = email_body.split('\n')
    cleaned_lines = [line for line in lines if not any(line.startswith(phrase) for phrase in lines_to_remove)]

    # Join the remaining lines
    cleaned_body = '\n'.join(cleaned_lines)

    return cleaned_body

def save_email_to_file(email, output_dir, count, progress_text, subject_listbox, content_text, removed_text_indices):
    filename = f"email{count}"
    file_path = output_dir / filename

    email_body = email.Body
    original_body = email.Body

    removed_text = remove_specific_lines(email_body)
    removed_text_indices[email.Subject] = [(match.start(), match.end()) for match in re.finditer(re.escape(removed_text), original_body)]

    with open(file_path.with_suffix('.txt'), 'w', encoding='utf-8') as file:
        file.write(f"Subject: {email.Subject}\n\n")
        file.write(f"Body:\n{original_body}")

    for attachment in email.Attachments:
        attachment_path = output_dir / f"{filename}_{attachment.FileName}"
        attachment.SaveAsFile(str(attachment_path))

    progress_text.insert(tk.END, f"Saved email {count} from {email.Parent.Name}\n")
    progress_text.yview(tk.END)

    subject_listbox.insert(tk.END, email.Subject)
    content_text[email.Subject] = original_body

def retrieve_emails(output_dir, progress_text, subject_listbox, content_text, removed_text_indices):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  

    folder_names = ["22. Assignments from Combitech", "23. Assignments from Broccoli", "24. Assignments from Levigo"]

    email_count = 0

    for name in folder_names:
        try:
            folder = inbox.Folders[name]
            for email in folder.Items:
                email_count += 1
                save_email_to_file(email, output_dir, email_count, progress_text, subject_listbox, content_text, removed_text_indices)
        except Exception as e:
            progress_text.insert(tk.END, f"Error accessing folder {name}: {e}\n")
            progress_text.yview(tk.END)

def run_script(output_folder, progress_text, subject_listbox, content_text, removed_text_indices):
    try:
        output_dir = Path(output_folder)
        output_dir.mkdir(parents=True, exist_ok=True)
        progress_text.delete(1.0, tk.END)
        subject_listbox.delete(0, tk.END)
        content_text.clear()
        removed_text_indices.clear()
        retrieve_emails(output_dir, progress_text, subject_listbox, content_text, removed_text_indices)
        messagebox.showinfo("Success", "Email retrieval completed.")
    except Exception as ex:
        messagebox.showerror("Error", f"An error occurred: {ex}")

def choose_folder(entry):
    folder = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, folder)

def on_subject_select(event, subject_listbox, content_text, removed_text_indices):
    selected_index = subject_listbox.curselection()[0]
    subject = subject_listbox.get(selected_index)
    original_body = content_text.get(subject, "")
    removed_indices = removed_text_indices.get(subject, [])
    EmailViewer(subject, original_body, removed_indices)

def main():
    root = tk.Tk()
    root.title("Email Retrieval GUI")

    frame = ttk.Frame(root, padding="10")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    output_folder_label = ttk.Label(frame, text="Output Folder:")
    output_folder_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

    output_folder_entry = ttk.Entry(frame, width=50)
    output_folder_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)

    choose_folder_button = ttk.Button(frame, text="Choose Folder", command=lambda: choose_folder(output_folder_entry))
    choose_folder_button.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)

    progress_label = ttk.Label(frame, text="Progress:")
    progress_label.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky=tk.W)

    progress_text = scrolledtext.ScrolledText(frame, height=5, width=60)
    progress_text.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

    subject_label = ttk.Label(frame, text="Emails:")
    subject_label.grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)

    subject_listbox = Listbox(frame, selectmode=tk.SINGLE, height=10, width=60)
    subject_listbox.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

    content_text = {}
    removed_text_indices = {}

    subject_listbox.bind("<ButtonRelease-1>", lambda event: on_subject_select(event, subject_listbox, content_text, removed_text_indices))

    run_button = ttk.Button(frame, text="Run Script", command=lambda: run_script(output_folder_entry.get(), progress_text, subject_listbox, content_text, removed_text_indices))
    run_button.grid(row=5, column=0, columnspan=3, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
