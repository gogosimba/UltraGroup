import win32com.client
from pathlib import Path
import tkinter as tk
from tkinter import Listbox

class EmailProcessor:
    def __init__(self):
        self.content_text = {}
        self.removed_text_indices = {}

    def remove_specific_lines(self, email_body, subject):
        lines_to_remove = ["Från:", "Skickat:", "Till:", "Ämne:"]
        removed_text_indices = []

        lines = email_body.split('\n')
        cleaned_lines = [line for line in lines if not any(line.startswith(phrase) for phrase in lines_to_remove)]

        cleaned_body = '\n'.join(cleaned_lines)

        start = 1
        for line in lines:
            if any(line.startswith(phrase) for phrase in lines_to_remove):
                end = start + len(line)
                removed_text_indices.append((start, end))
            start += len(line) + 1

        return cleaned_body, removed_text_indices

    def save_email_to_file(self, email, output_dir, count, progress_text, subject_listbox):
        filename = f"email{count}"
        file_path = output_dir / filename

        email_body = email.Body
        original_body = email.Body

        cleaned_body, removed_indices = self.remove_specific_lines(email_body, email.Subject)

        self.removed_text_indices[email.Subject] = removed_indices

        with open(file_path.with_suffix('.txt'), 'w', encoding='utf-8') as file:
            file.write(f"Subject: {email.Subject}\n\n")
            file.write(f"Body:\n{original_body}")

        for attachment in email.Attachments:
            attachment_path = output_dir / f"{filename}_{attachment.FileName}"
            attachment.SaveAsFile(str(attachment_path))

        progress_text.insert(tk.END, f"Saved email {count} from {email.Parent.Name}\n")
        progress_text.yview(tk.END)

        subject_listbox.insert(tk.END, email.Subject)
        self.content_text[email.Subject] = original_body

    def get_email_content(self, subject):
        return self.content_text.get(subject, ""), self.removed_text_indices.get(subject, [])

    def retrieve_emails(self, output_dir, progress_text, subject_listbox):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # 6 is the folder number for Inbox in Outlook

        folder_names = ["22. Assignments from Combitech", "23. Assignments from Broccoli", "24. Assignments from Levigo"]

        email_count = 0

        for name in folder_names:
            try:
                folder = inbox.Folders[name]
                for email in folder.Items:
                    email_count += 1
                    self.save_email_to_file(email, output_dir, email_count, progress_text, subject_listbox)
            except Exception as e:
                progress_text.insert(tk.END, f"Error accessing folder {name}: {e}\n")
                progress_text.yview(tk.END)

    def run_script(self, output_folder, progress_text, subject_listbox):
        output_dir = Path(output_folder)
        output_dir.mkdir(parents=True, exist_ok=True)
        progress_text.delete(1.0, tk.END)
        subject_listbox.delete(0, tk.END)
        self.content_text.clear()
        self.removed_text_indices.clear()
        self.retrieve_emails(output_dir, progress_text, subject_listbox)
