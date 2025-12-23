import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import threading
from extractor import extract_userid_from_filename, create_xml
from mail import send_bulk_emails, load_email_config, load_recipients_from_xml
import yaml


class BulkEmailGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Examitage")
        self.geometry("900x700")

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.pdf_directory = ctk.StringVar(value="attachments")
        self.pdf_files = []
        self.recipients_count = 0

        self.create_widgets()

    def create_widgets(self):
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        title_label = ctk.CTkLabel(
            main_frame,
            text="Klausurrückgabe",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(0, 20))

        self.tabview = ctk.CTkTabview(main_frame)
        self.tabview.pack(fill="both", expand=True)

        self.tabview.add("1. PDF Files")
        self.tabview.add("2. Email Config")
        self.tabview.add("3. Send Emails")

        self.setup_pdf_tab()
        self.setup_config_tab()
        self.setup_send_tab()

    def setup_pdf_tab(self):
        tab = self.tabview.tab("1. PDF Files")

        dir_frame = ctk.CTkFrame(tab)
        dir_frame.pack(fill="x", padx=10, pady=10)

        ctk.CTkLabel(
            dir_frame,
            text="PDF Directory:",
            font=ctk.CTkFont(weight="bold")
        ).pack(side="left", padx=5)

        dir_entry = ctk.CTkEntry(
            dir_frame,
            textvariable=self.pdf_directory,
            width=400
        )
        dir_entry.pack(side="left", padx=5, fill="x", expand=True)

        browse_btn = ctk.CTkButton(
            dir_frame,
            text="Browse",
            command=self.browse_directory,
            width=100
        )
        browse_btn.pack(side="left", padx=5)

        scan_btn = ctk.CTkButton(
            dir_frame,
            text="Scan PDFs",
            command=self.scan_pdfs,
            width=100
        )
        scan_btn.pack(side="left", padx=5)

        list_frame = ctk.CTkFrame(tab)
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(
            list_frame,
            text="Found PDF Files:",
            font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", padx=5, pady=5)

        self.pdf_textbox = ctk.CTkTextbox(list_frame, height=300)
        self.pdf_textbox.pack(fill="both", expand=True, padx=5, pady=5)

        stats_frame = ctk.CTkFrame(tab)
        stats_frame.pack(fill="x", padx=10, pady=10)

        self.stats_label = ctk.CTkLabel(
            stats_frame,
            text="No PDFs scanned yet",
            font=ctk.CTkFont(size=12)
        )
        self.stats_label.pack(pady=5)

        bottom_btn_frame = ctk.CTkFrame(tab)
        bottom_btn_frame.pack(fill="x", pady=10)

        self.generate_xml_btn = ctk.CTkButton(
            bottom_btn_frame,
            text="Generate XML File",
            command=self.generate_xml,
            state="disabled",
            height=40
        )
        self.generate_xml_btn.pack(side="left", padx=10, expand=True, fill="x")

        self.continue_tab1_btn = ctk.CTkButton(
            bottom_btn_frame,
            text="Continue →",
            command=lambda: self.tabview.set("2. Email Config"),
            state="disabled",
            height=40,
            fg_color="green",
            hover_color="darkgreen"
        )
        self.continue_tab1_btn.pack(side="left", padx=10, expand=True, fill="x")

    def setup_config_tab(self):
        tab = self.tabview.tab("2. Email Config")

        load_yaml_frame = ctk.CTkFrame(tab)
        load_yaml_frame.pack(fill="x", padx=10, pady=10)

        load_yaml_btn = ctk.CTkButton(
            load_yaml_frame,
            text="Load Configuration from YAML",
            command=self.load_config,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        load_yaml_btn.pack(fill="x", padx=5, pady=5)

        scroll_frame = ctk.CTkScrollableFrame(tab)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(
            scroll_frame,
            text="From Address:",
            font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", pady=(10, 5))
        self.from_entry = ctk.CTkEntry(scroll_frame, placeholder_text="your.email@example.com")
        self.from_entry.pack(fill="x", pady=(0, 10))
        self.from_entry.bind("<KeyRelease>", lambda e: self.validate_email_config())

        ctk.CTkLabel(
            scroll_frame,
            text="Subject:",
            font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", pady=(10, 5))
        self.subject_entry = ctk.CTkEntry(scroll_frame, placeholder_text="Email subject")
        self.subject_entry.pack(fill="x", pady=(0, 10))
        self.subject_entry.bind("<KeyRelease>", lambda e: self.validate_email_config())

        ctk.CTkLabel(
            scroll_frame,
            text="Email Body:",
            font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", pady=(10, 5))
        self.body_textbox = ctk.CTkTextbox(scroll_frame, height=150)
        self.body_textbox.pack(fill="x", pady=(0, 10))
        self.body_textbox.bind("<KeyRelease>", lambda e: self.validate_email_config())

        ctk.CTkLabel(
            scroll_frame,
            text="CC (optional):",
            font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", pady=(10, 5))
        self.cc_entry = ctk.CTkEntry(scroll_frame, placeholder_text="cc@example.com")
        self.cc_entry.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            scroll_frame,
            text="BCC (optional):",
            font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", pady=(10, 5))
        self.bcc_entry = ctk.CTkEntry(scroll_frame, placeholder_text="bcc@example.com")
        self.bcc_entry.pack(fill="x", pady=(0, 10))

        save_btn = ctk.CTkButton(
            scroll_frame,
            text="Save Configuration to YAML",
            command=self.save_config,
            height=40
        )
        save_btn.pack(fill="x", pady=20)

        self.continue_tab2_btn = ctk.CTkButton(
            scroll_frame,
            text="Continue →",
            command=lambda: self.tabview.set("3. Send Emails"),
            state="disabled",
            height=40,
            fg_color="green",
            hover_color="darkgreen"
        )
        self.continue_tab2_btn.pack(fill="x", pady=(0, 20))

    def setup_send_tab(self):
        tab = self.tabview.tab("3. Send Emails")

        info_frame = ctk.CTkFrame(tab)
        info_frame.pack(fill="x", padx=10, pady=10)

        self.send_info_label = ctk.CTkLabel(
            info_frame,
            text="Ready to send emails. Please configure email settings first.",
            font=ctk.CTkFont(size=13)
        )
        self.send_info_label.pack(pady=20)

        self.progress_bar = ctk.CTkProgressBar(tab)
        self.progress_bar.pack(fill="x", padx=10, pady=10)
        self.progress_bar.set(0)

        ctk.CTkLabel(
            tab,
            text="Log:",
            font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 5))

        self.log_textbox = ctk.CTkTextbox(tab, height=300)
        self.log_textbox.pack(fill="both", expand=True, padx=10, pady=5)

        self.send_btn = ctk.CTkButton(
            tab,
            text="Send All Emails",
            command=self.send_emails,
            height=50,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="green",
            hover_color="darkgreen"
        )
        self.send_btn.pack(pady=20, padx=10, fill="x")

    def browse_directory(self):
        directory = filedialog.askdirectory(initialdir=self.pdf_directory.get())
        if directory:
            self.pdf_directory.set(directory)

    def validate_email_config(self):
        from_address = self.from_entry.get().strip()
        subject = self.subject_entry.get().strip()
        body = self.body_textbox.get("1.0", "end-1c").strip()

        if from_address and subject and body:
            self.continue_tab2_btn.configure(state="normal")
        else:
            self.continue_tab2_btn.configure(state="disabled")

    def scan_pdfs(self):
        directory = self.pdf_directory.get()

        if not os.path.exists(directory):
            messagebox.showerror("Error", f"Directory not found: {directory}")
            return

        self.pdf_files = [f for f in os.listdir(directory) if f.endswith('.pdf')]

        self.pdf_textbox.delete("1.0", "end")

        if not self.pdf_files:
            self.pdf_textbox.insert("1.0", "No PDF files found in the directory.")
            self.stats_label.configure(text="No PDFs found")
            self.generate_xml_btn.configure(state="disabled")
            return

        valid_count = 0
        invalid_files = []

        for pdf_file in self.pdf_files:
            userid = extract_userid_from_filename(pdf_file)
            if userid:
                self.pdf_textbox.insert("end", f"✓ {pdf_file} → {userid}\n")
                valid_count += 1
            else:
                self.pdf_textbox.insert("end", f"✗ {pdf_file} → NO USER ID FOUND\n")
                invalid_files.append(pdf_file)

        total = len(self.pdf_files)
        invalid_count = len(invalid_files)
        self.stats_label.configure(
            text=f"Total: {total} PDFs | Valid: {valid_count} | Invalid: {invalid_count}"
        )

        if invalid_files:
            self.pdf_textbox.insert("end", f"\n⚠️ Warning: {invalid_count} file(s) without valid user IDs\n")

        self.generate_xml_btn.configure(state="normal")

    def generate_xml(self):
        if not self.pdf_files:
            messagebox.showwarning("Warning", "Please scan PDFs first")
            return

        try:
            create_xml(self.pdf_files, output_file='mail.xml')
            messagebox.showinfo("Success", "XML file generated successfully!\n\nFile: mail.xml")
            self.log("XML file generated successfully")
            self.continue_tab1_btn.configure(state="normal")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate XML:\n{str(e)}")

    def load_config(self):
        filename = filedialog.askopenfilename(
            title="Select YAML Config File",
            filetypes=[("YAML files", "*.yml *.yaml"), ("All files", "*.*")],
            initialfile="content.yml"
        )

        if not filename:
            return

        try:
            with open(filename, 'r') as f:
                config = yaml.safe_load(f)

            self.from_entry.delete(0, "end")
            self.from_entry.insert(0, config.get('from_address', ''))

            self.subject_entry.delete(0, "end")
            self.subject_entry.insert(0, config.get('subject', ''))

            self.body_textbox.delete("1.0", "end")
            self.body_textbox.insert("1.0", config.get('body', ''))

            self.cc_entry.delete(0, "end")
            self.cc_entry.insert(0, config.get('cc', ''))

            self.bcc_entry.delete(0, "end")
            self.bcc_entry.insert(0, config.get('bcc', ''))

            self.validate_email_config()

            messagebox.showinfo("Success", "Configuration loaded successfully")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load config:\n{str(e)}")

    def save_config(self):
        filename = filedialog.asksaveasfilename(
            title="Save YAML Config File",
            defaultextension=".yml",
            filetypes=[("YAML files", "*.yml"), ("All files", "*.*")],
            initialfile="content.yml"
        )

        if not filename:
            return

        config = {
            'from_address': self.from_entry.get(),
            'subject': self.subject_entry.get(),
            'body': self.body_textbox.get("1.0", "end-1c"),
            'cc': self.cc_entry.get() if self.cc_entry.get() else None,
            'bcc': self.bcc_entry.get() if self.bcc_entry.get() else None
        }

        try:
            with open(filename, 'w') as f:
                yaml.dump(config, f, default_flow_style=False)

            messagebox.showinfo("Success", "Configuration saved successfully")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save config:\n{str(e)}")

    def log(self, message):
        self.log_textbox.insert("end", f"{message}\n")
        self.log_textbox.see("end")

    def send_emails(self):
        if not self.from_entry.get():
            messagebox.showwarning("Warning", "Please enter a 'From' address")
            self.tabview.set("2. Email Config")
            return

        if not os.path.exists('mail.xml'):
            messagebox.showwarning("Warning", "Please generate XML file first")
            self.tabview.set("1. PDF Files")
            return

        try:
            recipients = load_recipients_from_xml('mail.xml')
            recipient_count = len(recipients)
        except:
            messagebox.showerror("Error", "Failed to load recipients from mail.xml")
            return

        if not messagebox.askyesno(
                "Confirm",
                f"Send {recipient_count} email(s)?\n\nThis will send real emails!"
        ):
            return

        temp_config = {
            'from_address': self.from_entry.get(),
            'subject': self.subject_entry.get(),
            'body': self.body_textbox.get("1.0", "end-1c"),
            'cc': self.cc_entry.get() if self.cc_entry.get() else None,
            'bcc': self.bcc_entry.get() if self.bcc_entry.get() else None
        }

        with open('temp_config.yml', 'w') as f:
            yaml.dump(temp_config, f, default_flow_style=False)

        self.send_btn.configure(state="disabled", text="Sending...")
        self.log_textbox.delete("1.0", "end")
        self.progress_bar.set(0)

        thread = threading.Thread(
            target=self.send_emails_thread,
            args=(recipient_count,)
        )
        thread.start()

    def send_emails_thread(self, total_count):
        try:
            from mail import (
                send_mail,
                load_email_config,
                load_recipients_from_xml
            )

            config = load_email_config('temp_config.yml')
            recipients = load_recipients_from_xml('mail.xml')

            from_address = config.get('from_address')
            subject = config.get('subject', '')
            body = config.get('body', '')
            cc = config.get('cc')
            bcc = config.get('bcc')

            success_count = 0
            fail_count = 0

            for idx, recipient in enumerate(recipients, 1):
                to_address = recipient['email'] + '@myubt.de'
                attachment_path = 'attachments/' + recipient['attachment']

                progress = idx / total_count
                self.after(0, self.progress_bar.set, progress)

                self.after(0, self.log, f"[{idx}/{total_count}] Sending to: {to_address}")

                attachments = [attachment_path] if attachment_path else []

                success = send_mail(
                    to_address=to_address,
                    from_address=from_address,
                    subject=subject,
                    body=body,
                    attachments=attachments
                )

                if success:
                    self.after(0, self.log, f"✓ Success: {to_address}")
                    success_count += 1
                else:
                    self.after(0, self.log, f"✗ Failed: {to_address}")
                    fail_count += 1

            self.after(0, self.log, "\n" + "=" * 50)
            self.after(0, self.log, "Summary:")
            self.after(0, self.log, f"  Total: {total_count}")
            self.after(0, self.log, f"  Successful: {success_count}")
            self.after(0, self.log, f"  Failed: {fail_count}")
            self.after(0, self.log, "=" * 50)

            self.after(0, self.send_btn.configure, {"state": "normal", "text": "Send All Emails"})

            if fail_count == 0:
                self.after(0, messagebox.showinfo, "Complete",
                           f"All {success_count} emails sent successfully!")
            else:
                self.after(0, messagebox.showwarning, "Complete with errors",
                           f"Sent: {success_count}\nFailed: {fail_count}")

        except Exception as e:
            self.after(0, self.log, f"\n❌ Error: {str(e)}")
            self.after(0, self.send_btn.configure, {"state": "normal", "text": "Send All Emails"})
            self.after(0, messagebox.showerror, "Error", f"Failed to send emails:\n{str(e)}")

        finally:
            if os.path.exists('temp_config.yml'):
                os.remove('temp_config.yml')


if __name__ == "__main__":
    app = BulkEmailGUI()
    app.mainloop()