
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import threading
import time
import webbrowser
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def remove_emojis(text):
    return ''.join(c for c in text if ord(c) <= 0xFFFF)

class TanwarWASenderApp:
    def __init__(self, root):
        self.root = root
        root.title("TanwarWASender")
        root.geometry("850x800")
        root.configure(bg="#f0f4f7")
        root.resizable(False, False)
        self.filename = None

        tk.Label(root, text="TanwarWASender", font=("Arial", 18, "bold"), bg="#f0f4f7", fg="#2c3e50").pack(pady=(15, 5))
        tk.Button(root, text="📂 Select Excel File", command=self.load_excel, font=("Arial", 11), bg="#2980b9", fg="white").pack()

        # Excel File Info Viewer
        self.file_display = scrolledtext.ScrolledText(root, height=4, width=100, font=("Arial", 9))
        self.file_display.pack(pady=(5, 10))

        tk.Label(root, text="✏️ Message Template (Use {name}, {brand}, {sender_name}, {mobile}, {email}, {linkedin}):", font=("Arial", 10, "bold"), bg="#f0f4f7", fg="#2c3e50").pack()
        self.message_input = scrolledtext.ScrolledText(root, height=8, width=100, font=("Arial", 10))
        self.message_input.insert(tk.END, "Dear {name} ji,\n\nGreetings from *{brand}*...\n\nWarm regards,\n*{sender_name}*\nMobile: {mobile}\nEmail: {email}\nLinkedIn: {linkedin}")
        self.message_input.pack(pady=5)

        entry_frame = tk.Frame(root, bg="#f0f4f7")
        entry_frame.pack(pady=5)

        labels = ["Firm / Brand Name:", "Sender Name:", "Mobile:", "Email:", "LinkedIn URL:"]
        defaults = [
            "Tanwar Ratan Singh & Associates",
            "CA Ratan Singh Tanwar",
            "+91-8875483332",
            "caratansinghtanwar@gmail.com",
            "https://www.linkedin.com/in/tanwarratansingh"
        ]
        self.entries = []

        for i, (label, default) in enumerate(zip(labels, defaults)):
            tk.Label(entry_frame, text=label, font=("Arial", 9), bg="#f0f4f7").grid(row=i, column=0, sticky="e", padx=10, pady=2)
            entry = tk.Entry(entry_frame, width=60)
            entry.insert(0, default)
            entry.grid(row=i, column=1, padx=10, pady=2)
            self.entries.append(entry)

        tk.Button(root, text="🚀 Start Messaging", command=self.start_sending, bg="#27ae60", fg="white", font=("Arial", 12)).pack(pady=(10, 5))
        tk.Button(root, text="📱 Contact Us on WhatsApp", bg="#25D366", fg="white", font=("Arial", 10),
                  command=lambda: webbrowser.open("https://wa.me/918875483332")).pack(pady=(0, 15))

        # Footer Layout with Icon Buttons
        footer = tk.Frame(root, bg="#f0f4f7")
        footer.pack(fill=tk.X, padx=20)

        left = tk.Frame(footer, bg="#f0f4f7")
        left.pack(side=tk.LEFT, anchor="w")

        tk.Label(left, text="🧑‍💻 Developed by:", font=("Arial", 9, "bold"), bg="#f0f4f7", fg="#2c3e50").pack(anchor="w")
        tk.Label(left, text="CA Ratan Singh Tanwar", font=("Arial", 9), bg="#f0f4f7", fg="#7f8c8d").pack(anchor="w")

        icon_frame = tk.Frame(left, bg="#f0f4f7")
        icon_frame.pack(anchor="w", pady=2)

        linkedin_btn = tk.Button(icon_frame, text="🔗", font=("Arial", 12), bg="#0077b5", fg="white", width=2, command=lambda: webbrowser.open("https://www.linkedin.com/in/tanwarratansingh"))
        linkedin_btn.grid(row=0, column=0, padx=3)

        whatsapp_btn = tk.Button(icon_frame, text="💬", font=("Arial", 12), bg="#25D366", fg="white", width=2, command=lambda: webbrowser.open("https://wa.me/918875483332"))
        whatsapp_btn.grid(row=0, column=1, padx=3)

        right = tk.Frame(footer, bg="#f0f4f7")
        right.pack(side=tk.RIGHT, anchor="e")
        tk.Label(right, text="Tanwar Ratan Singh & Associates", font=("Arial", 9, "bold"), bg="#f0f4f7", fg="#2c3e50").pack(anchor="e")
        tk.Label(right, text="📞 +91-8875483332", font=("Arial", 9), bg="#f0f4f7", fg="#2c3e50").pack(anchor="e")
        tk.Label(right, text="✉️ caratansinghtanwar@gmail.com", font=("Arial", 9), bg="#f0f4f7", fg="#2c3e50").pack(anchor="e")

    def load_excel(self):
        self.filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.filename:
            try:
                df = pd.read_excel(self.filename)
                self.file_display.delete("1.0", tk.END)
                self.file_display.insert(tk.END, f"Loaded file: {self.filename}\n")
                self.file_display.insert(tk.END, f"Total contacts: {len(df)}\n")
                self.file_display.insert(tk.END, "\n".join(df["Name of Assessee"].astype(str).head(5).tolist()))
            except Exception as e:
                messagebox.showerror("Error", f"Could not read Excel: {e}")

    def start_sending(self):
        if not self.filename:
            messagebox.showwarning("No File", "Please select an Excel file first.")
            return
        threading.Thread(target=self.run_script).start()

    def run_script(self):
        df = pd.read_excel(self.filename)
        df.dropna(subset=["Name of Assessee", "Mobile"], inplace=True)

        user_message = self.message_input.get("1.0", tk.END).strip()
        brand = self.entries[0].get().strip()
        sender = self.entries[1].get().strip()
        mobile = self.entries[2].get().strip()
        email = self.entries[3].get().strip()
        linkedin = self.entries[4].get().strip()

        if not user_message:
            messagebox.showwarning("Empty Message", "Please enter a message to send.")
            return

        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            driver.get("https://web.whatsapp.com")

            proceed = messagebox.askokcancel("Scan QR", "Scan the QR in Chrome then click OK.")
            if not proceed:
                driver.quit()
                return

            for _, row in df.iterrows():
                name = str(row["Name of Assessee"]).strip()
                number = ''.join(filter(str.isdigit, str(row["Mobile"]).split('.')[0]))

                if len(number) == 10:
                    number = "91" + number
                elif not number.startswith("91"):
                    continue

                url = f"https://web.whatsapp.com/send?phone={number}&app_absent=0"
                driver.get(url)
                time.sleep(8)

                if "shared via url is invalid" in driver.page_source:
                    continue

                try:
                    msg_box = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.XPATH, "//footer//div[@contenteditable='true']"))
                    )

                    msg = user_message.replace("{name}", name + " ji").replace("{brand}", brand).replace("{sender_name}", sender).replace("{mobile}", mobile).replace("{email}", email).replace("{linkedin}", linkedin)
                    msg = remove_emojis(msg)

                    msg_box.click()
                    for line in msg.strip().split("\n"):
                        msg_box.send_keys(line)
                        msg_box.send_keys(Keys.SHIFT, Keys.ENTER)
                    msg_box.send_keys(Keys.ENTER)
                    time.sleep(3)
                except Exception:
                    continue

            driver.quit()
        except Exception as e:
            messagebox.showerror("Script Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = TanwarWASenderApp(root)
    root.mainloop()
