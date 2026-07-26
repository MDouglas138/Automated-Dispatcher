import os
import sys
import re
import sqlite3
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *


def setup_database():
    conn = sqlite3.connect("automations.db")
    cursor = conn.cursor()
    cursor.execute("""
                   CREATE TABLE IF NOT EXISTS automations
                   (
                       id
                       INTEGER
                       PRIMARY
                       KEY
                       AUTOINCREMENT,
                       task_name
                       TEXT,
                       file_path
                       TEXT,
                       sender_email
                       TEXT,
                       schedule_time
                       TEXT,
                       precheck_offset
                       INTEGER
                   )
                   """)
    cursor.execute("""
                   CREATE TABLE IF NOT EXISTS recipients
                   (
                       id
                       INTEGER
                       PRIMARY
                       KEY
                       AUTOINCREMENT,
                       group_name
                       TEXT,
                       email_address
                       TEXT
                   )
                   """)
    cursor.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS task_recipients (task_id INTEGER, email_address TEXT)")

    cursor.execute("PRAGMA table_info(automations)")
    cols = [col[1] for col in cursor.fetchall()]
    if "frequency" not in cols:
        cursor.execute("ALTER TABLE automations ADD COLUMN frequency TEXT DEFAULT 'Daily'")
    if "schedule_details" not in cols:
        cursor.execute("ALTER TABLE automations ADD COLUMN schedule_details TEXT DEFAULT ''")
    # Failsafe in case the previous script didn't run properly on this machine
    if "last_status" not in cols:
        cursor.execute("ALTER TABLE automations ADD COLUMN last_status TEXT")

    conn.commit()
    conn.close()


def is_valid_email(email):
    pattern = r"^[a-zA-Z0-9!#$%&'*+/=?^_`{|}~.-]+@[a-zA-Z0-9.-]+\.[a-zA-Z0-9-]{2,}$"

    if not re.match(pattern, email):
        return False

    try:
        domain_part = email.split('@')[1]
    except IndexError:
        return False

    if ".." in email or domain_part.startswith(".") or domain_part.endswith(".") or domain_part.startswith("-"):
        return False

    if "--" in domain_part:
        return False

    return True


class DispatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automated Inventory & Restock Dispatcher")
        self.root.geometry("1100x700")
        self.target_file_path = ""
        self.create_menu()
        self.build_ui()
        self.load_emails()
        self.load_automations()

    def create_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Settings", command=self.open_settings_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        self.root.config(menu=menubar)

    def open_settings_dialog(self):
        dialog = ttk.Toplevel(self.root)
        dialog.overrideredirect(True)
        dialog.geometry("400x200+{}+{}".format(self.root.winfo_x() + 300, self.root.winfo_y() + 250))
        ttk.Label(dialog, text="Sender Email:").pack(pady=(15, 0))
        email_ent = ttk.Entry(dialog)
        email_ent.pack(pady=5, padx=20, fill=X)
        conn = sqlite3.connect("automations.db")
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key='sender_email'")
        res = cursor.fetchone()
        if res: email_ent.insert(0, res[0])
        conn.close()

        def save_settings():
            email = email_ent.get()
            if not is_valid_email(email):
                messagebox.showerror("Error", "Invalid Sender Email format.")
                return
            conn = sqlite3.connect("automations.db")
            cursor = conn.cursor()
            cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('sender_email', ?)", (email,))
            conn.commit()
            conn.close()
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Save Email", bootstyle="primary", command=save_settings).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="Close", bootstyle="secondary", command=dialog.destroy).pack(side=LEFT, padx=5)

    def build_ui(self):
        left_frame = ttk.Frame(self.root, padding=15)
        left_frame.pack(side=LEFT, fill=BOTH, expand=True)
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(side=BOTTOM, fill=X)
        ttk.Button(btn_frame, text="Add Group", command=self.add_group).pack(side=LEFT, padx=2)
        ttk.Button(btn_frame, text="Add Email", command=self.add_email).pack(side=LEFT, padx=2)
        ttk.Button(btn_frame, text="Delete", bootstyle="danger", command=self.delete_selected).pack(side=RIGHT, padx=2)
        self.tree = ttk.Treeview(left_frame, show="tree")
        self.tree.pack(side=TOP, fill=BOTH, expand=True, pady=(0, 10))
        self.tree.bind("<Double-1>", self.toggle_email)
        self.tree.bind("<space>", self.toggle_email)

        right_frame = ttk.Frame(self.root, padding=15)
        right_frame.pack(side=RIGHT, fill=BOTH, expand=True)
        setup_frame = ttk.Frame(right_frame)
        setup_frame.pack(fill=X, pady=(0, 20))
        file_frame = ttk.Frame(setup_frame)
        file_frame.pack(fill=X, pady=5)
        ttk.Button(file_frame, text="Select Excel Workbook", command=self.select_file).pack(side=LEFT)
        self.lbl_filename = ttk.Label(file_frame, text="No file selected", font=("Segoe UI", 10, "italic"))
        self.lbl_filename.pack(side=LEFT, padx=15)
        time_frame = ttk.Frame(setup_frame)
        time_frame.pack(fill=X, pady=10)
        ttk.Label(time_frame, text="Schedule Time:").pack(side=LEFT)
        self.time_hr = ttk.Combobox(time_frame, values=[f"{i:02d}" for i in range(1, 13)], width=3, state="readonly")
        self.time_hr.set("08")
        self.time_hr.pack(side=LEFT, padx=(10, 2))
        ttk.Label(time_frame, text=":").pack(side=LEFT)
        self.time_min = ttk.Combobox(time_frame, values=[f"{i:02d}" for i in range(0, 60, 5)], width=3,
                                     state="readonly")
        self.time_min.set("00")
        self.time_min.pack(side=LEFT, padx=2)
        self.time_ampm = ttk.Combobox(time_frame, values=["AM", "PM"], width=4, state="readonly")
        self.time_ampm.set("AM")
        self.time_ampm.pack(side=LEFT, padx=2)
        ttk.Label(time_frame, text="Freq:").pack(side=LEFT, padx=(15, 2))
        self.freq_var = tk.StringVar(value="Daily")
        self.combo_freq = ttk.Combobox(time_frame, textvariable=self.freq_var,
                                       values=["Once", "Daily", "Weekly", "Bi-Weekly", "Monthly"], width=9,
                                       state="readonly")
        self.combo_freq.pack(side=LEFT, padx=2)
        self.freq_var.trace_add("write", self.update_dynamic_schedule)
        ttk.Label(time_frame, text="Offset (hrs):").pack(side=LEFT, padx=(15, 2))
        self.ent_offset = ttk.Entry(time_frame, width=5)
        self.ent_offset.insert(0, "1")
        self.ent_offset.pack(side=LEFT)
        self.dyn_frame = ttk.Frame(setup_frame)
        self.dyn_frame.pack(fill=X, pady=5)
        self.weekly_frame = ttk.Frame(self.dyn_frame)
        self.day_vars = {day: tk.BooleanVar() for day in ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]}
        for d in self.day_vars:
            ttk.Checkbutton(self.weekly_frame, text=d, variable=self.day_vars[d]).pack(side=LEFT, padx=5)
        self.monthly_frame = ttk.Frame(self.dyn_frame)
        ttk.Label(self.monthly_frame, text="Day of Month (1-31 or 'Last'):").pack(side=LEFT)
        self.month_day_ent = ttk.Entry(self.monthly_frame, width=5)
        self.month_day_ent.insert(0, "1")
        self.month_day_ent.pack(side=LEFT, padx=5)
        ttk.Button(setup_frame, text="Save & Schedule Task", bootstyle="success", command=self.save_and_schedule).pack(
            pady=15, fill=X)

        list_frame = ttk.Frame(right_frame)
        list_frame.pack(fill=BOTH, expand=True)
        ttk.Label(list_frame, text="Active Scheduled Tasks", font=("Segoe UI", 12, "bold")).pack(anchor=W, pady=(0, 5))
        action_btn_frame = ttk.Frame(list_frame)
        action_btn_frame.pack(side=BOTTOM, fill=X)
        ttk.Button(action_btn_frame, text="Edit Task", command=self.edit_task).pack(side=LEFT, padx=2)
        ttk.Button(action_btn_frame, text="Delete Task", bootstyle="danger", command=self.delete_task).pack(side=LEFT,
                                                                                                            padx=2)

        # Updated to include Status column
        self.auto_list = ttk.Treeview(list_frame, columns=("Time", "Freq", "Details", "File", "Status"),
                                      show="headings")
        self.auto_list.heading("Time", text="Time")
        self.auto_list.heading("Freq", text="Frequency")
        self.auto_list.heading("Details", text="Details")
        self.auto_list.heading("File", text="Target File")
        self.auto_list.heading("Status", text="Status")

        self.auto_list.column("Time", width=70, anchor=CENTER)
        self.auto_list.column("Freq", width=70, anchor=CENTER)
        self.auto_list.column("Details", width=90, anchor=CENTER)
        self.auto_list.column("File", width=150, anchor=W)
        self.auto_list.column("Status", width=220, anchor=W)

        self.auto_list.pack(side=TOP, fill=BOTH, expand=True, pady=(0, 10))

    def update_dynamic_schedule(self, *args):
        self.weekly_frame.pack_forget()
        self.monthly_frame.pack_forget()
        val = self.freq_var.get()
        if val in ["Weekly", "Bi-Weekly"]:
            self.weekly_frame.pack(side=LEFT)
        elif val == "Monthly":
            self.monthly_frame.pack(side=LEFT)

    def toggle_email(self, event):
        selected = self.tree.focus()
        if not selected or self.tree.parent(selected) == "": return
        text = self.tree.item(selected, "text")
        if text.startswith("[ ] "):
            self.tree.item(selected, text=text.replace("[ ] ", "[X] ", 1))
        elif text.startswith("[X] "):
            self.tree.item(selected, text=text.replace("[X] ", "[ ] ", 1))

    def add_group(self):
        dlg = ttk.Toplevel(self.root)
        dlg.overrideredirect(True)
        dlg.geometry("300x120+{}+{}".format(self.root.winfo_x() + 350, self.root.winfo_y() + 250))
        ttk.Label(dlg, text="Enter Department/Group Name:").pack(pady=10)
        ent = ttk.Entry(dlg)
        ent.pack(pady=5, padx=20, fill=X)
        ttk.Button(dlg, text="OK",
                   command=lambda: [self.tree.insert("", "end", text=ent.get(), open=True), dlg.destroy()]).pack(pady=5)

    def add_email(self):
        selected = self.tree.focus()
        if not selected or self.tree.parent(selected) != "":
            messagebox.showwarning("Selection", "Select a group.")
            return
        dlg = ttk.Toplevel(self.root)
        dlg.overrideredirect(True)
        dlg.geometry("300x120+{}+{}".format(self.root.winfo_x() + 350, self.root.winfo_y() + 250))
        ttk.Label(dlg, text="Enter Email Address:").pack(pady=10)
        ent = ttk.Entry(dlg)
        ent.pack(pady=5, padx=20, fill=X)

        def on_submit():
            if is_valid_email(ent.get()):
                self.tree.insert(selected, "end", text=f"[ ] {ent.get()}")
                self.save_emails_to_db()
                dlg.destroy()
            else:
                messagebox.showerror("Error", "Invalid email format.")

        ttk.Button(dlg, text="OK", command=on_submit).pack(pady=5)

    def delete_selected(self):
        selected = self.tree.focus()
        if selected:
            self.tree.delete(selected)
            self.save_emails_to_db()

    def save_emails_to_db(self):
        conn = sqlite3.connect("automations.db")
        cursor = conn.cursor()
        cursor.execute("DELETE FROM recipients")
        for group_id in self.tree.get_children():
            g_name = self.tree.item(group_id, "text")
            for e_id in self.tree.get_children(group_id):
                email_text = self.tree.item(e_id, "text")[4:]
                cursor.execute("INSERT INTO recipients (group_name, email_address) VALUES (?, ?)", (g_name, email_text))
        conn.commit()
        conn.close()

    def load_emails(self):
        for row in self.tree.get_children(): self.tree.delete(row)
        conn = sqlite3.connect("automations.db")
        cursor = conn.cursor()
        cursor.execute("SELECT group_name, email_address FROM recipients")
        rows = cursor.fetchall()
        groups = {}
        for g_name, email in rows:
            if g_name not in groups: groups[g_name] = self.tree.insert("", "end", text=g_name, open=True)
            self.tree.insert(groups[g_name], "end", text=f"[ ] {email}")
        conn.close()

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Select Inventory Workbook",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if path:
            if not path.lower().endswith(('.xlsx', '.xls')):
                messagebox.showerror("Error", "Please select a valid Excel workbook (.xlsx or .xls).")
                return

            self.target_file_path = path
            self.lbl_filename.config(text=os.path.basename(path))

    def save_and_schedule(self):
        conn = sqlite3.connect("automations.db")
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key='sender_email'")
        res = cursor.fetchone()
        sender_email = res[0] if res else ""
        if not is_valid_email(sender_email):
            messagebox.showerror("Error", "Check your Sender Email in Settings.")
            conn.close()
            return
        if not self.target_file_path:
            messagebox.showerror("Error", "Select an Excel file.")
            conn.close()
            return
        selected_emails = []
        for group_id in self.tree.get_children():
            for e_id in self.tree.get_children(group_id):
                text = self.tree.item(e_id, "text")
                if text.startswith("[X] "): selected_emails.append(text[4:])
        if not selected_emails:
            messagebox.showerror("Error", "Select at least one recipient.")
            conn.close()
            return
        hr = int(self.time_hr.get())
        mn = self.time_min.get()
        if self.time_ampm.get() == "PM" and hr != 12:
            hr += 12
        elif self.time_ampm.get() == "AM" and hr == 12:
            hr = 0
        schedule_time_24 = f"{hr:02d}:{mn}"
        base_name = os.path.basename(self.target_file_path).replace(' ', '_').split('.')[0]
        task_name = f"AutoTask_{base_name}_{hr:02d}{mn}"
        freq = self.freq_var.get()
        details = ""
        trigger_cmd = ""
        if freq == "Once":
            trigger_cmd = f"New-ScheduledTaskTrigger -Once -At {schedule_time_24}"
        elif freq == "Daily":
            trigger_cmd = f"New-ScheduledTaskTrigger -Daily -At {schedule_time_24}"
        elif freq in ["Weekly", "Bi-Weekly"]:
            selected_days = [d for d, var in self.day_vars.items() if var.get()]
            if not selected_days:
                messagebox.showerror("Error", "Select at least one day.")
                conn.close()
                return
            details = ",".join(selected_days)
            ps_day_map = {"Mon": "Monday", "Tue": "Tuesday", "Wed": "Wednesday", "Thu": "Thursday", "Fri": "Friday",
                          "Sat": "Saturday", "Sun": "Sunday"}
            ps_days = ",".join([ps_day_map[d] for d in selected_days])
            interval = "2" if freq == "Bi-Weekly" else "1"
            trigger_cmd = f"New-ScheduledTaskTrigger -Weekly -DaysOfWeek {ps_days} -WeeksInterval {interval} -At {schedule_time_24}"
        elif freq == "Monthly":
            day = self.month_day_ent.get().strip()
            if day.lower() == "last":
                details = "Last"
                trigger_cmd = f"New-ScheduledTaskTrigger -Monthly -DaysOfMonth 28,29,30,31 -At {schedule_time_24}"
            elif day.isdigit() and (1 <= int(day) <= 31):
                details = day
                trigger_cmd = f"New-ScheduledTaskTrigger -Monthly -DaysOfMonth {day} -At {schedule_time_24}"
            else:
                messagebox.showerror("Error", "Invalid Day (1-31 or 'Last').")
                conn.close()
                return
        cursor.execute(
            "INSERT INTO automations (task_name, file_path, sender_email, schedule_time, precheck_offset, frequency, schedule_details, last_status) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (task_name, self.target_file_path, sender_email, schedule_time_24, int(self.ent_offset.get()), freq,
             details, "Pending"))
        task_id = cursor.lastrowid
        for email in selected_emails: cursor.execute(
            "INSERT INTO task_recipients (task_id, email_address) VALUES (?, ?)", (task_id, email))
        conn.commit()
        conn.close()
        ps_cmd = f"Register-ScheduledTask -TaskName '{task_name}' -Action (New-ScheduledTaskAction -Execute '{sys.executable}' -Argument '\"{os.path.abspath('dispatcher.py')}\" {task_id}') -Trigger ({trigger_cmd}) -Settings (New-ScheduledTaskSettingsSet -WakeToRun) -Force"
        subprocess.run(["powershell", "-Command", ps_cmd], creationflags=subprocess.CREATE_NO_WINDOW)
        self.load_automations()
        messagebox.showinfo("Success", "Task scheduled.")

    def delete_task(self):
        selected = self.auto_list.focus()
        if not selected: return

        # Updated indices because Status column pushed hidden values back by 1
        task_info = self.auto_list.item(selected, "values")
        task_name = task_info[5]
        task_id = task_info[6]

        conn = sqlite3.connect("automations.db")
        cursor = conn.cursor()

        cursor.execute("DELETE FROM task_recipients WHERE task_id=?", (task_id,))
        cursor.execute("DELETE FROM automations WHERE id=?", (task_id,))

        conn.commit()
        conn.close()

        subprocess.run(["powershell", "-Command", f"Unregister-ScheduledTask -TaskName '{task_name}' -Confirm:$false"],
                       creationflags=subprocess.CREATE_NO_WINDOW)
        self.load_automations()

    def edit_task(self):
        selected = self.auto_list.focus()
        if not selected: return
        # Updated index to 6
        task_id = self.auto_list.item(selected, "values")[6]
        conn = sqlite3.connect("automations.db")
        cursor = conn.cursor()
        cursor.execute(
            "SELECT file_path, schedule_time, precheck_offset, frequency, schedule_details FROM automations WHERE id=?",
            (task_id,))
        row = cursor.fetchone()
        if row:
            self.target_file_path = row[0]
            self.lbl_filename.config(text=os.path.basename(row[0]))
            hr, mn = map(int, row[1].split(':'))
            self.time_hr.set(f"{(hr % 12 or 12):02d}")
            self.time_min.set(f"{mn:02d}")
            self.time_ampm.set("PM" if hr >= 12 else "AM")
            self.ent_offset.delete(0, END)
            self.ent_offset.insert(0, str(row[2]))
            if row[3]: self.freq_var.set(row[3]); self.update_dynamic_schedule()
            details = row[4]
            if details:
                if row[3] in ["Weekly", "Bi-Weekly"]:
                    for d in self.day_vars: self.day_vars[d].set(d in details.split(","))
                elif row[3] == "Monthly":
                    self.month_day_ent.delete(0, END)
                    self.month_day_ent.insert(0, details)
            for group_id in self.tree.get_children():
                for e_id in self.tree.get_children(group_id):
                    text = self.tree.item(e_id, "text")
                    if text.startswith("[X] "): self.tree.item(e_id, text=text.replace("[X] ", "[ ] ", 1))
            cursor.execute("SELECT email_address FROM task_recipients WHERE task_id=?", (task_id,))
            task_emails = [r[0] for r in cursor.fetchall()]
            for group_id in self.tree.get_children():
                for e_id in self.tree.get_children(group_id):
                    text = self.tree.item(e_id, "text")
                    email_only = text[4:]
                    if email_only in task_emails: self.tree.item(e_id, text=f"[X] {email_only}")
        conn.close()

    def load_automations(self):
        for row in self.auto_list.get_children(): self.auto_list.delete(row)
        conn = sqlite3.connect("automations.db")
        cursor = conn.cursor()

        # Updated query to pull last_status
        cursor.execute(
            "SELECT id, task_name, schedule_time, file_path, frequency, schedule_details, last_status FROM automations")
        for row in cursor.fetchall():
            task_id, task_name, stime, fpath, freq, details, status = row
            hr, mn = map(int, stime.split(':'))
            disp = f"{(hr % 12 or 12):02d}:{mn:02d} {'PM' if hr >= 12 else 'AM'}"

            # Handle empty status text
            display_status = status if status else "Not run yet"

            # Insert the row into the Treeview with the display_status
            self.auto_list.insert("", END,
                                  values=(disp, freq, details, os.path.basename(fpath), display_status, task_name,
                                          task_id))
        conn.close()


if __name__ == "__main__":
    setup_database()
    root = ttk.Window(themename="superhero")
    app = DispatcherApp(root)
    root.mainloop()