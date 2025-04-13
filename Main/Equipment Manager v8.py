import customtkinter as ctk
import sqlite3 as sql
from datetime import datetime
from pathlib import Path
from tksheet import Sheet
import csv

stu_csv_path = Path("Main") / "Students"
eq_csv_path = Path("Main") / "Equipment"

data_path = Path("Main") / ("DB") / "data.db"
data_uri = data_path.resolve().as_uri()

con = sql.connect(data_uri, uri=True)
cur = con.cursor()

ctk.set_appearance_mode("dark")

#-----Temp Data Table-----#
cur.execute("""CREATE TABLE IF NOT EXISTS temp(
            SKU TEXT PRIMARY KEY NOT NULL, 
            stu_ID INTEGER NOT NULL, 
            date TEXT NOT NULL,
            location TEXT NOT NULL
            )""")
cur.execute("DELETE FROM temp")
con.commit()

#-----Equipment Info Data Table-----#
cur.execute("""CREATE TABLE IF NOT EXISTS equipment(
            name TEXT NOT NULL,
            SKU TEXT PRIMARY KEY NOT NULL,
            category TEXT,
            kit TEXT,
            kit_parts TEXT
            )""")
con.commit()

#-----Student ID Data Table-----#
cur.execute("""CREATE TABLE IF NOT EXISTS students( 
            stu_ID TEXT PRIMARY KEY NOT NULL, 
            stu_name TEXT NOT NULL
            )""")
con.commit()

#-----Master Data Table-----#
cur.execute("""CREATE TABLE IF NOT EXISTS master(
            SKU TEXT NOT NULL, 
            name TEXT NOT NULL,
            stu_ID INTEGER NOT NULL,
            stu_name TEXT, 
            date TEXT NOT NULL,
            location TEXT NOT NULL
            )""")
con.commit()

#-----Availability Data Table-----#
cur.execute("""CREATE TABLE IF NOT EXISTS avail(
            name TEXT NOT NULL,
            location TEXT NOT NULL
            )""")
con.commit()

#-----Student ID Import Table-----#
cur.execute("""CREATE TABLE IF NOT EXISTS import_students( 
            stu_ID TEXT PRIMARY KEY NOT NULL, 
            stu_name TEXT NOT NULL
            )""")
cur.execute("DELETE FROM import_students")
con.commit()

#-----Equipment Import Table-----#
cur.execute("""CREATE TABLE IF NOT EXISTS import_equipment(
            name TEXT NOT NULL,
            SKU TEXT PRIMARY KEY NOT NULL,
            category TEXT,
            kit TEXT,
            kit_parts TEXT
            )""")
cur.execute("DELETE FROM import_equipment")
con.commit()


#-----Start Window-----#
class start_window(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.geometry("800x480")
        self.title("Equipment Manager")
        self.attributes("-fullscreen", True)
        self.state("normal")
        
        #Frames
        self.frames = {}

        self.container = ctk.CTkFrame(self)
        self.container.pack(fill="both", expand=True)

        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        for F in (
            HomeFrame, SettingsFrame, InFrame, EqFrame, StuFrame, 
            StuImportFrame, StuModifyFrame, OutFrame, AvailFrame, 
            MasterListFrame, AvailModifyFrame
            ):
            frame = F(self.container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")
            frame.grid_remove()

    def show_frame(self, frame_class):

        for frame in self.frames.values():
            frame.grid_remove()
        frame = self.frames[frame_class]
        frame.grid()
        frame.tkraise()
        self.attributes("-fullscreen", True)

        if hasattr(frame, "on_open") and callable(frame.on_open):
            frame.on_open()

    def close(self):
        self.quit()
        self.destroy()

    def error(self, message):
        error_window = ctk.CTkToplevel(self)
        error_window.geometry("300x150")
        error_window.title("ERROR")
        error_window.attributes("-topmost", 1)
        error_window.grab_set()
        error_window.grid_columnconfigure(0, weight=1)
        error_window.grid_rowconfigure(0, weight=1)

        error_window.bind("<Button-1>", lambda event: error_window.destroy())

        error_label = ctk.CTkLabel(error_window, text=message, wraplength=250, justify="center", anchor="center")
        error_label.grid(column=0, row=0)
    

#-----Home Frame-----#
class HomeFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.controller = controller
        self.grid_columnconfigure(0, weight=1)

        self.in_button = ctk.CTkButton(
            self, text="Check-in", command=lambda: self.check_press("in"), width=200, height=30)
        self.in_button.grid(row=0, column=0, padx=20, pady=(40,0))

        self.out_button = ctk.CTkButton(
            self, text="Check-out", command=lambda: self.check_press("out"), width=200, height=30)
        self.out_button.grid(row=1, column=0, padx=20, pady=(40,0))

        self.avail_button = ctk.CTkButton(
            self, text="Availability", command=lambda: controller.show_frame(AvailFrame), width=200, height=30)
        self.avail_button.grid(row=2, column=0, padx=20, pady=(40,0))

        self.settings_button = ctk.CTkButton(
            self, text="Settings", command=self.settings_press, width=200, height=30)
        self.settings_button.grid(row=3, column=0, padx=20, pady=(40,0))

    def check_press(self, action):
    
        self.ID_popup = ID_window(self, self.controller, action)

    def settings_press(self):
    
        self.password_popup = password_window(self, self.controller)


#-----Student ID Entry Window-----#
class ID_window(ctk.CTkToplevel):
    def __init__(self, parent, controller, action):
        super().__init__(parent)

        self.controller = controller
        self.parent = parent
        self.action = action
        self.title("Identity Verification")
        self.geometry("400x200")
        self.grid_columnconfigure(0, weight=1)
        self.lift()

        self.grab_set()

        self.ID_text = ctk.CTkLabel(self, text="Scan Student ID Card:")
        self.ID_text.grid(column=0 ,row=0, pady=20, columnspan=2)
    
        self.ID_entry = ctk.CTkEntry(self)
        self.ID_entry.grid(column=0 ,row=1, columnspan=2)
        self.ID_entry.bind("<Return>", lambda event: self.ID_submit())
        self.after(100, self.ID_entry.focus_set)

        self.ID_submit_button = ctk.CTkButton(
            self, text="Submit", command=self.ID_submit, width=100, height=30)
        self.ID_submit_button.grid(column=1 ,row=2, padx=(0,60), pady=(30,20), sticky="w")

        self.ID_cancel_button = ctk.CTkButton(
            self, text="Cancel", command=self.cancel, width=100, height=30)
        self.ID_cancel_button.grid(column=0 ,row=2, padx=(60,0), pady=(30,20), sticky="w")

    def ID_submit(self):

        enter_ID = self.ID_entry.get()
        cur.execute("SELECT stu_name FROM students WHERE stu_ID = ?", (enter_ID,))
        stu_ID = cur.fetchone()
        
        if stu_ID:

            global read_stu_ID
            read_stu_ID = self.ID_entry.get()

            if self.action == "in":
                self.controller.show_frame(InFrame)
            if self.action == "out":
                self.controller.show_frame(OutFrame)

            self.destroy()

        else:

            self.controller.error("No matching student found.")

    def cancel(self):

        self.destroy()


#-----Equipment Check-in Frame-----#
class InFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.controller = controller
        self.grid_columnconfigure(0, weight=1)

        self.in_tag_entry = ctk.CTkEntry(self, width=140, height=30)
        self.in_tag_entry.grid(column=0, row=0, pady=(20,0))
        self.in_tag_entry.bind("<Return>", self.in_scan_barcode)

        self.in_list = ctk.CTkTextbox(self, height=300, width=300, state="disabled")
        self.in_list.grid(row=1, column=0, pady=(40,0))

        self.in_tag_button = ctk.CTkButton(self, text="Done", command=self.in_tag_press, width=200, height=30)
        self.in_tag_button.grid(row=2, column=0, pady=(40,0))
    
    def on_open(self):

        self.in_tag_entry.focus_set()

    def in_tag_press(self):
        
        cur.execute("""INSERT INTO master (SKU, name, stu_ID, stu_name, date, location) 
        SELECT temp.SKU, equipment.name, temp.stu_ID, students.stu_name, temp.date, temp.location 
        FROM temp
        JOIN equipment ON temp.SKU = equipment.SKU
        JOIN students ON temp.stu_ID = students.stu_ID""")
        
        cur.execute("""UPDATE avail SET location = "in" WHERE name IN
            (SELECT name FROM equipment WHERE SKU IN (SELECT SKU FROM temp))""")
        
        con.commit()

        self.controller.show_frame(HomeFrame)
        
    def in_update_list(self):

        current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur.execute("INSERT INTO temp (SKU, stu_ID, date, location) VALUES (?, ?, ?, ?)", (self.barcode, read_stu_ID, current_date, "in"))
        con.commit()
        
        self.in_list.configure(state="normal")
        cur.execute("SELECT SKU FROM temp ORDER BY date DESC LIMIT 1")
        row = cur.fetchone()
        bar_text = row[0] if row else None

        cur.execute("SELECT name FROM equipment WHERE SKU = ?", (bar_text,))
        row = cur.fetchone()

        if row:  
            name_text = str(row[0])
            self.in_list.insert("end", f"{name_text}\n")

        self.in_list.configure(state="disabled")

    def in_scan_barcode(self, event):

        self.barcode = self.in_tag_entry.get().strip()
        if self.barcode:
            
            cur.execute("SELECT SKU FROM temp WHERE SKU = ?", (self.barcode,))
            past_bar = cur.fetchall()

            if past_bar:

                self.controller.error("Same barcode scanned.")

            else:
                
                cur.execute("SELECT name FROM equipment WHERE SKU = ?", (self.barcode,))
                past_name = cur.fetchone()

                if past_name:
                    
                    cur.execute("SELECT location from avail WHERE name = ?", (past_name))
                    location = cur.fetchone()

                    if location and location[0] == "out":

                        self.kit_scan()
                    
                    else:
                        self.controller.error(f"Item {past_name[0]} is not checked out.")
                
                else:
                    self.controller.error("No equipment found.")

        self.in_tag_entry.delete(0, "end")

    def kit_scan(self):

        cur.execute("SELECT kit FROM equipment WHERE SKU = ?", (self.barcode,))
        kit_type = cur.fetchone()

        if kit_type and kit_type[0]:

            kit_type = kit_type[0].lower()
            
            if kit_type == "kmain":
                self.kit_in_popup = kit_window(self, self.controller, self.barcode, "in")
            elif kit_type == "kpart":
                self.controller.error("Scanned item is a part of a kit.")
            else:
                self.controller.error(f"Unknown kit type: {kit_type}.")
        else:
            
            self.in_update_list()


#-----Equipment Check-out Frame-----#
class OutFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.controller = controller
        self.grid_columnconfigure(0, weight=1)

        self.in_tag_entry = ctk.CTkEntry(self, width=140, height=30)
        self.in_tag_entry.grid(column=0, row=0, pady=(20,0))
        self.in_tag_entry.bind("<Return>", self.in_scan_barcode)

        self.in_list = ctk.CTkTextbox(self, height=300, width=300, state="disabled")
        self.in_list.grid(row=1, column=0, pady=(40,0))

        self.in_tag_button = ctk.CTkButton(self, text="Done", command=self.in_tag_press, width=200, height=30)
        self.in_tag_button.grid(row=2, column=0, pady=(40,0))
    
    def on_open(self):

        self.in_tag_entry.focus_set()

    def in_tag_press(self):
        
        cur.execute("""INSERT INTO master (SKU, name, stu_ID, stu_name, date, location) 
        SELECT temp.SKU, equipment.name, temp.stu_ID, students.stu_name, temp.date, temp.location 
        FROM temp
        JOIN equipment ON temp.SKU = equipment.SKU
        JOIN students ON temp.stu_ID = students.stu_ID""")
        
        cur.execute("""UPDATE avail SET location = "out" WHERE name IN
            (SELECT name FROM equipment WHERE SKU IN (SELECT SKU FROM temp))""")
        
        con.commit()

        self.controller.show_frame(HomeFrame)
        
    def in_update_list(self):

        current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur.execute("INSERT INTO temp (SKU, stu_ID, date, location) VALUES (?, ?, ?, ?)", (self.barcode, read_stu_ID, current_date, "out"))
        con.commit()

        self.in_list.configure(state="normal")
        cur.execute("SELECT SKU FROM temp ORDER BY date DESC LIMIT 1")
        row = cur.fetchone()
        bar_text = row[0] if row else None

        cur.execute("SELECT name FROM equipment WHERE SKU = ?", (bar_text,))
        row = cur.fetchone()

        if row:  
            name_text = str(row[0])
            self.in_list.insert("end", f"{name_text}\n")

        self.in_list.configure(state="disabled")

    def in_scan_barcode(self, event):

        self.barcode = self.in_tag_entry.get().strip()
        if self.barcode:
            
            cur.execute("SELECT SKU FROM temp WHERE SKU = ?", (self.barcode,))
            past_bar = cur.fetchall()

            if past_bar:

                self.controller.error("Same barcode scanned.")

            else:
                
                cur.execute("SELECT name FROM equipment WHERE SKU = ?", (self.barcode,))
                past_name = cur.fetchone()

                if past_name:
                    
                    cur.execute("SELECT location from avail WHERE name = ?", (past_name))
                    location = cur.fetchone()

                    if location and location[0] == "in":

                        self.kit_scan()
                    
                    else:
                        self.controller.error(f"Item {past_name[0]} is not checked in.")
                
                else:
                    self.controller.error("No equipment found.")

        self.in_tag_entry.delete(0, "end")

    def kit_scan(self):

        cur.execute("SELECT kit FROM equipment WHERE SKU = ?", (self.barcode,))
        kit_type = cur.fetchone()

        if kit_type and kit_type[0]:

            kit_type = kit_type[0].lower()
            
            if kit_type == "kmain":
                self.kit_in_popup = kit_window(self, self.controller, self.barcode, "out")
            elif kit_type == "kpart":
                self.controller.error("Scanned item is a part of a kit.")
            else:
                self.controller.error(f"Unknown kit type: {kit_type}.")
        else:
            self.in_update_list()


#-----Kit Window-----#
class kit_window(ctk.CTkToplevel):
    def __init__(self, parent, controller, barcode, action):
        super().__init__(parent)

        self.controller = controller
        self.parent = parent
        self.barcode = barcode
        self.action = action
        self.title("Kit Check-in" if action == "in" else "Kit Check-out")
        self.geometry("400x300")
        self.grid_columnconfigure(0, weight=1)
        self.lift()

        self.grab_set()

        self.in_tag_entry = ctk.CTkEntry(self, width=140, height=30)
        self.in_tag_entry.grid(column=0, row=0, pady=(20,0))
        self.in_tag_entry.bind("<Return>", self.kit_scan)
        self.after(100, self.in_tag_entry.focus_set)

        self.text = ctk.CTkLabel(self, text="Scan in the following items:")
        self.text.grid(column=0 ,row=1, pady=(20,0), columnspan=2)

        self.in_list = ctk.CTkTextbox(self, height=100, width=300, state="disabled")
        self.in_list.grid(row=2, column=0, pady=(20,0))

        self.in_tag_button = ctk.CTkButton(self, text="Back", command=self.destroy, width=200, height=30)
        self.in_tag_button.grid(row=3, column=0, pady=(20,0))

        self.print_parts()
        

    def print_parts(self):

        cur.execute("SELECT kit_parts FROM equipment WHERE SKU = ?", (self.barcode,))
        self.kit_parts = cur.fetchone()
        
        if self.kit_parts and self.kit_parts[0]: 

            parts_list = self.kit_parts[0].split(",")

            for part_sku in parts_list:
                part_sku = part_sku.strip()
                cur.execute("SELECT name FROM equipment WHERE SKU = ?", (part_sku,))
                part_name = cur.fetchone()

                if part_name and part_name[0]:
                    self.in_list.configure(state="normal")
                    self.in_list.insert("end", f"{part_name[0]}\n")
                    self.in_list.configure(state="disabled")
                else:
                    self.controller.error(f"Unknown part SKU: {part_sku}.")

        else:

            self.controller.error("No kit parts found for kit.")
            self.destroy()

    def kit_scan(self, event):

        self.barcode = self.in_tag_entry.get().strip()

        if self.barcode:

            if self.kit_parts and self.kit_parts[0]:

                parts_list = self.kit_parts[0].split(",")

                if self.barcode in parts_list:
                    self.remove_part_from_list(self.barcode)
                else:
                    self.controller.error(f"Barcode {self.barcode} is not part of the kit.")
            else:
                self.controller.error("No kit parts found for the given SKU.")

        if self.in_tag_entry.winfo_exists():
            self.in_tag_entry.delete(0, "end")

    def remove_part_from_list(self, barcode):

        cur.execute("SELECT name FROM equipment WHERE SKU = ?", (barcode,))
        part_name = cur.fetchone()

        if part_name and part_name[0]:

            part_name = part_name[0]

            self.in_list.configure(state="normal")
            lines = self.in_list.get("1.0", "end").splitlines()
            self.in_list.delete("1.0", "end")

            for line in lines:
                if line.strip() != part_name:
                    self.in_list.insert("end", f"{line}\n")

            self.in_list.configure(state="disabled")

            remaining_lines = self.in_list.get("1.0", "end").strip()
            if not remaining_lines:
                self.close_window()
            
        else:
            self.controller.error(f"ERROR: No name found for barcode {barcode}.")

    def close_window(self):

        if self.action == "in":
            in_frame = self.controller.frames[InFrame]
            in_frame.in_update_list()
        elif self.action == "out":
            out_frame = self.controller.frames[OutFrame]
            out_frame.in_update_list()

        self.destroy()


#-----Settings Password Window-----#
class password_window(ctk.CTkToplevel):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.controller = controller
        self.parent = parent
        self.title("Admin Entry Only")
        self.geometry("400x200")
        self.grid_columnconfigure(0, weight=1)
        self.lift()

        self.grab_set()

        self.text = ctk.CTkLabel(self, text="Input Password:")
        self.text.grid(column=0 ,row=0, pady=20, columnspan=2)
    
        self.entry = ctk.CTkEntry(self)
        self.entry.grid(column=0 ,row=1, columnspan=2)
        self.entry.bind("<Return>", lambda event: self.submit())
        self.after(100, self.entry.focus_set)

        self.submit_button = ctk.CTkButton(
            self, text="Submit", command=self.submit, width=100, height=30)
        self.submit_button.grid(column=1 ,row=2, padx=(0,60), pady=(30,20), sticky="w")

        self.cancel_button = ctk.CTkButton(
            self, text="Cancel", command=self.cancel, width=100, height=30)
        self.cancel_button.grid(column=0 ,row=2, padx=(60,0), pady=(30,20), sticky="w")

    def submit(self):

        if self.entry.get() == "8273":

            self.controller.show_frame(SettingsFrame)
            self.destroy()

    def cancel(self):

        self.destroy()


#-----Availibility Frame-----#
class AvailFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.controller = controller

        self.frame = ctk.CTkFrame(self)
        self.frame.grid(row=0, column=0, padx=(20,20), pady=(20,0))
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_rowconfigure(0, weight=1)

        self.cancel_button = ctk.CTkButton(
            self, text="Back", command=lambda: controller.show_frame(HomeFrame), width=200, height=30)
        self.cancel_button.grid(row=1, pady=(20,20))

    def on_open(self):

        self.create_table()

    def create_table(self):

        if hasattr(self, "sheet") and self.sheet.winfo_exists():
            self.sheet.destroy()

        cur.execute("SELECT name, location FROM avail ORDER BY name ASC")
        list = cur.fetchall()

        name = [item[0] for item in list]
        location = [item[1] for item in list]

        data = [[name, location] for name, location in zip(name, location)]

        self.sheet = Sheet(self.frame,
                           data=data,
                           headers=["Name", "Location"],
                           width=760,
                           height=410,
                           theme="dark blue")
        self.sheet.disable_bindings()
        self.sheet.set_all_cell_sizes_to_text()
        self.frame.grid(row=0, column=0, sticky="nswe")
        self.sheet.grid(row=0, column=0, sticky="nswe")


#-----Settings Frame-----#
class SettingsFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        
        self.controller = controller
        self.grid_columnconfigure(0, weight=1)

        self.modify_item_button = ctk.CTkButton(
            self, text="Equipment Import", command=lambda: controller.show_frame(EqFrame), width=200, height=30)
        self.modify_item_button.grid(row=0, column=0, pady=(40,0))

        self.student_list_button = ctk.CTkButton(
            self, text="Student List", command=lambda: controller.show_frame(StuFrame), width=200, height=30)
        self.student_list_button.grid(row=1, column=0, pady=(40,0))

        self.master_list_button = ctk.CTkButton(
            self, text="Master List", command=lambda: controller.show_frame(MasterListFrame), width=200, height=30)
        self.master_list_button.grid(row=2, column=0, pady=(40,0))

        self.avail_modify_button = ctk.CTkButton(
            self, text="Modify Availability", command=lambda: controller.show_frame(AvailModifyFrame), width=200, height=30)
        self.avail_modify_button.grid(row=3, column=0, pady=(40,0))

        self.back_button = ctk.CTkButton(
            self, text="Back", command=lambda: controller.show_frame(HomeFrame), width=200, height=30)
        self.back_button.grid(row=4, column=0, pady=(40,0))

        self.quit_button = ctk.CTkButton(
            self, text="Quit", command=self.controller.close, width=200, height=30)
        self.quit_button.grid(row=5, column=0, pady=(40,0))
        

#-----Equipment Frame-----#
class EqFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.controller = controller
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.frame = ctk.CTkFrame(self)
        self.frame.grid(row=0, column=1, padx=(20,20), pady=(20,0))
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_rowconfigure(0, weight=1)

        self.cancel_button = ctk.CTkButton(
            self, text="Cancel", command=lambda: controller.show_frame(SettingsFrame), width=200, height=30)
        self.cancel_button.grid(row=1, pady=(20,20), padx=(70,0), sticky="w")
        self.save_button = ctk.CTkButton(
            self, text="Save Changes", command=self.save, width=200, height=30)
        self.save_button.grid(row=1, pady=(20,20), padx=(0,70), sticky="e")

    def on_open(self):
        
        self.import_csv()
        self.create_table()
    
    def import_csv(self):
        
        cur.execute("DELETE FROM import_equipment")
        con.commit()
        for eq_csv in eq_csv_path.glob("*.csv"):
            with open(eq_csv, mode='r', newline='', encoding='utf-8') as file:
                reader = csv.reader(file)
                
                for _ in range(19):
                    next(reader, None)
                
                for row in reader:
                    csv_name = row[0].strip()
                    csv_category = row[4].strip()
                    csv_kit = row[5].strip()
                    csv_SKU = row[6].strip()
                    csv_kit_parts = ",".join([part.strip() for part in row[12:] if part.strip()])

                    if not (csv_name or csv_SKU or csv_category or csv_kit or csv_kit_parts):
                        continue

                    cur.execute("""
                        INSERT OR IGNORE INTO import_equipment 
                        (name, SKU, category, kit, kit_parts) 
                        VALUES (?, ?, ?, ?, ?)""", 
                        (csv_name, csv_SKU, csv_category, csv_kit, csv_kit_parts))
                con.commit()

    def create_table(self):

        if hasattr(self, "sheet") and self.sheet.winfo_exists():
            self.sheet.destroy()

        cur.execute("SELECT name, SKU, category, kit, kit_parts FROM import_equipment ORDER BY SKU ASC")
        list = cur.fetchall()

        name = [equipment[0] for equipment in list]
        SKU = [equipment[1] for equipment in list]
        category = [equipment[2] for equipment in list]
        kit = [equipment[3] for equipment in list]
        kit_parts = [equipment[4] for equipment in list]

        data = [[name, SKU, category, kit, kit_parts]
            for name, SKU, category, kit, kit_parts in 
            zip(name, SKU, category, kit, kit_parts)]

        self.sheet = Sheet(self.frame,
                           data=data,
                           headers=["Name", "SKU", "Category", "Kit", "Kit Parts"],
                           width=760,
                           height=410,
                           theme="dark blue")
        self.sheet.enable_bindings()
        self.sheet.set_all_cell_sizes_to_text()
        self.frame.grid(row=0, column=0, sticky="nswe")
        self.sheet.grid(row=0, column=0, sticky="nswe")

    def save(self):

        cur.execute("DELETE FROM equipment")
        cur.execute("DELETE FROM avail")
        con.commit()
        
        new_data = self.sheet["A1"].expand().data

        new_name = [equipment[0] for equipment in new_data]
        new_SKU = [equipment[1] for equipment in new_data]
        new_category = [equipment[2] for equipment in new_data]
        new_kit = [equipment[3] for equipment in new_data]
        new_kit_parts = [equipment[4] for equipment in new_data]

        for name, SKU, category, kit, kit_parts in zip(new_name, new_SKU, new_category, new_kit, new_kit_parts):

            cur.execute("""INSERT INTO equipment (name, SKU, category, kit, kit_parts) 
                VALUES (?, ?, ?, ?, ?)""", (name, SKU, category, kit, kit_parts))
            cur.execute("DELETE FROM import_equipment")

            cur.execute("INSERT INTO avail (name, location) VALUES (?, ?)", (name, "in"))

        con.commit()

        self.controller.show_frame(SettingsFrame)


#-----Student Frame-----#
class StuFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.controller = controller
        self.grid_columnconfigure(0, weight=1)

        self.import_button = ctk.CTkButton(
            self, text="Import Students", command=lambda: controller.show_frame(StuImportFrame), width=200, height=30)
        self.import_button.grid(row=0, column=0, pady=(40,0))

        self.modify_button = ctk.CTkButton(
            self, text="Modify Students", command=lambda: controller.show_frame(StuModifyFrame), width=200, height=30)
        self.modify_button.grid(row=1, column=0, pady=(40,0))

        self.cancel_button = ctk.CTkButton(
            self, text="Back", command=lambda: controller.show_frame(SettingsFrame), width=200, height=30)
        self.cancel_button.grid(row=2, column=0, pady=(40,0))


#-----Student Import Frame-----#
class StuImportFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.controller = controller
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.frame = ctk.CTkFrame(self)
        self.frame.grid(row=0, column=1, padx=(20,20), pady=(20,0))
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_rowconfigure(0, weight=1)

        self.cancel_button = ctk.CTkButton(
            self, text="Cancel", command=lambda: controller.show_frame(StuFrame), width=200, height=30)
        self.cancel_button.grid(row=1, pady=(20,20), padx=(70,0), sticky="w")
        self.save_button = ctk.CTkButton(
            self, text="Save Changes", command=self.save, width=200, height=30)
        self.save_button.grid(row=1, pady=(20,20), padx=(0,70), sticky="e")
        self.check_state = ctk.StringVar(self, value="off")
        self.del_prev_stu = ctk.CTkCheckBox(
            self, text="Delete Previous Student List", variable=self.check_state, onvalue="on", offvalue="off")
        self.del_prev_stu.grid(row=1)

    def on_open(self):
        
        self.import_csv()
        self.create_table()
    
    def import_csv(self):
        
        cur.execute("DELETE FROM import_students")
        con.commit()
        for stu_csv in stu_csv_path.glob("*.csv"):
            with open(stu_csv, mode='r', newline='', encoding='utf-8') as file:
                reader = csv.reader(file)
                
                for _ in range(7):
                    next(reader, None)
                
                for row in reader:
                    if row and row[0] and row[1]:
                        csv_name = row[0].strip()
                        csv_id = row[1].strip()
                
                        cur.execute("INSERT OR IGNORE INTO import_students (stu_ID, stu_name) VALUES (?, ?)", (csv_id, csv_name))
                con.commit()

    def create_table(self):

        if hasattr(self, "sheet") and self.sheet.winfo_exists():
            self.sheet.destroy()

        cur.execute("SELECT stu_name, stu_ID FROM import_students ORDER BY stu_name ASC")
        list = cur.fetchall()

        name = [student[0] for student in list]
        ID = [student[1] for student in list]

        data = [[name, ID] for name, ID in zip(name, ID)]

        self.sheet = Sheet(self.frame,
                           data=data,
                           headers=["Student Name", "ID"],
                           width=760,
                           height=410,
                           theme="dark blue")
        self.sheet.enable_bindings()
        self.sheet.set_all_cell_sizes_to_text()
        self.frame.grid(row=0, column=0, sticky="nswe")
        self.sheet.grid(row=0, column=0, sticky="nswe")

    def save(self):

        if self.check_state.get() == "on":
            cur.execute("DELETE FROM students")
            con.commit()
        
        new_data = self.sheet["A1"].expand().data

        new_name = [student[0] for student in new_data]
        new_ID = [student[1] for student in new_data]

        for stu_id, stu_name in zip(new_ID, new_name):
            cur.execute("INSERT INTO students (stu_ID, stu_name) VALUES (?, ?)", (stu_id, stu_name))
            cur.execute("DELETE FROM import_students")
        con.commit()

        self.controller.show_frame(SettingsFrame)


#-----Student Modify Frame-----#
class StuModifyFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.controller = controller

        self.frame = ctk.CTkFrame(self)
        self.frame.grid(row=0, column=1, padx=(20,20), pady=(20,0))
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_rowconfigure(0, weight=1)

        self.cancel_button = ctk.CTkButton(
            self, text="Cancel", command=lambda: controller.show_frame(StuFrame), width=200, height=30)
        self.cancel_button.grid(row=1, pady=(20,20), padx=(70,0), sticky="w")
        self.save_button = ctk.CTkButton(
            self, text="Save Changes", command=self.save, width=200, height=30)
        self.save_button.grid(row=1, pady=(20,20), padx=(0,70), sticky="e")

    def save(self):

        new_data = self.sheet["A1"].expand().data

        new_name = [student[0] for student in new_data]
        new_ID = [student[1] for student in new_data]

        cur.execute("DELETE FROM students")
        for stu_id, stu_name in zip(new_ID, new_name):
            cur.execute("INSERT INTO students (stu_ID, stu_name) VALUES (?, ?)", (stu_id, stu_name))
        con.commit()

        self.controller.show_frame(SettingsFrame)

    def on_open(self):

        self.create_table()

    def create_table(self):

        if hasattr(self, "sheet") and self.sheet.winfo_exists():
            self.sheet.destroy()

        cur.execute("SELECT stu_name, stu_ID FROM students ORDER BY stu_name ASC")
        list = cur.fetchall()

        name = [student[0] for student in list]
        ID = [student[1] for student in list]

        data = [[name, ID] for name, ID in zip(name, ID)]

        self.sheet = Sheet(self.frame,
                           data=data,
                           headers=["Student Name", "ID"],
                           width=760,
                           height=410,
                           theme="dark blue")
        self.sheet.enable_bindings()
        self.sheet.set_all_cell_sizes_to_text()
        self.frame.grid(row=0, column=0, sticky="nswe")
        self.sheet.grid(row=0, column=0, sticky="nswe")


#-----Master List Frame-----#
class MasterListFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.controller = controller

        self.frame = ctk.CTkFrame(self)
        self.frame.grid(row=0, column=0, padx=(20,20), pady=(20,0))
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_rowconfigure(0, weight=1)

        self.cancel_button = ctk.CTkButton(
            self, text="Back", command=lambda: controller.show_frame(SettingsFrame), width=200, height=30)
        self.cancel_button.grid(row=1, pady=(20,20))

    def on_open(self):

        self.create_table()

    def create_table(self):

        if hasattr(self, "sheet") and self.sheet.winfo_exists():
            self.sheet.destroy()

        cur.execute("SELECT SKU, name, stu_ID, stu_name, date, location FROM master ORDER BY date DESC")
        list = cur.fetchall()

        SKU = [item[0] for item in list]
        name = [item[1] for item in list]
        stu_ID = [item[2] for item in list]
        stu_name = [item[3] for item in list]
        date = [item[4] for item in list]
        location = [item[5] for item in list]

        data = [[SKU, name, stu_ID, stu_name, date, location]
            for SKU, name, stu_ID, stu_name, date, location
            in zip(SKU, name, stu_ID, stu_name, date, location)]

        self.sheet = Sheet(self.frame,
                           data=data,
                           headers=["SKU", "Equipment Name", "Student ID", "Student Name", "Date", "Location"],
                           width=760,
                           height=410,
                           theme="dark blue")
        self.sheet.disable_bindings()
        self.sheet.set_all_cell_sizes_to_text()
        self.frame.grid(row=0, column=0, sticky="nswe")
        self.sheet.grid(row=0, column=0, sticky="nswe")


#-----Availibility Modify Frame-----#
class AvailModifyFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.controller = controller

        self.frame = ctk.CTkFrame(self)
        self.frame.grid(row=0, column=1, padx=(20,20), pady=(20,0))
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_rowconfigure(0, weight=1)

        self.cancel_button = ctk.CTkButton(
            self, text="Cancel", command=lambda: controller.show_frame(SettingsFrame), width=200, height=30)
        self.cancel_button.grid(row=1, pady=(20,20), padx=(70,0), sticky="w")
        self.save_button = ctk.CTkButton(
            self, text="Save Changes", command=self.save, width=200, height=30)
        self.save_button.grid(row=1, pady=(20,20), padx=(0,70), sticky="e")

    def on_open(self):

        self.create_table()
    
    def save(self):

        new_data = self.sheet["A1"].expand().data

        new_name = [item[0] for item in new_data]
        new_location = [item[1] for item in new_data]

        cur.execute("DELETE FROM avail")
        for name, location in zip(new_name, new_location):
            cur.execute("INSERT INTO avail (name, location) VALUES (?, ?)", (name, location))
        con.commit()

        self.controller.show_frame(SettingsFrame)

    def create_table(self):

        if hasattr(self, "sheet") and self.sheet.winfo_exists():
            self.sheet.destroy()

        cur.execute("SELECT name, location FROM avail ORDER BY name ASC")
        list = cur.fetchall()

        name = [item[0] for item in list]
        location = [item[1] for item in list]

        data = [[name, location] for name, location in zip(name, location)]

        self.sheet = Sheet(self.frame,
                           data=data,
                           headers=["Name", "Location"],
                           width=760,
                           height=410,
                           theme="dark blue")
        self.sheet.enable_bindings()
        self.sheet.set_all_cell_sizes_to_text()
        self.frame.grid(row=0, column=0, sticky="nswe")
        self.sheet.grid(row=0, column=0, sticky="nswe")


#-----Start Program-----#
if __name__ == "__main__":
    start_app = start_window()
    start_app.show_frame(HomeFrame)
    start_app.mainloop()
    con.close()