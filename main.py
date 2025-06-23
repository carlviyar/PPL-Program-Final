# Python program to create a basic GUI 
# application using the customtkinter module

import customtkinter as ctk
from tkinter import messagebox
import tkinter as tk
import excel
from datetime import date
import os

ctk.set_appearance_mode("System") 
ctk.set_default_color_theme("blue") 
APP_WIDTH, APP_HEIGHT = 900, 900
NUM_CUSTOM_LABS = 5

def calculate_sum(labs):
    sum = 0
    for lab in labs:
        sum += lab["price"]
    return sum


# App Class
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        labs, lab_info = excel.import_excel_data()
        num_columns = len(labs)

        self.title("PPL Invoice Form")
        self.geometry(f"{APP_HEIGHT}x{APP_WIDTH}")
        self.grid_columnconfigure(list(range(num_columns)), weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.lab_info = lab_info

        self.name_frame = NameFrame(self)
        self.name_frame.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="nswe", columnspan=num_columns)

        self.labs_frame = ScrollableFrame(self, num_columns, labs)
        self.labs_frame.grid(row=1, column=0, padx=10, pady=(5, 5), sticky="nswe", columnspan=num_columns)

        self.button = ctk.CTkButton(self, text="Create Invoice", command=self.button_callback)
        self.button.grid(row=3, column=0, padx=10, pady=10, sticky="ew", columnspan=num_columns)

    def button_callback(self):
        namebox = self.name_frame.get()
        checked_lab_codes, custom_labs = self.labs_frame.get()

        all_checked_labs = [self.lab_info[int(checked_lab_code)] for checked_lab_code in checked_lab_codes]
        all_checked_labs += custom_labs

        total_price = calculate_sum(all_checked_labs)
        # For base price
        total_price += excel.get_base_price()
        lab_count = len(all_checked_labs)

        # test to see that file exists already
        file_name = excel.output_file_name(namebox["name"])
        for root, dirs, files in os.walk("."):
            for file in files:
                # file_name[2:] because the first two chars are "./"
                if file == file_name[2:]:
                    file_exists = messagebox.Message(
                        title="File already exists",
                        message=f"The file \"{file}\" already exists in the current directory. This program will overwrite that existing file.",
                        detail=f"Continue with saving the new file?",
                        type=messagebox.OKCANCEL
                        )
                    res = file_exists.show()
                    if res == messagebox.CANCEL:
                        return
        
        alert = messagebox.Message(
            title="Finalize Invoice",
            message=f"Number of tests ordered: {lab_count}\nTotal cost: {total_price}\nPrepaid?: {namebox["prepaid"]}",
            detail=f"For patient {namebox["name"]} on {date.today().strftime("%m/%d/%y")}",
            type=messagebox.OKCANCEL
            )
    
        res = alert.show()
        if res == messagebox.OK:
            excel.create_workbook({
                "name": namebox["name"],
                "notes": namebox["notes"],
                "ordered_labs": all_checked_labs,
                "prepaid": namebox["prepaid"]
            }) 
            self.reset()
            # sys.exit()
    
    def reset(self):
        # set name to empty, prepaid to ON, and notes to empty
        self.name_frame.name_box.delete(0, ctk.END)
        self.name_frame.prepaid.select()
        self.name_frame.notes_box.delete("0.0", ctk.END)

        # empty out all labs
        # lab: MyCheckboxFrame
        for lab in self.labs_frame.checkbox_frames:
            # checkbox: CTkCheckBox
            for checkbox in lab.checkboxes:
                checkbox.deselect()

        # lab: CustomLabFrame
        for lab in self.labs_frame.custom_labs_frame.custom_lab_frames:
            lab.test_code_box.delete(0, ctk.END)
            lab.desc_box.delete(0, ctk.END)
            lab.price_box.delete(0, ctk.END)

class ScrollableFrame(ctk.CTkScrollableFrame):
    def __init__(self, master, num_columns, labs):
        super().__init__(master)
        self.grid_columnconfigure(list(range(num_columns)), weight=1, minsize=275)
        self.grid_rowconfigure(0, weight=1)
        
        self.checkbox_frames = []
        for i, labs_column in enumerate(labs):
            new_cbox_frame = MyCheckboxFrame(self, values=labs_column)
            new_cbox_frame.grid(row=0, column=i, padx=5, pady=(0, 10), sticky="nsew")
            self.checkbox_frames.append(new_cbox_frame)

        self.custom_labs_frame = CustomLabsFrame(self)
        self.custom_labs_frame.grid(row=1, column=0, padx=10, pady=(15, 5), sticky="nswe", columnspan=num_columns)
    
    def get(self):
        checked_lab_codes = []
        for i, frame in enumerate(self.checkbox_frames):
            frame_checked_lab_codes = frame.get()
            checked_lab_codes += frame_checked_lab_codes
        
        custom_labs = self.custom_labs_frame.get()

        if len(checked_lab_codes + custom_labs) == 0:
            messagebox.showwarning("No labs", "Warning: There are no labs selected or manually inputted.")
        
        return checked_lab_codes, custom_labs

class CustomLabFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        self.grid_columnconfigure(3, weight=1)
        self.test_code_label = ctk.CTkLabel(self, text="Test Code")
        self.test_code_label.grid(row=0, column=0, padx=10, pady=(5, 5))
        self.test_code_box = ctk.CTkEntry(self, placeholder_text="Test Code")
        self.test_code_box.grid(row=0, column=1, padx=(0, 10), pady=(5, 5))
        self.desc_label = ctk.CTkLabel(self, text="Description")
        self.desc_label.grid(row=0, column=2,  padx=10, pady=(5, 5))
        self.desc_box = ctk.CTkEntry(self, placeholder_text="Description")
        self.desc_box.grid(row=0, column=3, sticky="nsew",  padx=10, pady=(5, 5))
        self.price_label = ctk.CTkLabel(self, text="Price")
        self.price_label.grid(row=0, column=4, padx=10, pady=(5, 5))
        self.price_box = ctk.CTkEntry(self, placeholder_text="Price")
        self.price_box.grid(row=0, column=5, sticky="nse", padx=10, pady=(5, 5))

    def get(self):
        return {
            "test_code": self.test_code_box.get(),
            "description": self.desc_box.get(),
            "price": self.price_box.get()
        }


class CustomLabsFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        self.grid_rowconfigure(list(range(NUM_CUSTOM_LABS)), weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.custom_lab_frames = []

        for i in range(NUM_CUSTOM_LABS):
            new_lab_frame = CustomLabFrame(self)
            new_lab_frame.grid(row=i, column=0, padx=0, pady=(5, 0), sticky="nsew")
            self.custom_lab_frames.append(new_lab_frame)
    
    def get(self):
        custom_labs = []
        for lab_frame in self.custom_lab_frames:
            lab = lab_frame.get()
            if lab["test_code"] == '' and lab["description"] == '' and lab["price"] == '':
                continue
            try:
                float_price = float(lab["price"])
                print(float_price)
            except:
                messagebox.showerror(
                    title="Error with custom labs",
                    message=f"The program was unable to convert the price of the following lab into a number. Please double-check your custom labs!\n\n{lab["test_code"]}: {lab["description"]}, ${lab["price"]}",
                )
                return
            lab["price"] = float_price
            custom_labs.append(lab)
        return custom_labs


class NameFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)

        self.grid_columnconfigure(1, weight=1)
        self.name_label = ctk.CTkLabel(self, text="Name")
        self.name_label.grid(row=0, column=0, sticky="nsw", padx=10, pady=(10, 0))
        self.name_box = ctk.CTkEntry(self, placeholder_text="Jane Doe", width=600)
        self.name_box.grid(row=0, column=1, sticky="nsw", padx=(0, 10), pady=(10, 0), columnspan=2)
        self.prepaid = ctk.CTkCheckBox(self, text="Prepaid?")
        self.prepaid.grid(row=0, column=2, sticky="nse", padx=10, pady=(10, 0))
        self.prepaid.select()

        self.notes_label = ctk.CTkLabel(self, text="Notes")
        self.notes_label.grid(row=1, column=0, sticky="nsw", padx=10, pady=(5, 10))
        self.notes_box = ctk.CTkTextbox(self, height=50)
        self.notes_box.grid(row=1, column=1, sticky="nswe", padx=(0, 10), pady=(5, 10), columnspan=2)

    def get(self):
        name = self.name_box.get()
        if name == "":
            messagebox.showwarning("Missing Name", "Warning: There is no patient name inputted.")
        return {
            "name": name,
            "notes": self.notes_box.get("0.0", ctk.END),
            "prepaid": "YES" if self.prepaid.get() == 1 else "NO"
        }

class MyCheckboxFrame(ctk.CTkFrame):
    def __init__(self, master, values):
        super().__init__(master)
        self.labs = values
        self.checkboxes = []

        for i, lab in enumerate(self.labs):
            checkbox = ctk.CTkCheckBox(self, text=f"{lab["test_code"]}: {lab["description"]}")
            checkbox.grid(row=i, column=0, padx=10, pady=(10, 0), sticky="w")
            self.checkboxes.append(checkbox)

    def get(self):
        checked_checkboxes = []
        for checkbox in self.checkboxes:
            if checkbox.get() == 1:
                checkbox_text = checkbox.cget("text")
                test_code = checkbox_text.split(':')[0]
                checked_checkboxes.append(test_code)
        return checked_checkboxes

def main():
	app = App()
	app.mainloop()
