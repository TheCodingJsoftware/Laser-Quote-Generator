import configparser
import json
import os
import shutil
import socket
import sys
import threading
import time
import tkinter
import tkinter as tk
from functools import partial
from tkinter import filedialog, messagebox, ttk
from tkinter.constants import *

import sv_ttk
from PIL import Image, ImageTk

program_directory = os.path.dirname(os.path.realpath(sys.argv[0]))


global_variables = configparser.ConfigParser()
global_variables.read(f"{program_directory}/global_variables.cfg")
materials = global_variables["GLOBAL VARIABLES"]["materials"].split(",")
input_dialogs = {}


class VerticalScrolledFrame(ttk.Frame):
    """A pure Tkinter scrollable frame that actually works!
    * Use the 'interior' attribute to place widgets inside the scrollable frame.
    * Construct and pack/place/grid normally.
    * This frame only allows vertical scrolling.
    """

    def __init__(self, parent, *args, **kw):
        ttk.Frame.__init__(self, parent, *args, **kw)

        # Create a canvas object and a vertical scrollbar for scrolling it.
        vscrollbar = ttk.Scrollbar(self, orient=VERTICAL)
        vscrollbar.pack(fill=Y, side=RIGHT, expand=FALSE)
        canvas = tk.Canvas(
            self, bd=0, highlightthickness=0, yscrollcommand=vscrollbar.set
        )
        canvas.pack(side=LEFT, fill=BOTH, expand=TRUE)
        vscrollbar.config(command=canvas.yview)

        # Reset the view
        canvas.xview_moveto(0)
        canvas.yview_moveto(0)

        # Create a frame inside the canvas which will be scrolled with it.
        self.interior = interior = ttk.Frame(canvas)
        interior_id = canvas.create_window(0, 0, window=interior, anchor=NW)

        # Track changes to the canvas and frame width and sync them,
        # also updating the scrollbar.
        def _configure_interior(event):
            # Update the scrollbars to match the size of the inner frame.
            size = (1300, interior.winfo_reqheight())
            canvas.config(scrollregion="0 0 %s %s" % size)
            if 1300 != canvas.winfo_width():
                # Update the canvas's width to fit the inner frame.
                canvas.config(width=1300, height=600)

        interior.bind("<Configure>", _configure_interior)

        def _configure_canvas(event):
            if 1300 != canvas.winfo_width():
                # Update the inner frame's width to fill the canvas.
                canvas.itemconfigure(interior_id, width=canvas.winfo_width())

        canvas.bind("<Configure>", _configure_canvas)


class WrappingLabel(ttk.Label):
    """a type of Label that automatically adjusts the wrap to the size"""

    def __init__(self, master=None, **kwargs):
        tk.Label.__init__(self, master, **kwargs)
        self.bind("<Configure>", lambda e: self.config(wraplength=self.winfo_width()))


class ToggleButton(ttk.Button):
    """> A ttk.Button that toggles between two states when clicked"""

    ON_config = {
        "text": "RECUT",
    }
    OFF_config = {
        "text": "NOT RECUT",
    }

    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.toggled = False
        self.config = self.OFF_config
        self.config_button()
        self.text = kwargs["text"].split(";")[0]
        self.json_file_path = kwargs["text"].split(";")[1]

        self.bind("<Button-1>", self.toggle)

    def toggle(self, *args):
        check_part_number_boolean(self.json_file_path, self.text)
        self.config = self.OFF_config if self.toggled else self.ON_config
        self.toggled = not self.toggled
        return self.config_button()

    def config_button(self):
        self["text"] = self.config["text"]
        return "break"

    def __str__(self):
        return f"{self['text']}, {self['bg']}, {self['relief']}"


def check_part_number_boolean(json_file_path, part_name) -> None:
    """
    It opens the json file, loads the data, changes the value of the key "checked" to the opposite of
    what it was, and then writes the data back to the json file

    Args:
      json_file_path: The path to the json file
      part_name: The name of the part you want to check/uncheck.
    """
    with open(json_file_path) as f:
        data = json.load(f)

    data[part_name]["recut"] = not data[part_name]["recut"]

    with open(json_file_path, "w") as f:
        json.dump(data, f, sort_keys=True, ensure_ascii=False, indent=4)


def upload_file(command: str, json_file_path: str) -> None:
    """
    It sends a command, a file path, and a file size to the server, then sends the file in chunks of
    8192 bytes

    Args:
      command (str): str = "upload"
      json_file_path (str): The path to the file you want to send.
    """
    SERVER_IP: str = "10.0.0.93"
    SERVER_PORT: int = 80
    BUFFER_SIZE: int = 8192
    SEPARATOR = "<SEPARATOR>"
    server = (SERVER_IP, SERVER_PORT)
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(15)
    s.connect(server)
    filesize = os.path.getsize(json_file_path)

    s.send(f"{command}{SEPARATOR}{json_file_path}{SEPARATOR}{filesize}".encode())
    time.sleep(1)  # ! IMPORTANT
    with open(json_file_path, "rb") as f:
        while True:
            if bytes_read := f.read(BUFFER_SIZE):
                s.sendall(bytes_read)
            else:
                # file transmitting is done
                time.sleep(1)  # ! IMPORTANT
                break
    s.sendall("FINSIHED!".encode("utf-8"))
    response = s.recv(1024).decode("utf-8")
    s.shutdown(2)
    s.close()
    if response == "Batch sent successfully":
        messagebox.showinfo(
            "Success",
            "Batch was sent successfully.",
        )
    else:
        messagebox.showerror(
            "Oh no",
            f"{response}",
        )


def go_button_pressed(root, json_file_path, material_type) -> None:
    """
    It starts a new thread that calls the upload_file function with the arguments
    "laser_parts_list_upload" and the json_file_path

    Args:
      json_file_path: The path to the JSON file that you want to upload.
    """
    with open(f"{program_directory}/action", "w") as f:
        f.write("go")
    with open(json_file_path) as f:
        data = json.load(f)
    for item in list(data.keys()):
        data[item]["material"] = material_type.get()
    with open(json_file_path, "w") as f:
        json.dump(data, f, sort_keys=True, indent=4)
    threading.Thread(
        target=upload_file, args=["laser_parts_list_upload", json_file_path]
    ).start()
    root.destroy()
    messagebox.showinfo(
        "Sending",
        "I am sending file hold on",
    )


def make_quote_button_pressed(root, json_file_path, material_type) -> None:
    """
    This function updates a JSON file with a selected material type and writes "quote" to a separate
    file before destroying the root window.

    Args:
      root: The root parameter is typically a reference to the main window or frame of a GUI
    application. It is used to access and modify the widgets and properties of the application.
      json_file_path: The file path to a JSON file that contains data to be modified.
      material_type: It is a variable that contains the selected material type. It is likely a tkinter
    StringVar() object that is used to store the value of a dropdown menu or radio button selection. The
    value of this variable is used to update the "material" field in a JSON file.
    """
    with open(json_file_path) as f:
        data = json.load(f)
    for item in list(data.keys()):
        data[item]["material"] = material_type.get()
    with open(json_file_path, "w") as f:
        json.dump(data, f, sort_keys=True, indent=4)
    with open(f"{program_directory}/action", "w") as f:
        f.write("quote")
    root.destroy()


def get_total_sheet_count(json_file_path) -> int:
    """
    > It opens the JSON file, loads the data, and then sums the quantity_multiplier

    Args:
      json_file_path: The path to the JSON file that contains the data for the parts.

    Returns:
      The total number of sheets in the json file.
    """
    with open(json_file_path) as f:
        data = json.load(f)
    sheet_count: int = sum(
        data[part_name]["quantity_multiplier"]
        for part_name in list(data.keys())
        if part_name[0] == "_"
    )
    return sheet_count


def quantity_change(json_file_path, part_name: str) -> None:
    """
    This function updates the quantity of a specific part in a JSON file based on user input.

    Args:
      json_file_path: The file path of the JSON file that contains the data to be modified.
      part_name (str): The parameter `part_name` is a string that represents the name of a part in a
    JSON file. The function `quantity_change` updates the quantity of this part in the JSON file based
    on user input.
    """
    with open(json_file_path) as f:
        data = json.load(f)

    data[part_name]["quantity"] = int(input_dialogs[part_name].get())

    with open(json_file_path, "w") as f:
        json.dump(data, f, sort_keys=True, ensure_ascii=False, indent=4)


def load_gui(json_file_path: str, selected_material_type: str) -> None:
    """
    It loads a JSON file, then creates a GUI with a scrollable frame, and then populates the frame with
    the data from the JSON file.

    Args:
      json_file_path (str): str
    """
    root = tkinter.Tk()
    root.title("Laser Quote Generator - Add parts to Inventory")
    root.lift()
    root.attributes("-topmost", True)
    width, height = 950, 810
    root.geometry(f"{width}x{height}")
    root.minsize(width, height)
    root.maxsize(width, height)
    # This is where the magic happens
    # sv_ttk.set_theme("dark")

    with open(json_file_path) as f:
        data = json.load(f)
    part_to_get_sheet_dim = ""
    for part_name in list(data.keys()):
        if part_name[0] == "_":
            continue
        part_to_get_sheet_dim = part_name
    panel = ttk.Label(
        root,
        text=f"Total Sheet Count: {get_total_sheet_count(json_file_path)} - Sheet Size: {data[list(data.keys())[0]]['sheet_dim']} - Thickness: {data[list(data.keys())[0]]['gauge']}",
    )
    panel.pack()
    panel = ttk.Label(
        root,
        text="\nMaterial:",
    )
    panel.pack()
    material_type = tk.StringVar(root)
    material_type.set(selected_material_type)  # default value
    dropdown_material = ttk.OptionMenu(
        root, material_type, *[selected_material_type] + materials
    )
    dropdown_material.pack()
    panel = ttk.Label(
        root,
        text="",
    )
    panel.pack()

    frame = VerticalScrolledFrame(root)
    frame.pack()

    for col_i, header in enumerate(["Item", "Part Name", "Quantity", "Recut or not"]):
        panel = ttk.Label(frame.interior, text=header)
        panel.grid(row=0, column=col_i)
        panel.grid_rowconfigure(0, weight=1)
        panel.grid_columnconfigure(0, weight=1)
    for row_i, part_name in enumerate(list(data.keys()), start=1):
        if part_name[0] == "_":
            continue
        img = Image.open(
            f"{program_directory}/images/{data[part_name]['image_index']}.jpeg"
        )
        img = img.resize((64, 64), Image.ANTIALIAS)
        img = ImageTk.PhotoImage(img)
        panel = ttk.Label(frame.interior, image=img)
        panel.image = img
        panel.grid_rowconfigure(0, weight=1)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid(row=row_i, column=0, padx=50, pady=5)
        panel = WrappingLabel(
            frame.interior, text=part_name, wraplength=300, justify="center"
        )
        panel.grid_rowconfigure(0, weight=1)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid(row=row_i, column=1, padx=50, pady=5)

        var = tk.DoubleVar(root)
        panel = tk.Spinbox(
            frame.interior,
            from_=0,
            to=99999999,
            textvariable=var,
            command=partial(quantity_change, json_file_path, part_name),
        )
        var.set(str(data[part_name]["quantity"]))
        input_dialogs[part_name] = panel
        panel.grid_rowconfigure(0, weight=1)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid(row=row_i, column=2, padx=50, pady=5)
        # panel.pack()
        panel = ToggleButton(frame.interior, text=f"{part_name};{json_file_path}")
        panel.grid_rowconfigure(0, weight=1)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid(row=row_i, column=3, padx=50, pady=5)

    # NOTE Make work order with col hidden and send to inventory
    recut_button = ttk.Button(
        root,
        text="Send to Inventory &\nGenerate workorder!!",
        command=partial(go_button_pressed, root, json_file_path, material_type),
    )
    recut_button.place(rely=1.0, relx=1.0, x=-10, y=-10, anchor=SE, width=150, height=80)

    # NOTE Quote NOT GO TO INVENTORY BUT JUST MAKE EXCEL FILE
    quote_button = ttk.Button(
        root,
        text="Generate Quote!",
        command=partial(make_quote_button_pressed, root, json_file_path, material_type),
    )
    quote_button.place(rely=1.0, relx=1.0, x=-170, y=-10, anchor=SE, width=150, height=80)
    root.mainloop()


if __name__ == "__main__":
    load_gui(
        r"F:\Code\Python-Projects\Laser-Quote-Generator\excel files\2023-04-26-11-05-37.json",
        "304 SS",
    )
