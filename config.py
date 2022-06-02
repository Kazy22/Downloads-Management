import os
import win32com.client
from tkinter import *
from tkinter import messagebox
import customtkinter
from pathlib import Path
from tkinter import filedialog
import json
import subprocess

customtkinter.set_appearance_mode("Dark")
customtkinter.set_default_color_theme("green")

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.APPDATA = os.getenv("APPDATA")
        self.DOWNLOAD_PATH = str(Path.home() / "Downloads")
        self.JSON = "config.json"

        with open(self.JSON, "r") as file:
            self.settings = json.load(file)

        self.title("File Downloads management")
        self.geometry("421x364")
        self.resizable(0, 0)
        self.grid_columnconfigure(1, weight = 1)
        self.grid_rowconfigure(0, weight = 1)

        # ======= Main Frame =======

        self.main_frame = Frame(self, bd = 0, bg = "#1f1f1f")
        self.main_frame.pack(pady = (20,0))

        # ======= Paths Set Frame =======

        self.paths_frame = Frame(self.main_frame, bd = 0, bg = "#1f1f1f")
        self.paths_frame.pack()

        # ======= Images Folder Frame =======

        self.image_frame = Frame(self.paths_frame, bd = 0, bg = "#1f1f1f")
        self.image_frame.grid(row = 1, pady = (20, 0))

        self.image_label = customtkinter.CTkLabel(self.image_frame, text = "Images Path:")
        self.image_label.grid(row = 0, column = 1, sticky = W, padx = (0, 0))

        self.image_var = IntVar()
        self.image_var.set(self.settings["dir_image_check"])
        self.image_checkbox = customtkinter.CTkCheckBox(self.image_frame, text = "", variable = self.image_var, command = self.image_check)
        self.image_checkbox.grid(row = 1, column = 0)
            
        self.image_entry = customtkinter.CTkEntry(self.image_frame, width = 220)
        self.image_entry.grid(row = 1, column = 1, padx = (10, 0))
        self.image_entry.insert(0, self.settings["dest_dir_image"])
        self.image_entry.bind("<Key>", lambda e: "break")

        self.image_browse = customtkinter.CTkButton(self.image_frame, text = "Browse...", width = 10, bd = 0, command = self.image_dialog)
        self.image_browse.grid(row = 1, column = 2, padx = 10)

        if not self.settings["dir_image_check"]:
            self.image_entry.delete(0, END)
            self.image_entry.config(state = DISABLED)
            self.image_browse.config(state = DISABLED)

        # ======= Video Folder Frame =======

        self.video_frame = Frame(self.paths_frame, bd = 0, bg = "#1f1f1f")
        self.video_frame.grid(row = 2)

        self.video_label = customtkinter.CTkLabel(self.video_frame, text = "Videos Path:")
        self.video_label.grid(row = 0, column = 1, sticky = W, padx = (0, 0))

        self.video_var = IntVar()
        self.video_var.set(self.settings["dir_video_check"])
        self.video_checkbox = customtkinter.CTkCheckBox(self.video_frame, text = "", variable = self.video_var, command = self.video_check)
        self.video_checkbox.grid(row = 1, column = 0)

        self.video_entry = customtkinter.CTkEntry(self.video_frame, width = 220)
        self.video_entry.grid(row = 1, column = 1, padx = (10, 0))
        self.video_entry.insert(0, self.settings["dest_dir_video"])
        self.video_entry.bind("<Key>", lambda e: "break")

        self.video_browse = customtkinter.CTkButton(self.video_frame, text = "Browse...", width = 10, bd = 0, command = self.video_dialog)
        self.video_browse.grid(row = 1, column = 2, padx = 10)

        # ======= Musics Folder Frame =======

        self.music_frame = Frame(self.paths_frame, bd = 0, bg = "#1f1f1f")
        self.music_frame.grid(row = 3)

        self.music_label = customtkinter.CTkLabel(self.music_frame, text = "Musics Path:")
        self.music_label.grid(row = 0, column = 1, sticky = W, padx = (0, 0))

        self.music_var = IntVar()
        self.music_var.set(self.settings["dir_music_check"])
        self.music_checkbox = customtkinter.CTkCheckBox(self.music_frame, text = "", variable = self.music_var, command = self.music_chek)
        self.music_checkbox.grid(row = 1, column = 0)

        self.music_entry = customtkinter.CTkEntry(self.music_frame, width = 220)
        self.music_entry.grid(row = 1, column = 1, padx = (10, 0))
        self.music_entry.insert(0, self.settings["dest_dir_music"])
        self.music_entry.bind("<Key>", lambda e: "break")

        self.music_browse = customtkinter.CTkButton(self.music_frame, text = "Browse...", width = 10, bd = 0, command = self.music_dialog)
        self.music_browse.grid(row = 1, column = 2, padx = 10)

        if not self.settings["dir_music_check"]:
            self.music_entry.delete(0, END)
            self.music_entry.config(state = DISABLED)
            self.music_browse.config(state = DISABLED)

        # ======= Documents Folder Frame =======

        self.document_frame = Frame(self.paths_frame, bd = 0, bg = "#1f1f1f")
        self.document_frame.grid(row = 4)

        self.document_label = customtkinter.CTkLabel(self.document_frame, text = "Documents Path:")
        self.document_label.grid(row = 0, column = 1, sticky = W, padx = 10)

        self.document_var = IntVar()
        self.document_var.set(self.settings["dir_documents_check"])
        self.document_checkbox = customtkinter.CTkCheckBox(self.document_frame, text = "", variable = self.document_var, command = self.document_check)
        self.document_checkbox.grid(row = 1, column = 0)

        self.document_entry = customtkinter.CTkEntry(self.document_frame, width = 220)
        self.document_entry.grid(row = 1, column = 1, padx = (10, 0))
        self.document_entry.insert(0, self.settings["dest_dir_documents"])
        self.document_entry.bind("<Key>", lambda e: "break")

        self.document_browse = customtkinter.CTkButton(self.document_frame, text = "Browse...", width = 10, bd = 0, command = self.document_dialog)
        self.document_browse.grid(row = 1, column = 2, padx = 10)

        if not self.settings["dir_documents_check"]:
            self.document_entry.delete(0, END)
            self.document_entry.config(state = DISABLED)
            self.document_browse.config(state = DISABLED)

        # ======= Add to Startup =======

        self.startup_var = IntVar()
        self.startup_var.set(self.settings["startup_check"])
        self.add_startup = customtkinter.CTkCheckBox(self.paths_frame, text = "Add to Startup", corner_radius = 50, variable = self.startup_var)
        self.add_startup.grid(row = 5, pady = (30,0), sticky = W)
        self.add_startup.configure(border_color = "#8AE0C3")

        # ======= ok/cancel buttons =======

        self.second_frame = Frame(self.main_frame, bd = 0, bg = "#1f1f1f")
        self.second_frame.pack(pady = (10, 0), side = RIGHT)

        self.ok_button = customtkinter.CTkButton(self.second_frame, text = "Ok", width = 60, command = self.save)
        self.ok_button.grid(row = 0, column = 0, padx = 10)
        self.ok_button.configure(fg_color = "#1a8aba")

        self.cancel_button = customtkinter.CTkButton(self.second_frame, text = "Cancel", width = 60, command = self.destroy)
        self.cancel_button.grid(row = 0, column = 1, padx = (0, 10))
        self.cancel_button.configure(fg_color = "#1a8aba")


    def add_to_startup(self, file_path=f"{os.path.dirname(os.path.realpath(__file__))}"):
        path = r"{}\Microsoft\Windows\Start Menu\Programs\Startup".format(self.APPDATA)
        path = os.path.join(path, "DownloadManagement.lnk")
        target = r"{}\fileAutomation.pyw".format(file_path)
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = target
        shortcut.WorkingDirectory  = file_path
        shortcut.save()

    def image_dialog(self):
        path = filedialog.askdirectory(initialdir = f"{Path.home()}")
        if path:
            self.image_entry.delete(0, END)
            self.image_entry.insert(0, path)

    def video_dialog(self):
        path = filedialog.askdirectory(initialdir = f"{Path.home()}")
        if path:
            self.video_entry.delete(0, END)
            self.video_entry.insert(0, path)

    def music_dialog(self):
        path = filedialog.askdirectory(initialdir = f"{Path.home()}")
        if path:
            self.music_entry.delete(0, END)
            self.music_entry.insert(0, path)

    def document_dialog(self):
        path = filedialog.askdirectory(initialdir = f"{Path.home()}")
        if path:
            self.document_entry.delete(0, END)
            self.document_entry.insert(0, path)

    def image_check(self):
        if not self.settings["dir_image_check"]:
            self.image_browse.config(state = NORMAL)
            self.image_entry.config(state = NORMAL)
            self.image_entry.delete(0, END)
            self.image_entry.insert(0, self.settings["dest_dir_image"])

            self.settings["dir_image_check"] = 1

            with open(self.JSON, 'w') as file:
                json.dump(self.settings, file, indent = 2)

        else:
            self.image_entry.delete(0, END)
            self.image_browse.config(state = DISABLED)
            self.image_entry.config(state = DISABLED)

            self.settings["dir_image_check"] = 0

            with open(self.JSON, 'w') as file:
                json.dump(self.settings, file, indent = 2)

    def video_check(self):
        if not self.settings["dir_video_check"]:
            self.video_browse.config(state = NORMAL)
            self.video_entry.config(state = NORMAL)
            self.video_entry.delete(0, END)
            self.video_entry.insert(0, self.settings["dest_dir_video"])

            self.settings["dir_video_check"] = 1

            with open(self.JSON, 'w') as file:
                json.dump(self.settings, file, indent = 2)

        else:
            self.video_entry.delete(0, END)
            self.video_browse.config(state = DISABLED)
            self.video_entry.config(state = DISABLED)

            self.settings["dir_video_check"] = 0

            with open(self.JSON, 'w') as file:
                json.dump(self.settings, file, indent = 2)

    def music_chek(self):
        if not self.settings["dir_music_check"]:
            self.music_browse.config(state = NORMAL)
            self.music_entry.config(state = NORMAL)
            self.music_entry.delete(0, END)
            self.music_entry.insert(0, self.settings["dest_dir_music"])

            self.settings["dir_music_check"] = 1

            with open(self.JSON, 'w') as file:
                json.dump(self.settings, file, indent = 2)

        else:
            self.music_entry.delete(0, END)
            self.music_browse.config(state = DISABLED)
            self.music_entry.config(state = DISABLED)

            self.settings["dir_music_check"] = 0

            with open(self.JSON, 'w') as file:
                json.dump(self.settings, file, indent = 2)

    def document_check(self):
        if not self.settings["dir_documents_check"]:
            self.document_browse.config(state = NORMAL)
            self.document_entry.config(state = NORMAL)
            self.document_entry.delete(0, END)
            self.document_entry.insert(0, self.settings["dest_dir_documents"])

            self.settings["dir_documents_check"] = 1

            with open(self.JSON, 'w') as file:
                json.dump(self.settings, file, indent = 2)

        else:
            self.document_entry.delete(0, END)
            self.document_browse.config(state = DISABLED)
            self.document_entry.config(state = DISABLED)

            self.settings["dir_documents_check"] = 0

            with open(self.JSON, 'w') as file:
                json.dump(self.settings, file, indent = 2)

    def getTasks(self, name):
        r = os.popen("tasklist /v").read().strip().split("\n")
        for i in range(len(r)):
            s = r[i]
            if name in r[i]:
                print (f"{name} in r[i]")
                return r[i]
        return []

    def save(self):
        with open(self.JSON, "w") as file:
            self.settings["source_dir"] = str(os.path.join(f"{Path.home()}\\Downloads"))
            self.settings["dest_dir_image"] = self.image_entry.get() if self.settings["dir_image_check"] == 1 else self.settings["dest_dir_image"]
            self.settings["dest_dir_video"] = self.video_entry.get() if self.settings["dir_video_check"] == 1 else self.settings["dest_dir_video"]
            self.settings["dest_dir_music"] = self.music_entry.get() if self.settings["dir_music_check"] == 1 else self.settings["dest_dir_music"]
            self.settings["dest_dir_documents"] = self.document_entry.get() if self.settings["dir_documents_check"] == 1 else self.settings["dest_dir_documents"]
            self.settings["startup_check"] = self.startup_var.get()
            json.dump(self.settings, file, indent = 2)

        with open(self.JSON, "r") as file:
            self.settings = json.load(file)

        if self.settings["dir_image_check"] and not self.image_entry.get() or self.settings["dir_video_check"] and not self.video_entry.get() or self.settings["dir_music_check"] and not self.video_entry.get() or self.settings["dir_documents_check"] and not self.document_entry.get():
            messagebox.showerror("Error", "Select a Path or uncheck a option to ignore these type of files")

        else:
            self.destroy()

            if self.settings["startup_check"]:
                self.add_to_startup()
            
            else:
                try:
                    os.remove(r"{}\Microsoft\Windows\Start Menu\Programs\Startup\DownloadManagement.lnk".format(self.APPDATA))
                except FileNotFoundError:
                    pass

            task = self.getTasks("pythonw.exe")
            if task:
                os.system("taskkill /f /im pythonw.exe")
                os.system("cls")
            
            subprocess.call("pyw fileAutomation.pyw", shell = True)

    def start(self):
        self.mainloop()

if __name__ == "__main__":
    app = App()
    app.start()
