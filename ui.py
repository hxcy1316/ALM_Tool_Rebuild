import tkinter as tk
import main
import pickle

# Iverson edit

FILENAME = "previous.pickle"


class GUI(tk.Frame):
    def __init__(self, master=tk.Tk()):
        master.title("Thin Client QA ALM Tool")
        tk.Frame.__init__(self, master)
        self.grid()
        self.previous_data = {}
        self.restore_state()
        self.url_Label = tk.Label(master, text="ALM Address:")
        self.url_Label.grid(row=0, column=0, padx=10, pady=10)
        self.url_Entry = tk.StringVar(value=self.__get_previous_value("url"))
        self.url_Entry = tk.Entry(textvariable=self.url_Entry, width=50)
        self.url_Entry.grid(row=0, column=1, padx=10, pady=10)

        self.username_Label = tk.Label(master, text="User Name:")
        self.username_Label.grid(row=1, column=0, padx=10, pady=10)
        self.username_Entry = tk.StringVar(value=self.__get_previous_value("username"))
        self.username_Entry = tk.Entry(textvariable=self.username_Entry, width=50)
        self.username_Entry.grid(row=1, column=1, padx=10, pady=10)

        self.password_Label = tk.Label(master, text="Password:")
        self.password_Label.grid(row=2, column=0, padx=10, pady=10)
        self.password_Entry = tk.StringVar(value=self.__get_previous_value("password"))
        self.password_Entry = tk.Entry(textvariable=self.password_Entry, show="*", width=50)
        self.password_Entry.grid(row=2, column=1, padx=10, pady=10)

        self.domain_Label = tk.Label(master, text="Domain:")
        self.domain_Label.grid(row=3, column=0, padx=10, pady=10)
        self.domain_Entry = tk.StringVar(value=self.__get_previous_value("domain"))
        self.domain_Entry = tk.Entry(textvariable=self.domain_Entry, width=50)
        self.domain_Entry.grid(row=3, column=1, padx=10, pady=10)

        self.project_Label = tk.Label(master, text="Project:")
        self.project_Label.grid(row=4, column=0, padx=10, pady=10)
        self.project_Entry = tk.StringVar(value=self.__get_previous_value("project"))
        self.project_Entry = tk.Entry(textvariable=self.project_Entry, width=50)
        self.project_Entry.grid(row=4, column=1, padx=10, pady=10)

        self.instance_path_Label = tk.Label(master, text="Instance Path:")
        self.instance_path_Label.grid(row=5, column=0, padx=10, pady=10)
        self.instance_path_Entry = tk.StringVar(value=self.__get_previous_value("path"))
        self.instance_path_Entry = tk.Entry(textvariable=self.instance_path_Entry, width=50)
        self.instance_path_Entry.grid(row=5, column=1, padx=10, pady=10)
        self.submitButton = tk.Button(master,
                                      text="Generate",
                                      command=self.buttonClick)
        self.submitButton.grid(row=6, column=0, padx=10, pady=10)
        master.wm_protocol("WM_DELETE_WINDOW", self.save_state)

    def __get_previous_value(self, key_name):
        if key_name in self.previous_data:
            return self.previous_data[key_name]
        else:
            return None

    def buttonClick(self):
        url = self.url_Entry.get()
        self.previous_data["url"] = url
        username = self.username_Entry.get()
        self.previous_data["username"] = username
        password = self.password_Entry.get()
        self.previous_data["password"] = password
        domain = self.domain_Entry.get()
        self.previous_data["domain"] = domain
        project = self.project_Entry.get()
        self.previous_data["project"] = project
        path = self.instance_path_Entry.get()
        self.previous_data["path"] = path
        main.Get_Execution_Summary(url, username, password, domain, project, path)

    def save_state(self):
        try:
            data = self.previous_data
            with open(FILENAME, "wb") as f:
                pickle.dump(data, f)
        except Exception as e:
            print("error saving state:", str(e))
        self.master.destroy()

    def restore_state(self):
        try:
            with open(FILENAME, "rb") as f:
                data = pickle.load(f)
            self.previous_data = data
        except Exception as e:
            print("error loading saved state:", str(e))


if __name__ == "__main__":
    guiFrame = GUI()
    guiFrame.mainloop()
