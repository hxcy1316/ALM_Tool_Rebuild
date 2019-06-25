import tkinter as tk
import main


class GUI(tk.Frame):
    def __init__(self, master=tk.Tk()):
        master.title("Thin Client QA ALM Tool")
        tk.Frame.__init__(self, master)
        self.grid()

        self.url_Label = tk.Label(master, text="ALM Address:")
        self.url_Label.grid(row=0, column=0, padx=10, pady=10)
        self.url_Entry = tk.StringVar()
        self.url_Entry = tk.Entry(textvariable=self.url_Entry, width=50)
        self.url_Entry.grid(row=0, column=1, padx=10, pady=10)

        self.username_Label = tk.Label(master, text="User Name:")
        self.username_Label.grid(row=1, column=0, padx=10, pady=10)
        self.username_Entry = tk.StringVar()
        self.username_Entry = tk.Entry(textvariable=self.username_Entry, width=50)
        self.username_Entry.grid(row=1, column=1, padx=10, pady=10)

        self.password_Label = tk.Label(master, text="Password:")
        self.password_Label.grid(row=2, column=0, padx=10, pady=10)
        self.password_Entry = tk.StringVar()
        self.password_Entry = tk.Entry(textvariable=self.password_Entry, width=50)
        self.password_Entry.grid(row=2, column=1, padx=10, pady=10)

        self.domain_Label = tk.Label(master, text="Domain:")
        self.domain_Label.grid(row=3, column=0, padx=10, pady=10)
        self.domain_Entry = tk.StringVar()
        self.domain_Entry = tk.Entry(textvariable=self.domain_Entry, width=50)
        self.domain_Entry.grid(row=3, column=1, padx=10, pady=10)

        self.project_Label = tk.Label(master, text="Project:")
        self.project_Label.grid(row=4, column=0, padx=10, pady=10)
        self.project_Entry = tk.StringVar()
        self.project_Entry = tk.Entry(textvariable=self.project_Entry, width=50)
        self.project_Entry.grid(row=4, column=1, padx=10, pady=10)

        self.instance_path_Label = tk.Label(master, text="Instance Path:")
        self.instance_path_Label.grid(row=5, column=0, padx=10, pady=10)
        self.instance_path_Entry = tk.StringVar()
        self.instance_path_Entry = tk.Entry(textvariable=self.url_Entry, width=50)
        self.instance_path_Entry.grid(row=5, column=1, padx=10, pady=10)

        def buttonClick():
            url = self.url_Entry.get()
            username = self.username_Entry.get()
            password = self.password_Entry.get()
            domain = self.domain_Entry.get()
            project = self.project_Entry.get()
            path = self.instance_path_Entry.get()
            main.Get_Execution_Summary(url, username, password, domain, project, path)

        self.submitButton = tk.Button(master,
                                      text="Generate",
                                      command=buttonClick)
        self.submitButton.grid(row=6, column=0, padx=10, pady=10)


if __name__ == "__main__":
    guiFrame = GUI()
    guiFrame.mainloop()
