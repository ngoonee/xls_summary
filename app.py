import tkinter as tk
from tkinter import filedialog
from xls_summary.summarize import do_a_summary

class Application(tk.Frame):
    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.pack()
        self.createWidgets()
        self.dir = None

    def createWidgets(self):
        self.labelframe = tk.LabelFrame(self, text="Welcome to Excel Summary")
        self.labelframe.pack(side="top")
        self.feedback = tk.Label(self.labelframe, wraplength=600,
                                 justify='center')
        self.feedback["text"] = ("Please choose a directory.")
        self.feedback.pack(side="top")
        self.choose_dir = tk.Button(self.labelframe,
                                    text="Choose Directory",
                                    command=self.choose_directory,
                                    ).pack(side="top")
        self.QUIT = tk.Button(self, text="QUIT", fg="red",
                                            command=root.destroy)
        self.QUIT.pack(side="bottom")

    def choose_directory(self):
        options = {}
        options['mustexist'] = False
        options['parent'] = self
        options['title'] = 'Choose the directory'
        self.dir = filedialog.askdirectory(**options)
        feedback = do_a_summary(self.dir)
        self.feedback["text"] = feedback

if __name__ == '__main__':
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()
