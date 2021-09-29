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
        self.multiplier_instruction = tk.Label(self.labelframe, wraplength=600,
                                               justify='center')
        self.multiplier_instruction["text"] = ("Please enter the multiplier to be used.")
        self.multiplier_instruction.pack(side="top")
        self.multiplier = tk.Entry(self.labelframe, justify='center')
        self.multiplier.insert(0, 100)
        self.multiplier.pack(side="top")
        self.sorter_instruction = tk.Label(self.labelframe, wraplength=600,
                                           justify='center')
        self.sorter_instruction["text"] = ("Please choose which column to sort by.")
        self.sorter_instruction.pack(side="top")
        self.sorter = tk.Entry(self.labelframe, justify='center')
        self.sorter.insert(0, 'A')
        self.sorter.pack(side="top")
        self.feedback = tk.Label(self.labelframe, wraplength=600,
                                 justify='center')
        self.feedback["text"] = ("Please choose a directory, then wait for processing to be done.")
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
        multiplier = self.multiplier.get()
        try:
            multiplier = int(multiplier)
        except Exception as e:
            feedback = "Please enter a number for the multiplier."
            self.feedback["text"] = feedback
        sorter = self.sorter.get()
        self.dir = filedialog.askdirectory(**options)
        #self.feedback["text"] = 'Doing the summary now, please wait...'
        feedback = do_a_summary(self.dir, multiplier=multiplier, sorter=sorter)
        self.feedback["text"] = feedback

if __name__ == '__main__':
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()
