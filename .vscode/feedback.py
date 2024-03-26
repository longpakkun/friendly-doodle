from tkinter import *
from tkinter import messagebox
import csv
from datetime import datetime

class feedback:
    def __init__(self, root, pw, ph, osname):
        self.root = root
        root.geometry("500x500+{}+{}".format(pw-250, ph-250))
        root.title("Feedback")

        self.frmName = Frame(root)
        self.frmName.pack(side="top", fill="x")
        self.lblName = Label(self.frmName, text="Name:", font=("Arial", 15, "bold"), justify="left")
        self.lblName.pack(side="left", padx=5, pady=5, ipadx=16)
        self.txtName = Entry(self.frmName, font=("Arial", 15))
        self.txtName.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        self.txtName.insert(0, osname)
        self.txtName.config(state="readonly")
        
        self.frmContent = Frame(root)
        self.frmContent.pack(side="top", fill="both", expand=True)
        self.frmLBLContent = Frame(self.frmContent)
        self.frmLBLContent.pack(side="left", fill="both", expand=True)
        self.lblContent = Label(self.frmLBLContent, text="Content:", font=("Arial", 15, "bold"))
        self.lblContent.pack(padx=5, pady=5, side="top", ipadx=5)
        self.txtContent = Text(self.frmContent, font=("Arial", 10))
        self.txtContent.pack(side="left", padx=5, pady=5, expand=True, fill="both")

        self.frmButton = Frame(root)
        self.frmButton.pack(side="top", fill="x")
        self.butSubmit = Button(self.frmButton, text="Submit", font=("Arial", 15, "bold"), command=self.butSubmit_Click)
        self.butSubmit.pack(side="right", padx=5, pady=5)

    def butSubmit_Click(self):
        today = datetime.today().strftime("%Y/%m/%d, %H:%M:%S")
        name = ""
        content = ""
        if self.txtName.get() == "" or self.txtContent.get("1.0", "end-1c") == "":
            messagebox.showerror(title="Error!", message="You do not input your name or content. \nPlease try agian!")
            return
        name = self.txtName.get()
        content = self.txtContent.get("1.0", "end-1c")

        txt = [today, name, content]

        with open(r"\\10.177.113.250\data$\Happy\python\Feedback\feedback.csv", "a", encoding="utf-8", newline='') as file:
            csv_writer = csv.writer(file)

            csv_writer.writerow(txt)
        
        self.txtName.delete(0, END)
        self.txtContent.delete("1.0", END)
        self.root.destroy()

if __name__ == "__main__":
    root = Tk()
    feedback(root, -1200, 300, "Happy_Luu")
    root.mainloop()