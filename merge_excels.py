import os
import glob
import pandas as pd
from tkinter import *
from tkinter.filedialog import askdirectory, asksaveasfilename


GUI = Tk()
GUI.title("Merge Excels")
GUI.geometry("500x220")
EA_welcome = Label(GUI, text="Merge Excels", font=("Courier", 14, "bold"))
EA_welcome.pack(pady=10, padx=60)

class filesclass:
	def __init__(self):
		self.filenames = []

FC = filesclass()
cwd = os.getcwd()


def newlocation():
	# Autoselect one dir above
	location = askdirectory(title="Select a folder to merge", initialdir=os.path.dirname(cwd))
	selectedfolder_label.configure(text=location)
	# Select only .xlsx files
	pathlocation = os.path.normpath(location) + "\\*.xlsx"
	FC.filenames = glob.glob(pathlocation)
	selectedfolder_label.configure(text=location + "\n" + str(len(FC.filenames)) + " excel files have been found.")
	fileselection_label.configure(text="")


select_button = Button(GUI, text="Select a folder to merge", font=("Courier", 12), command=newlocation)
select_button.pack(pady=8)

selectedfolder_label = Label(GUI, text="Please select a folder where the excel files are located.", font=("Courier", 10))
selectedfolder_label.pack()


def generate_excel():
	if(len(FC.filenames) > 0):
		fileselection_label.configure(text="Please wait...", fg="black")
		df = pd.DataFrame()
		for file in FC.filenames:
			# Copy the excels together
			df1 = pd.read_excel(file)
			df = pd.concat([df, df1], ignore_index=True)
			df.fillna(value="N/A")
		try:
			output_file = asksaveasfilename(filetypes=(("Excel document", "*.xlsx"), ("All files", "*.*")), defaultextension=".xlsx", initialdir=cwd, initialfile="Merged", title="Save merged Excel")
			df.to_excel(output_file, index=False)
			fileselection_label.configure(text=output_file + "\nSuccessfully generated.", fg="black", font=("Courier", 10))
		except ValueError:
			fileselection_label.configure(text="Please select a file.", fg="red")
	else:
		fileselection_label.configure(text="Please select a folder first.", fg="red")


generate_button = Button(GUI, text="Generate merged excel", font=("Courier", 12), command=generate_excel)
generate_button.pack(pady=8)

fileselection_label = Label(GUI, font=("Courier", 12))
fileselection_label.pack()

GUI.mainloop()
