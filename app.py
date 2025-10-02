import tkinter as tk

from tkinter.scrolledtext import ScrolledText
from tkinter import ttk
import csv
from time import strftime

CSV_FILE = "users.csv"   # your CSV file

FIELDS = [
	"register_date", "district", "address", "exam_date",
	"time", "group", "arc", "birthday", "name", "phone", "status"
]

class App:
	def __init__(self, root):
		self.root = root
		self.root.title("CSV Viewer")
		self.root.geometry("700x450")
		self.root.resizable(False, False)

		self.log_box = tk.Text(root, height=5, wrap=tk.WORD)
		self.log_box.config(spacing1=5)
		self.log_box.config(spacing2=5)
		self.log_box.config(spacing3=5)
		self.log_box.pack(side=tk.BOTTOM, fill=tk.X,pady=20,padx=20)

		self.time_label = tk.Label(self.root, font=("Arial", 20), fg='red')
		self.time_label.pack(pady=20)

		run_button = tk.Button(root, text="CHƠI!", font=("Arial", 10,'bold'), fg='blue',padx=20)
		run_button.pack(pady=5)

		# split window into left and right
		self.left_frame = tk.Frame(root)
		self.left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

		self.right_frame = tk.Frame(root)
		self.right_frame.pack(side=tk.LEFT, fill=tk.Y,   padx=10, pady=10)

		# listbox on the left
		self.listbox = tk.Listbox(self.left_frame, width=30, height=10)
		self.listbox.pack()



		# details on the right
		self.detail_vars = {}
		for i, field in enumerate(FIELDS):
			tk.Label(self.right_frame, text=field + ":").grid(row=i//2, column=0+2*(i%2),)
			var = tk.StringVar()
			entry = tk.Entry(self.right_frame, textvariable=var, width=30, state="readonly")
			entry.grid(row=i//2, column=1+2*(i%2), pady=5)
			if (i==len(FIELDS)-1):
				entry.config(fg='red')
			self.detail_vars[field] = var


		# load data
		self.records = self.get_user_from_csv()

		# fill listbox (e.g., show name + phone)
		for idx, record in enumerate(self.records):
			display_text = f"{idx+1}. {record['name']} ({record['arc']})"
			self.listbox.insert(tk.END, display_text)

		# bind selection
		self.listbox.bind("<<ListboxSelect>>", self.show_details)

	def get_user_from_csv(self,file_name="users.csv"):
		keys = [
			"register_date", "district", "address", "exam_date",
			"time", "group", "arc", "birthday", "name", "phone", "email"
		]
		result = []
		log = ""
		count = 0
		with open(file_name, "r", encoding="utf-8") as f:
			reader = csv.reader(f)
			for row in reader:
				entry = dict(zip(keys, row))
				entry['status']= "Not Register Yet"
				result.append(entry)
				count += 1

		return result

	def show_details(self, event):
		selection = self.listbox.curselection()
		if not selection:
			return
		index = selection[0]
		record = self.records[index]
		for field in FIELDS:
			self.detail_vars[field].set(record.get(field, ""))

	def update_gui(self):
		# update clock
		self.time_label.config(text="Time: " + strftime("%H:%M:%S"))

		root.after(500, self.update_gui)  # repeat
if __name__ == "__main__":
	root = tk.Tk()
	app = App(root)
	app.update_gui()
	root.mainloop()

