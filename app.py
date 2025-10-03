import tkinter as tk
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
from time import strftime
from multiprocessing import Process,Manager
import threading
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import math
import csv
import requests
import io

import os
import csv
from time import strftime

CSV_FILE = "users.csv"   # your CSV file

FIELDS = [
	"register_date", "district", "address", "exam_date",
	"time", "group", "arc", "birthday", "name", "phone", "status"
]


# --- Selenium automation function ---
def auto_regist(i,shared_user,n_user,shared_log,win_h,win_w):
	user = shared_user.values()[i]
	try:
		options = Options()
		options.add_argument("--ignore-certificate-errors")
		options.add_argument("--ignore-ssl-errors")
		options.add_argument("--disable-gpu")
		options.add_argument("--disable-notifications")
		options.add_argument("--disable-background-networking")
		options.add_argument("--disable-gcm")
		options.add_argument("--disable-extensions")
		options.add_argument("--disable-popup-blocking")
		options.add_argument("--disable-infobars")
		options.add_argument("--log-level=3")
		options.add_experimental_option('excludeSwitches', ['enable-logging'])


		screen_width = win_h - 20
		screen_height = win_w - 20

		cols = min(3, n_user)
		rows = math.ceil(n_user / cols)
		width = screen_width / cols
		height = screen_height / rows
		col = i % cols
		row = i // cols
		x_pos = col * width
		y_pos = row * height

		try:
			s = Service(ChromeDriverManager().install())
		except:
			try:
				s = Service("chromedriver/chromedriver.exe")
			except:
				print("CHROME DRIVER ERROR")
		driver = webdriver.Chrome(service=s, options=options)
		driver.set_window_size(width, height)
		driver.set_window_position(x_pos, y_pos)

		# 打開網頁
		driver.get("https://www.mvdis.gov.tw/m3-emv-trn/exm/locations#gsc.tab=0")
		try:

			# 等待「報考照類」下拉選單可點擊
			WebDriverWait(driver, 10).until(
				EC.element_to_be_clickable((By.ID, 'licenseTypeCode'))
			)

			# 定位「報考照類」並選擇
			select_license_type = Select(driver.find_element(By.ID, 'licenseTypeCode'))
			select_license_type.select_by_visible_text('普通重型機車')  # 選擇普通重型機車

			# 定位預計考試日期輸入框並填寫
			date_input = driver.find_element(By.ID, 'expectExamDateStr')
			date_input.clear()
			date_input.send_keys(user['register_date'])  # 填入預計考試日期

			# 定位第一個考試地點選單並選擇臺北區監理所（北宜花）
			select_region = Select(driver.find_element(By.ID, 'dmvNoLv1'))
			select_region.select_by_visible_text(user['district'])

			# 等待第二個下拉選單更新並包含特定選項
			WebDriverWait(driver, 10).until(
				EC.text_to_be_present_in_element(
					(By.ID, 'dmvNo'),
					user['address']
				)
			)

			# 定位第二個考試地點選單並選擇具體的監理站
			select_station = Select(driver.find_element(By.ID, 'dmvNo'))
			select_station.select_by_visible_text(user['address'])

			# 定位「查詢場次」按鈕並點擊
			search_button = driver.find_element(By.CSS_SELECTOR, "a[onclick='query();']")

			search_button.click()
			time.sleep(1)
			# 等待彈出視窗元素加載完成
			WebDriverWait(driver, 10).until(
				EC.presence_of_element_located((By.CLASS_NAME, "blockUI"))
			)

			# 使用JavaScript移除遮罩層
			driver.execute_script("""
			var overlay = document.querySelector('.blockUI.blockOverlay');
			if (overlay) {
				overlay.style.display = 'none';
			}
			var messageBox = document.querySelector('.blockUI.blockMsg.blockPage');
			if (messageBox) {
				messageBox.style.display = 'none';
			}
			""")
		except Exception as e:
			user['status'] = "FAIL: Cannot select LOCATION"
			shared_log.append(f"{user['arc']} {user['name']} FAIL: Cannot select LOCATION\n")

		exam_date = user["exam_date"]
		time_val = user["time"]
		group_val = user["group"]

		xpath = f"//a[@onclick=\"preAdd('{exam_date}', '{time_val}', '{group_val}')\"]"

		try:
			signup_button = driver.find_element(By.XPATH, xpath)
			signup_button.click()
			time.sleep(2)
			# 等待彈出視窗元素加載完成
			WebDriverWait(driver, 10).until(
				EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".blockUI"))
			)

			# 使用JavaScript強制移除所有遮罩層
			driver.execute_script("""
			Array.from(document.querySelectorAll('.blockUI')).forEach(el => el.remove());
			""")

			time.sleep(1)

			# 身分證字號
			id_input = driver.find_element(By.ID, 'idNo')
			id_input.send_keys(user['arc'])

			# 出生年月日
			birthday_input = driver.find_element(By.ID, 'birthdayStr')
			birthday_input.send_keys(user['birthday'])

			# 姓名
			name_input = driver.find_element(By.ID, 'name')
			name_input.send_keys(user['name'])

			# 聯絡電話/手機
			phone_input = driver.find_element(By.ID, 'contactTel')
			phone_input.send_keys(user['phone'])

			# Email
			email_input = driver.find_element(By.ID, 'email')
			email_input.send_keys('')
			driver.execute_script("window.scrollTo(0, document.body.scrollHeight / 2);")

			# 點擊最後的報名按鈕
			# final_signup_button = driver.find_element(By.CSS_SELECTOR, "a[onclick='add()']")
			# final_signup_button.click()
			final_signup_button = driver.find_element(By.CSS_SELECTOR, "a[onclick='add()']")

			# IMPORTANT: UNCOMMENT THIS LINE FOR AUTO SUBMIT
			# driver.execute_script("arguments[0].click();", final_signup_button)

			user['status'] = "SUCCESS"
			shared_log.append(f"{user['arc']} {user['name']} SUCCESS\n")
			# # 等待一些時間以便觀察
			time.sleep(15)
			# # 完成後 關閉瀏覽器
			driver.quit()
		except:
			driver.quit()
			user['status'] = "FAIL: Date,Time,Group NOT AVAILABLE"
			shared_log.append(f"{user['arc']} {user['name']} FAIL:  Date,Time,Group NOT AVAILABLE\n")


	except Exception as e:
		shared_log.append(f"{user['arc']} {user['name']} FAIL: {str(e)}\n")



class App:
	def __init__(self, root):
		self.root = root
		self.root.title("CSV Viewer")
		self.root.attributes("-topmost", True)

		# self.root.geometry("700x550")
		self.root.resizable(False, False)

		self.win_h, self.win_w = self.root.winfo_screenwidth(), self.root.winfo_screenheight(),

		self.manager = Manager()
		self.shared_dict = self.manager.dict()
		self.shared_log = self.manager.list()

		self.log_box = tk.Text(root, height=8, wrap=tk.WORD)
		self.log_box.config(spacing1=3)
		self.log_box.config(spacing2=3)
		self.log_box.config(spacing3=3)
		self.log_box.pack(side=tk.BOTTOM, fill=tk.X,pady=10,padx=10)



		self.time_label = tk.Label(self.root, font=("Arial", 20), fg='red')
		self.time_label.pack(pady=20)

		self.run_button = tk.Button(root, text="RUN!", font=("Arial", 10,'bold'), fg='blue',padx=20, command=lambda: self.run_in_thread())
		self.run_button.pack(pady=5)

		var_auto_submit = tk.BooleanVar()
		self.check_auto_submit = tk.Checkbutton(root,variable=var_auto_submit,text="Auto Submit",font=("Arial", 12), fg='red', justify=tk.LEFT)
		self.check_auto_submit.pack(pady=5)


		# split window into left and right
		self.left_frame = tk.Frame(root)
		self.left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=5)



		self.right_frame = tk.Frame(root)
		self.right_frame.pack(side=tk.RIGHT, fill=tk.Y,   padx=10, pady=5)

		# listbox on the left
		self.listbox = tk.Listbox(self.left_frame, width=30, height=10)
		self.listbox.pack()
		load_button = tk.Button(self.left_frame, text="Reload data", font=("Arial", 10), command=self.get_user_from_csv)
		load_button.pack(pady=10)

		self.status_entry = None

		# details on the right
		self.detail_vars = {}
		for i, field in enumerate(FIELDS):
			tk.Label(self.right_frame, text=field + ":").grid(row=i//2, column=0+2*(i%2),)
			var = tk.StringVar()
			entry = tk.Entry(self.right_frame, textvariable=var, width=30, state="readonly")
			entry.grid(row=i//2, column=1+2*(i%2), pady=5)

			if (i==len(FIELDS)-1):
				entry.config(fg='red')
				self.status_entry = entry

			self.detail_vars[field] = var


		# load data
		self.get_user_from_csv()

		# fill listbox (e.g., show name + phone)
		for k,record in self.shared_dict.items():

			display_text = f"{record['name']} ({record['arc']})"
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
				entry = self.manager.dict()
				new_entry = dict(zip(keys, row))

				for k,v in new_entry.items():
					entry[k] = v

				entry['status']= "Not Register Yet"

				self.shared_dict[entry['arc']] = entry

				count += 1

		self.show_details(None)

	def show_details(self, event):
		selection = self.listbox.curselection()
		if not selection:
			return
		index = selection[0]
		record = self.shared_dict.values()[index]
		for field in FIELDS:
			self.detail_vars[field].set(record.get(field, ""))

			if ("FAIL" in record.get(field, "")):
				self.status_entry.config(fg='red')

			if ("SUCCESS" in record.get(field, "")):
				self.status_entry.config(fg='green')

	def update_gui(self):
		# update clock
		self.time_label.config(text="Time: " + strftime("%H:%M:%S"))
		if (len(self.shared_log) > 0):
			self.log_box.delete(1.0, tk.END)
			# Insert new content
			self.log_box.insert(tk.END, "REPORT:\n")

			for i, data in enumerate(self.shared_log):
				self.log_box.insert(tk.END, f"{i + 1}. {data}")
			# Auto scroll to the bottom

			self.log_box.insert(tk.END, f"-- FINISH --")

		root.after(500, self.update_gui)  # repeat


	# --- Multiprocessing runner ---
	def run_processes(self):
		self.run_button['state'] = tk.DISABLED

		self.log_box.delete(1.0, tk.END)
		# Insert new content
		self.log_box.insert(tk.END, "Running auto registration. Please wait ...\n")

		num_users = len(self.shared_dict)
		processes = []
		for i in range(len(self.shared_dict)):
			p = Process(target=auto_regist, args=(i,self.shared_dict, num_users,self.shared_log,self.win_h, self.win_w))
			p.start()
			processes.append(p)
		for p in processes:
			p.join()

		print("DONE")
		self.show_details(None)
		self.run_button['state'] = tk.NORMAL

	# --- Run in background thread ---
	def run_in_thread(self):
		threading.Thread(target=self.run_processes, args=(), daemon=True).start()


if __name__ == "__main__":
	root = tk.Tk()
	app = App(root)
	app.update_gui()
	root.mainloop()

