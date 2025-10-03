import tkinter as tk
from tkinter.scrolledtext import ScrolledText
from time import strftime
from multiprocessing import Process, Manager
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

import re

session_map = {
	"1": "上午場次",
	"2": "下午場次"
}


# --- Selenium automation function ---
def check(user,shared_dict):
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
		options.add_argument("--headless")
		options.add_argument("--log-level=3")
		options.add_experimental_option('excludeSwitches', ['enable-logging'])
		# try:
		# 	s = Service(ChromeDriverManager().install())
		# except:
		# 	try:
		# 		s = Service("chromedriver/chromedriver.exe")
		# 	except:
		# 		print("CHROME DRIVER ERROR")

		s = Service(ChromeDriverManager().install())

		driver = webdriver.Chrome(service=s, options=options)

		# 打開網頁
		url = "https://www.mvdis.gov.tw/m3-emv-trn/exm/query#anchor&gsc.tab=0"
		driver.get(url)
		try:

			# 等待「報考照類」下拉選單可點擊
			WebDriverWait(driver, 5).until(
				EC.element_to_be_clickable((By.ID, 'idNo'))
			)

			# 定位預計考試日期輸入框並填寫
			date_input = driver.find_element(By.ID, 'idNo')
			date_input.clear()
			date_input.send_keys(user['arc'])  # 填入預計考試日期

			# 定位預計考試日期輸入框並填寫
			date_input = driver.find_element(By.ID, 'birthdayStr')
			date_input.clear()
			date_input.send_keys(user['birthday'])  # 填入預計考試日期

			# 定位「查詢場次」按鈕並點擊
			search_button = driver.find_element(By.CSS_SELECTOR, "a[onclick='query();']")

			search_button.click()
			time.sleep(5)
			# 等待彈出視窗元素加載完成
			# 嘗試直接搜尋整個文字「取消報名 Cancel」
			try:
				cancel_elem = driver.find_element(By.XPATH, "//*[text()='取消報名 Cancel']")
				shared_dict[user['arc']] = "DONE"
				driver.quit()
				return True
			except:
				shared_dict[user['arc']] = "FAIL"
				driver.quit()
				return False

			driver.quit()
		except Exception as e:
			driver.quit()
			print("error 2")
			print (e)


	except Exception as e:
		print("error 1")
		print (e)



def run_in_thread():
	threading.Thread(target=run_processes, args=(), daemon=True).start()

def run_processes():
	status_label.config(text="Checking ...")
	status_label.update_idletasks()
	manager = Manager()
	shared_dict = manager.dict()
	processes = []

	for k,user in users.items():
		p = Process(target=check, args=(user,shared_dict))
		p.start()
		processes.append(p)
	for p in processes:
		p.join()


	for k,v in shared_dict.items():
		user = users[k]
		treeview.item(k,values=(
		v, user['name'], user['arc'], user['phone'], user["district"], user["address"], user["exam_date"],
		user["time"], user["group"]))

		treeview.update_idletasks()

	status_label.config(text="DONE ...")
	status_label.update_idletasks()


def read_google_sheet_csv( csv_url):
	resp = requests.get(csv_url)
	resp.raise_for_status()  # raise error if request fails

	# Ensure UTF-8 decoding for Chinese characters
	text = resp.content.decode('utf-8')  # decode bytes to string
	f = io.StringIO(text)

	reader = csv.reader(f)
	rows = list(reader)
	return rows


def get_user_from_csv():
	treeview.delete(*treeview.get_children())
	keys = [
		"register_date", "district", "address", "exam_date",
		"time", "group", "arc", "birthday", "name", "phone", "email"
	]
	url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTMfUPlcL32jW93PhYBCIBOPTj1iIZRzVBBqqPSh1rM4cpR56ss2aYWA1zJjmlfbq9IAbmw9p6RAbRB/pub?output=csv"
	rows = read_google_sheet_csv(url)

	data = rows[1:]

	for row in data:
		entry = dict(zip(keys, row))
		entry['status'] = ""
		users[entry['arc']]= entry

		treeview.insert('', tk.END,iid=entry['arc'], values=("",entry['name'],entry['arc'],entry['phone'],entry["district"],entry["address"],entry["exam_date"],entry["time"], entry["group"]))

users = {}


from tkinter import ttk

if __name__=="__main__":

	root = tk.Tk()
	root.title("Check DONE register")
	root.geometry("1000x450")
	root.attributes("-topmost", True)
	# root.resizable(False, False)



	# Run button
	run_button = tk.Button(root, text="Check", font=("Arial", 10), command=lambda: run_in_thread())
	run_button.pack(pady=10)

	status_label = tk.Label(root, font=("Arial", 10), fg='red')
	status_label.pack(pady=5)

	treeview = ttk.Treeview()

	treeview = ttk.Treeview(columns=('STATUS','name','arc','phone',"District","Address","Date","Time", "Group"), show="headings")

	treeview.heading("STATUS", text="STATUS")
	treeview.heading("name", text="name")
	treeview.heading("phone", text="phone")
	treeview.heading("arc", text="arc")
	treeview.heading("District", text="District")
	treeview.heading("Address", text="Address")
	treeview.heading("Date", text="Date")
	treeview.heading("Time", text="Time")
	treeview.heading("Group", text="Group")


	treeview.column("STATUS", width=50)
	treeview.column("name", width=50)
	treeview.column("phone", width=70)
	treeview.column("arc", width=70)
	treeview.column("District", width=100)
	treeview.column("District", width=200)
	treeview.column("Address", width=250)
	treeview.column("Date", width=120)
	treeview.column("Time", width=50)
	treeview.column("Group", width=50)

	treeview.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)

	get_user_from_csv()
	root.mainloop()
