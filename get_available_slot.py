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
def get_valid_slot(date='1141030',district='新竹區監理所（桃竹苗）',shared_dict=None):
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
        try:
            s = Service(ChromeDriverManager().install())
        except:
            try:
                s = Service("chromedriver/chromedriver.exe")
            except:
                print("CHROME DRIVER ERROR")
        driver = webdriver.Chrome(service=s, options=options)

        # 打開網頁
        driver.get("https://www.mvdis.gov.tw/m3-emv-trn/exm/locations#gsc.tab=0")
        try:

            # 等待「報考照類」下拉選單可點擊
            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, 'licenseTypeCode'))
            )

            # 定位「報考照類」並選擇
            select_license_type = Select(driver.find_element(By.ID, 'licenseTypeCode'))
            select_license_type.select_by_visible_text('普通重型機車')  # 選擇普通重型機車

            # 定位預計考試日期輸入框並填寫
            date_input = driver.find_element(By.ID, 'expectExamDateStr')
            date_input.clear()
            date_input.send_keys(date)  # 填入預計考試日期

            # 定位第一個考試地點選單並選擇臺北區監理所（北宜花）
            select_region = Select(driver.find_element(By.ID, 'dmvNoLv1'))
            select_region.select_by_visible_text(district)
            # wait until dropdown has more than 1 option
            WebDriverWait(driver, 20).until(
                lambda d: len(Select(d.find_element(By.ID, "dmvNo")).options) > 1
            )

            select_station = Select(driver.find_element(By.ID, 'dmvNo'))
            options = select_station.options

            # extract text
            all_address = [opt.text for opt in options]

            for address in all_address[1:]:

                shared_dict[district][address] = []

                # print(district, address)
                select_station = Select(driver.find_element(By.ID, 'dmvNo'))
                select_station.select_by_visible_text(address)
                # 定位「查詢場次」按鈕並點擊
                search_button = driver.find_element(By.CSS_SELECTOR, "a[onclick='query();']")

                search_button.click()
                time.sleep(1)

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

                # 等待彈出視窗元素加載完成
                # wait until table appears
                table = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "trnTable"))
                )

                # find all <a> with class="std_btn" inside the table
                links = table.find_elements(By.CSS_SELECTOR, "a.std_btn")

                # get all rows (skip header if any)
                rows = table.find_elements(By.TAG_NAME, "tr")

                results = []
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if not cells:
                        continue

                    # Example structure: [date, session info, slots, link]
                    date_text = cells[0].text.strip()
                    session_info = cells[1].text.strip()
                    available_slots = cells[2].text.strip()

                    if (available_slots!="額滿"):
                        # get link and onclick parameters
                        link = cells[3].find_element(By.CSS_SELECTOR, "a.std_btn")
                        onclick_value = link.get_attribute("onclick")
                        params = re.findall(r"'([^']*)'", onclick_value)  # extract values inside quotes

                        results.append({
                            "date": date_text,
                            "session": params[-2],
                            "group": params[-1],
                            "available_slots": available_slots,
                        })

                shared_dict[district][address]= results

                        # show result


                # print("Available slots:", len(results))
                # for r in results:
                #     print(r)

            driver.quit()
        except Exception as e:
            driver.quit()
            print("error 2")
            print (e)


    except Exception as e:
        print("error 1")
        print (e)

# --- Run in background thread ---
def run_in_thread(manager,shared_dict):
    threading.Thread(target=run_processes, args=(manager,shared_dict), daemon=True).start()

def run_processes(manager,shared_dict):
    list_district = ['臺北市區監理所（含金門馬祖）','臺北區監理所（北宜花）','新竹區監理所（桃竹苗）','臺中區監理所（中彰投）','嘉義區監理所（雲嘉南）','高雄市區監理所','高雄區監理所（高屏澎東）']
    processes = []

    status_label.config(text="Checking ...")
    status_label.update_idletasks()

    for district in list_district:
        shared_dict[district]= manager.dict()
        p = Process(target=get_valid_slot, args=('1141027',district,shared_dict))
        p.start()
        processes.append(p)
    for p in processes:
        p.join()

    print("DONE")
    status_label.config(text="Done!")
    status_label.update_idletasks()

    for k,v in shared_dict.items():
        for k2, v2 in v.items():
            for item in v2:
                treeview.insert('', tk.END, values=( k,k2,item['date'],session_map[item['session']],"組別 "+item['group'],item['available_slots'] ))

    treeview.update_idletasks()


from tkinter import ttk

if __name__=="__main__":
    manager = Manager()
    shared_dict = manager.dict()

    root = tk.Tk()
    root.title("Check Available Slots")
    root.geometry("850x450")
    root.attributes("-topmost", True)
    # root.resizable(False, False)



    # Run button
    run_button = tk.Button(root, text="Check", font=("Arial", 10), command=lambda: run_in_thread(manager,shared_dict))
    run_button.pack(pady=10)

    status_label = tk.Label(root, font=("Arial", 10), fg='red')
    status_label.pack(pady=5)

    treeview = ttk.Treeview()

    treeview = ttk.Treeview(columns=("District","Address","Date","Time", "Group","Avai. Slots"), show="headings")

    treeview.heading("District", text="District")
    treeview.heading("Address", text="Address")
    treeview.heading("Date", text="Date")
    treeview.heading("Time", text="Time")
    treeview.heading("Group", text="Group")
    treeview.heading("Avai. Slots", text="Avai. Slots")

    treeview.column("District", width=200)
    treeview.column("Address", width=250)
    treeview.column("Date", width=120)
    treeview.column("Time", width=50)
    treeview.column("Group", width=50)
    treeview.column("Avai. Slots", width=50)

    treeview.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)

    root.mainloop()
