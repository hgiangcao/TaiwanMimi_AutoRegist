import tkinter as tk
from tkinter.scrolledtext import ScrolledText
from time import strftime
import multiprocessing
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


# --- Selenium automation function ---
def auto_regist(i, user, n_user, shared_log):
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

        root = tk.Tk()
        screen_width = root.winfo_screenwidth() - 20
        screen_height = root.winfo_screenheight() - 20
        root.destroy()

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
        except:
            print(user['arc'], user['name'], "FAIL: Cannot select LOCATION")
            shared_log.append(f"{user['arc']} {user['name']} FAIL: Wrong LOCATION\n")
            

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

            shared_log.append(f"{user['arc']} {user['name']} SUCCESS!!!\n")
            # # 等待一些時間以便觀察
            time.sleep(120)
            # # 完成後 關閉瀏覽器
            driver.quit()
        except:
            driver.quit()
            # print(f"{user['arc']} {user['name']} FAIL: Cannot select Exam Date - Time - Group")
            shared_log.append(f"{user['arc']} {user['name']} FAIL: Date,Time,Group NOT AVAILABLE  \n")
            

    except Exception as e:
        shared_log.append(f"{user['arc']} {user['name']} FAIL: {str(e)}\n")
        print(f"{user['arc']} {user['name']} FAIL: {str(e)}\n")
        

# --- CSV reading ---
def get_user_from_csv(file_name="users.csv"):
    keys = [
        "register_date", "district", "address", "exam_date",
        "time", "group", "arc", "birthday", "name", "phone", "email"
    ]
    result = []
    log=""
    count =0
    with open(file_name, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        for row in reader:
            if row[-1] != "DONE":
                entry = dict(zip(keys, row))
                result.append(entry)
                log += f"{count+1}. {entry['arc']} {entry['name']}\n"
                count+=1

    log_box.delete(1.0, tk.END)
    log_box.insert(tk.END, f"Found {count} records:\n")
    log_box.insert(tk.END, log)
    log_box.see(tk.END)
    log_box.update_idletasks()

    return result

def update_log():
    if (len(shared_log)>0):
        # Clear current content
        log_box.delete(1.0, tk.END)
        # Insert new content
        log_box.insert(tk.END, "REPORT:\n")

        for i,data in enumerate(shared_log):
            log_box.insert(tk.END,f"{i+1}. {data}")
        # Auto scroll to the bottom

        log_box.insert(tk.END,f"-- FINISH --")

        log_box.see(tk.END)

# --- Multiprocessing runner ---
def run_processes(shared_log):
    users = get_user_from_csv()
    num_users = len(users)

    print("records",num_users)

    processes = []
    for i in range(num_users):
        p = multiprocessing.Process(target=auto_regist, args=(i, users[i], num_users, shared_log))
        p.start()
        processes.append(p)
    for p in processes:
        p.join()




# --- Run in background thread ---
def run_in_thread(shared_log):
    threading.Thread(target=run_processes, args=(shared_log,), daemon=True).start()

# --- Tkinter GUI ---


if __name__ == "__main__":
    manager = multiprocessing.Manager()
    shared_log = manager.list()  # shared list for logging

    root = tk.Tk()
    root.title("Clock & Run")
    root.geometry("500x450")
    root.attributes("-topmost", True)
    root.resizable(False, False)

    # Clock label
    time_label = tk.Label(root, font=("Arial", 15),fg='red')
    time_label.pack(pady=10)

    # Run button
    run_button = tk.Button(root, text="Run", font=("Arial", 10), command=lambda: run_in_thread(shared_log))
    run_button.pack(pady=10)

    # Log label (multi-line)
    log_box = ScrolledText(root, width=60, height=10, wrap=tk.WORD)
    log_box.config(spacing1=5)
    log_box.config(spacing2=5)
    log_box.config(spacing3=5)
    log_box.pack(pady=10)

    load_button = tk.Button(root, text="Reload data", font=("Arial", 10), command=get_user_from_csv)
    load_button.pack(pady=10)

    get_user_from_csv()

    # Update clock and log
    def update_gui():
        # update clock
        time_label.config(text="Time: " + strftime("%H:%M:%S"))

        update_log()

        root.after(500, update_gui)  # repeat

    update_gui()
    root.mainloop()
