from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
import pandas as pd
import time

PATH = "C:/Program Files (x86)/chromedriver.exe"
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=options)

driver.get("https://assessprice.treasury.go.th")

data = []

try:
    # รอจนกว่า overlay จะหายไป
    WebDriverWait(driver, 10).until(
        EC.invisibility_of_element_located((By.CLASS_NAME, "amos-panel-underlay"))
    )

    # รอและคลิกปุ่ม close
    close_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "amos-dialog-icon-close"))
    )
    close_button.click()

except Exception as e:
    print("เกิดข้อผิดพลาดในการคลิก close button:", e)

try:
    # รอและคลิกปุ่ม hamburger โดยใช้ JavaScript
    hamburger_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "hamburger"))
    )
    driver.execute_script("arguments[0].click();", hamburger_button)

except Exception as e:
    print("เกิดข้อผิดพลาดในการคลิก hamburger button:", e)

try:
    # รอและคลิกปุ่มห้องชุด
    room_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//div[@class='label GIS5-WF05' and text()='ห้องชุด']"))
    )
    driver.execute_script("arguments[0].click();", room_button)

except Exception as e:
    print("เกิดข้อผิดพลาดในการคลิกปุ่มห้องชุด:", e)

try:
    # รอจนกว่า input ที่มี id เป็น "uniqName_19_0" จะปรากฏและพร้อมใช้งาน
    input_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "uniqName_19_0"))
    )
    # กรอกข้อความ "กรุงเทพมหานคร" ลงในช่อง input
    input_element.send_keys("กรุงเทพมหานคร")

    # รอจนกระทั่งปุ่มค้นหาปรากฏขึ้นและสามารถคลิกได้
    search_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@data-dojo-attach-point='labelContainer' and @class='amos-button-label-container']"))
    )
    # คลิกที่ปุ่มค้นหา
    driver.execute_script("arguments[0].click();", search_button)

    # วนลูปเพื่อดึงข้อมูลจากแต่ละหน้า
    while True:
        # รอให้ผลการค้นหาในหน้าปัจจุบันปรากฏขึ้น
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "amos-result-title"))
        )

        # ดึงข้อมูลจากทุก div ที่มี class "amos-result-title"
        result_titles = driver.find_elements(By.CLASS_NAME, "amos-result-title")
        for title in result_titles:
            data.append(title.text)

        # ตรวจสอบว่ามีปุ่ม "ถัดไป" หรือไม่
        try:
            next_button = driver.find_element(By.XPATH, "//span[@class='amos-icon-angle-right']")
            driver.execute_script("arguments[0].click();", next_button)  # คลิกที่ปุ่ม "ถัดไป" โดยใช้ JavaScript
            time.sleep(2)  # รอให้หน้าใหม่โหลดขึ้นมา
        except StaleElementReferenceException:
            print("พบปัญหา Stale Element Reference - กำลังพยายามใหม่...")
            continue  # ลองใหม่หากเกิด StaleElementReferenceException
        except:
            break  # ถ้าไม่พบปุ่ม "ถัดไป" ให้ออกจาก loop

except Exception as e:
    print("เกิดข้อผิดพลาด:", e)

# สร้าง DataFrame และบันทึกเป็นไฟล์ Excel เสมอไม่ว่าจะเกิดข้อผิดพลาดหรือไม่
df = pd.DataFrame(data, columns=["Result Title"])
df.to_excel("results.xlsx", index=False)
print("ข้อมูลถูกบันทึกลงในไฟล์ results.xlsx เรียบร้อยแล้ว")

# driver.quit() # ปิด browser หลังใช้งาน



