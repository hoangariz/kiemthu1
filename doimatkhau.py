from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
import time
import os
import re

# Đường dẫn đến file Excel
excel_path = r"C:\Users\Admin\Desktop\doimatkhau.xlsx"

# Tải file Excel
workbook = load_workbook(excel_path)
sheet = workbook.active

# Khởi tạo Selenium WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()

# URL trang đăng nhập và đổi mật khẩu
login_url = "https://my.thanhnien.vn/page/login.html?redirect_url=https://thanhnien.vn/"
change_password_url = "https://my.thanhnien.vn/doi-mat-khau"


# Hàm đăng nhập
def login(email="kiemthu@vomoto.com", password="123456"):
    try:
        driver.get(login_url)
        time.sleep(0.5)

        email_input = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input.input-email"))
        )
        email_input.clear()
        time.sleep(0.5)
        email_input.send_keys(email)
        time.sleep(0.5)

        password_input = driver.find_element(By.CSS_SELECTOR, "input.input-password")
        password_input.clear()
        time.sleep(0.5)
        password_input.send_keys(password)
        time.sleep(0.5)

        login_button = driver.find_element(By.CSS_SELECTOR, "div.btn-login a.link-btn")
        login_button.click()
        time.sleep(0.5)

        toast_msg_element = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "p.toast__msg"))
        )
        toast_msg = toast_msg_element.text
        if "Bạn đã đăng nhập thành công." in toast_msg:
            print("Đăng nhập thành công.")
            return True
        else:
            print(f"Đăng nhập thất bại: {toast_msg}")
            return False
    except:
        print("Lỗi khi đăng nhập.")
        return False


# Hàm đăng xuất
def logout(max_attempts=3):
    attempt = 1
    while attempt <= max_attempts:
        try:
            dropdown = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "span.icon-down"))
            )
            dropdown.click()
            time.sleep(0.5)

            logout_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a#logout_user"))
            )
            logout_button.click()
            time.sleep(0.5)

            WebDriverWait(driver, 5).until(
                EC.url_contains("login.html")
            )
            print("Đăng xuất thành công, quay lại trang đăng nhập.")
            return True
        except:
            print(f"Thử đăng xuất lần {attempt}/{max_attempts} thất bại, thử lại.")
            driver.get(login_url)
            time.sleep(0.5)
            attempt += 1

    print(f"Không thể đăng xuất sau {max_attempts} lần thử.")
    return False


# Hàm đổi mật khẩu về mặc định
def reset_password(old_password, default_password="123456"):
    try:
        driver.get(change_password_url)
        time.sleep(0.5)

        old_pass_input = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[data-ng-model='userInfo.oldPass']"))
        )
        old_pass_input.clear()
        time.sleep(0.5)
        old_pass_input.send_keys(old_password)
        time.sleep(0.5)

        new_pass_input = driver.find_element(By.CSS_SELECTOR, "input[data-ng-model='userInfo.newPass']")
        new_pass_input.clear()
        time.sleep(0.5)
        new_pass_input.send_keys(default_password)
        time.sleep(0.5)

        confirm_pass_input = driver.find_element(By.CSS_SELECTOR, "input[data-ng-model='userInfo.confirmPass']")
        confirm_pass_input.clear()
        time.sleep(0.5)
        confirm_pass_input.send_keys(default_password)
        time.sleep(0.5)

        save_button = driver.find_element(By.CSS_SELECTOR, "a.btn-save")
        save_button.click()
        time.sleep(0.5)

        toast_msg_element = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "p.toast__msg"))
        )
        toast_msg = toast_msg_element.text
        print(f"Reset mật khẩu: {toast_msg}")
        return toast_msg == "Mật khẩu của bạn đã được thay đổi."
    except:
        print("Lỗi khi reset mật khẩu.")
        return False


# Hàm xử lý giá trị ô
def process_cell_value(cell_value):
    if cell_value is None or not str(cell_value).strip():
        return "", "", ""
    cell_value = str(cell_value)
    old_pass_pattern = r'Mật khẩu cũ: "(.*?)"'
    new_pass_pattern = r'Mật khẩu mới: "(.*?)"'
    confirm_pattern = r'Nhập lại: "(.*?)"'

    old_pass_match = re.search(old_pass_pattern, cell_value)
    new_pass_match = re.search(new_pass_pattern, cell_value)
    confirm_match = re.search(confirm_pattern, cell_value)

    old_password = old_pass_match.group(1) if old_pass_match else ""
    new_password = new_pass_match.group(1) if new_pass_match else ""
    confirm_password = confirm_match.group(1) if confirm_match else ""

    if old_password.lower() == "(trống)":
        old_password = ""
    if new_password.lower() == "(trống)":
        new_password = ""
    if confirm_password.lower() == "(trống)":
        confirm_password = ""

    return old_password, new_password, confirm_password


# Đăng nhập ban đầu
if not login():
    print("Không thể đăng nhập, thoát chương trình.")
    driver.quit()
else:
    # Đọc các test case từ Excel
    test_cases = []
    row = 2
    while True:
        cell_value = sheet[f"C{row}"].value
        print(f"Debug - Row {row}: Cell='{cell_value}'")
        if cell_value is None or not str(cell_value).strip():
            break
        old_password, new_password, confirm_password = process_cell_value(cell_value)
        test_cases.append((old_password, new_password, confirm_password, row))
        row += 1

    # Thực hiện các test đổi mật khẩu
    for index, (old_password, new_password, confirm_password, row) in enumerate(test_cases, start=1):
        print(
            f"\nChạy Test Case {index}: Mật khẩu cũ='{old_password}', Mật khẩu mới='{new_password}', Nhập lại='{confirm_password}'")

        # Truy cập trang đổi mật khẩu
        driver.get(change_password_url)
        time.sleep(0.5)

        # Tìm ô nhập mật khẩu cũ
        try:
            old_pass_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[data-ng-model='userInfo.oldPass']"))
            )
            old_pass_input.clear()
            time.sleep(0.5)
            if old_password:
                old_pass_input.send_keys(old_password)
            time.sleep(0.5)
        except:
            print(f"Test Case {index}: Không tìm thấy ô Mật khẩu cũ, bỏ qua test case.")
            sheet[f"E{row}"].value = "Lỗi: Không tìm thấy ô Mật khẩu cũ"
            continue

        # Tìm ô nhập mật khẩu mới
        try:
            new_pass_input = driver.find_element(By.CSS_SELECTOR, "input[data-ng-model='userInfo.newPass']")
            new_pass_input.clear()
            time.sleep(0.5)
            if new_password:
                new_pass_input.send_keys(new_password)
            time.sleep(0.5)
        except:
            print(f"Test Case {index}: Không tìm thấy ô Mật khẩu mới, bỏ qua test case.")
            sheet[f"E{row}"].value = "Lỗi: Không tìm thấy ô Mật khẩu mới"
            continue

        # Tìm ô nhập xác nhận mật khẩu
        try:
            confirm_pass_input = driver.find_element(By.CSS_SELECTOR, "input[data-ng-model='userInfo.confirmPass']")
            confirm_pass_input.clear()
            time.sleep(0.5)
            if confirm_password:
                confirm_pass_input.send_keys(confirm_password)
            time.sleep(0.5)
        except:
            print(f"Test Case {index}: Không tìm thấy ô Nhập lại mật khẩu, bỏ qua test case.")
            sheet[f"E{row}"].value = "Lỗi: Không tìm thấy ô Nhập lại mật khẩu"
            continue

        # Nhấn nút lưu thay đổi
        try:
            save_button = driver.find_element(By.CSS_SELECTOR, "a.btn-save")
            save_button.click()
            time.sleep(0.5)
        except:
            print(f"Test Case {index}: Không tìm thấy nút Lưu thay đổi, bỏ qua test case.")
            sheet[f"E{row}"].value = "Lỗi: Không tìm thấy nút Lưu thay đổi"
            continue

        # Bắt thông báo toast
        toast_msg = ""
        try:
            toast_msg_element = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "p.toast__msg"))
            )
            toast_msg = toast_msg_element.text
            print(f"Test Case {index}: Toast message - {toast_msg}")
        except:
            toast_msg = "Không bắt được toast"
            print(f"Test Case {index}: {toast_msg}")

        # Ghi kết quả vào cột E
        sheet[f"E{row}"].value = toast_msg

        # Nếu đổi mật khẩu thành công, reset về mật khẩu mặc định
        if toast_msg == "Mật khẩu của bạn đã được thay đổi.":
            print(f"Test Case {index}: Đổi mật khẩu thành công, reset về mật khẩu mặc định.")
            if not reset_password(new_password):
                sheet[f"E{row}"].value = f"{toast_msg} (Lỗi khi reset mật khẩu về 123456)"
        else:
            print(f"Test Case {index}: Đổi mật khẩu thất bại, tải lại trang.")
            driver.get(change_password_url)
            time.sleep(0.5)

# Lưu file Excel
workbook.save(excel_path)

# Đóng trình duyệt
driver.quit()
print(f"\nHoàn thành {len(test_cases)} test case. Kết quả đã được ghi vào cột E của file Excel.")