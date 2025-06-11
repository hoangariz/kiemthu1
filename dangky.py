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
excel_path = r"C:\Users\Admin\Desktop\dangky.xlsx"

# Tải file Excel
workbook = load_workbook(excel_path)
sheet = workbook.active

# Khởi tạo Selenium WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()

# URL trang đăng nhập/đăng ký
login_url = "https://my.thanhnien.vn/page/login.html?redirect_url=https://thanhnien.vn/"
success_url = "https://my.thanhnien.vn/"


# Hàm xử lý giá trị ô
def process_cell_value(cell_value):
    if cell_value is None or not str(cell_value).strip():
        return "", "", "", ""
    # Chuyển giá trị thành chuỗi
    cell_value = str(cell_value)
    # Tách các trường bằng regex, giữ nguyên dấu cách trong ngoặc kép
    email_pattern = r'Email: "(.*?)"'
    name_pattern = r'Tên: "(.*?)"'
    password_pattern = r'Mật khẩu: "(.*?)"'
    confirm_pattern = r'Xác nhận: "(.*?)"'

    email_match = re.search(email_pattern, cell_value)
    name_match = re.search(name_pattern, cell_value)
    password_match = re.search(password_pattern, cell_value)
    confirm_match = re.search(confirm_pattern, cell_value)

    # Lấy giá trị gốc, không trim
    email = email_match.group(1) if email_match else ""
    name = name_match.group(1) if name_match else ""
    password = password_match.group(1) if password_match else ""
    confirm_password = confirm_match.group(1) if confirm_match else ""

    # Xử lý trường hợp "(trống)"
    if email.lower() == "(trống)":
        email = ""
    if name.lower() == "(trống)":
        name = ""
    if password.lower() == "(trống)":
        password = ""
    if confirm_password.lower() == "(trống)":
        confirm_password = ""

    return email, name, password, confirm_password


# Hàm đăng xuất
def logout(max_attempts=3):
    attempt = 1
    while attempt <= max_attempts:
        try:
            # Nhấn vào biểu tượng dropdown
            dropdown = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "span.icon-down"))
            )
            dropdown.click()
            time.sleep(0.5)  # Trễ 0,5 giây

            # Nhấn vào liên kết đăng xuất
            logout_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a#logout_user"))
            )
            logout_button.click()
            time.sleep(0.5)  # Trễ 0,5 giây

            # Chờ quay lại trang đăng nhập/đăng ký
            WebDriverWait(driver, 5).until(
                EC.url_contains("login.html")
            )
            print("Đăng xuất thành công, quay lại trang đăng nhập.")
            return True
        except:
            print(f"Thử đăng xuất lần {attempt}/{max_attempts} thất bại, thử lại.")
            driver.get(login_url)  # Tải lại trang đăng nhập/đăng ký
            time.sleep(0.5)  # Trễ 0,5 giây
            attempt += 1

    print(f"Không thể đăng xuất sau {max_attempts} lần thử, tiếp tục test case tiếp theo.")
    return False


# Đọc các test case từ Excel bắt đầu từ C2
test_cases = []
row = 2  # Bắt đầu từ hàng 2
while True:
    cell_value = sheet[f"C{row}"].value
    print(f"Debug - Row {row}: Cell='{cell_value}'")
    if cell_value is None or not str(cell_value).strip():
        break  # Dừng nếu ô trống
    email, name, password, confirm_password = process_cell_value(cell_value)
    test_cases.append((email, name, password, confirm_password, row))
    row += 1  # Chuyển sang ô tiếp theo (mỗi ô là một test case)

# Thực hiện các test đăng ký
for index, (email, name, password, confirm_password, row) in enumerate(test_cases, start=1):
    print(
        f"\nChạy Test Case {index}: Email='{email}', Tên='{name}', Mật khẩu='{password}', Xác nhận='{confirm_password}'")

    # Truy cập trang đăng ký
    driver.get(login_url)
    time.sleep(0.5)  # Trễ 0,5 giây

    # Nhấn vào tab Đăng ký
    try:
        register_tab = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "li.tabs-login-item a[href='#menu_2']"))
        )
        register_tab.click()
        time.sleep(0.5)  # Trễ 0,5 giây
    except:
        print(f"Test Case {index}: Không tìm thấy tab Đăng ký, bỏ qua test case.")
        sheet[f"E{row}"].value = "Lỗi: Không tìm thấy tab Đăng ký"
        continue

    # Tìm ô nhập email và nhập giá trị
    try:
        email_input = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[data-ng-model='registration.email']"))
        )
        email_input.clear()
        time.sleep(0.5)  # Trễ 0,5 giây
        if email:  # Chỉ nhập nếu email không rỗng
            email_input.send_keys(email)
        time.sleep(0.5)  # Trễ 0,5 giây
    except:
        print(f"Test Case {index}: Không tìm thấy ô Email, bỏ qua test case.")
        sheet[f"E{row}"].value = "Lỗi: Không tìm thấy ô Email"
        continue

    # Tìm ô nhập tên và nhập giá trị
    try:
        name_input = driver.find_element(By.CSS_SELECTOR, "input[data-ng-model='registration.displayName']")
        name_input.clear()
        time.sleep(0.5)  # Trễ 0,5 giây
        if name:  # Chỉ nhập nếu tên không rỗng
            name_input.send_keys(name)
        time.sleep(0.5)  # Trễ 0,5 giây
    except:
        print(f"Test Case {index}: Không tìm thấy ô Tên, bỏ qua test case.")
        sheet[f"E{row}"].value = "Lỗi: Không tìm thấy ô Tên"
        continue

    # Tìm ô nhập mật khẩu và nhập giá trị
    try:
        password_input = driver.find_element(By.CSS_SELECTOR, "input[data-ng-model='registration.password']")
        password_input.clear()
        time.sleep(0.5)  # Trễ 0,5 giây
        if password:  # Chỉ nhập nếu mật khẩu không rỗng
            password_input.send_keys(password)
        time.sleep(0.5)  # Trễ 0,5 giây
    except:
        print(f"Test Case {index}: Không tìm thấy ô Mật khẩu, bỏ qua test case.")
        sheet[f"E{row}"].value = "Lỗi: Không tìm thấy ô Mật khẩu"
        continue

    # Tìm ô nhập xác nhận mật khẩu và nhập giá trị
    try:
        confirm_input = driver.find_element(By.CSS_SELECTOR, "input[data-ng-model='registration.confirmPassword']")
        confirm_input.clear()
        time.sleep(0.5)  # Trễ 0,5 giây
        if confirm_password:  # Chỉ nhập nếu xác nhận không rỗng
            confirm_input.send_keys(confirm_password)
        time.sleep(0.5)  # Trễ 0,5 giây
    except:
        print(f"Test Case {index}: Không tìm thấy ô Xác nhận mật khẩu, bỏ qua test case.")
        sheet[f"E{row}"].value = "Lỗi: Không tìm thấy ô Xác nhận mật khẩu"
        continue

    # Nhấn nút đăng ký
    try:
        register_button = driver.find_element(By.CSS_SELECTOR, "a[data-ng-click='signUp()']")
        register_button.click()
        time.sleep(0.5)  # Trễ 0,5 giây
    except:
        print(f"Test Case {index}: Không tìm thấy nút Đăng ký, bỏ qua test case.")
        sheet[f"E{row}"].value = "Lỗi: Không tìm thấy nút Đăng ký"
        continue

    # Bắt thông báo toast (nếu có)
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

    # Ghi kết quả toast vào cột E
    sheet[f"E{row}"].value = toast_msg

    # Kiểm tra chuyển hướng để xác định đăng ký thành công
    try:
        WebDriverWait(driver, 5).until(
            EC.url_to_be(success_url)
        )
        print(f"Test Case {index}: Chuyển hướng đến {success_url}, đăng ký thành công.")
        # Thực hiện đăng xuất
        if not logout():
            sheet[f"E{row}"].value = f"{toast_msg} (Không thể đăng xuất sau 3 lần thử)"
    except:
        print(f"Test Case {index}: Không chuyển hướng đến {success_url}, không đăng xuất.")

# Lưu file Excel
workbook.save(excel_path)

# Đóng trình duyệt
driver.quit()
print(f"\nHoàn thành {len(test_cases)} test case. Kết quả đã được ghi vào cột E của file Excel.")