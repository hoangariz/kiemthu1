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
excel_path = r"C:\Users\Admin\Desktop\dangnhap.xlsx"

# Tải file Excel
workbook = load_workbook(excel_path)
sheet = workbook.active

# Khởi tạo Selenium WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()

# URL trang đăng nhập
login_url = "https://my.thanhnien.vn/page/login.html?redirect_url=https://thanhnien.vn/"


# Hàm xử lý giá trị ô
def process_cell_value(cell_value):
    if cell_value is None or not str(cell_value).strip():
        return "", ""
    # Chuyển giá trị thành chuỗi
    cell_value = str(cell_value)
    # Tách email và mật khẩu bằng regex, giữ nguyên dấu cách trong ngoặc kép
    email_pattern = r'Email: "(.*?)"'
    password_pattern = r'Mật khẩu: "(.*?)"'
    email_match = re.search(email_pattern, cell_value)
    password_match = re.search(password_pattern, cell_value)

    # Lấy giá trị gốc, không trim
    email = email_match.group(1) if email_match else ""
    password = password_match.group(1) if password_match else ""

    # Xử lý trường hợp "(trống)"
    if email.lower() == "(trống)":
        email = ""
    if password.lower() == "(trống)":
        password = ""

    return email, password


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

            # Chờ quay lại trang đăng nhập
            WebDriverWait(driver, 5).until(
                EC.url_contains("login.html")
            )
            print("Đăng xuất thành công, quay lại trang đăng nhập.")
            return True
        except:
            print(f"Thử đăng xuất lần {attempt}/{max_attempts} thất bại, thử lại.")
            driver.get(login_url)  # Tải lại trang đăng nhập
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
    email, password = process_cell_value(cell_value)
    test_cases.append((email, password, row))
    row += 1  # Chuyển sang ô tiếp theo (mỗi ô là một test case)

# Thực hiện các test đăng nhập
for index, (email, password, row) in enumerate(test_cases, start=1):
    print(f"\nChạy Test Case {index}: Email='{email}', Mật khẩu='{password}'")

    # Truy cập trang đăng nhập
    driver.get(login_url)
    time.sleep(0.5)  # Trễ 0,5 giây

    # Tìm ô nhập email và nhập giá trị
    try:
        email_input = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input.input-email"))
        )
        email_input.clear()
        time.sleep(0.5)  # Trễ 0,5 giây
        if email:  # Chỉ nhập nếu email không rỗng
            email_input.send_keys(email)
        time.sleep(0.5)  # Trễ 0,5 giây
    except:
        print(f"Test Case {index}: Không tìm thấy ô email, bỏ qua test case.")
        sheet[f"E{row}"].value = "Lỗi: Không tìm thấy ô email"
        continue

    # Tìm ô nhập mật khẩu và nhập giá trị
    try:
        password_input = driver.find_element(By.CSS_SELECTOR, "input.input-password")
        password_input.clear()
        time.sleep(0.5)  # Trễ 0,5 giây
        if password:  # Chỉ nhập nếu mật khẩu không rỗng
            password_input.send_keys(password)
        time.sleep(0.5)  # Trễ 0,5 giây
    except:
        print(f"Test Case {index}: Không tìm thấy ô mật khẩu, bỏ qua test case.")
        sheet[f"E{row}"].value = "Lỗi: Không tìm thấy ô mật khẩu"
        continue

    # Nhấn nút đăng nhập
    try:
        login_button = driver.find_element(By.CSS_SELECTOR, "div.btn-login a.link-btn")
        login_button.click()
        time.sleep(0.5)  # Trễ 0,5 giây
    except:
        print(f"Test Case {index}: Không tìm thấy nút đăng nhập, bỏ qua test case.")
        sheet[f"E{row}"].value = "Lỗi: Không tìm thấy nút đăng nhập"
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

    # Nếu đăng nhập thành công, đăng xuất
    if toast_msg == "Bạn đã đăng nhập thành công.":
        if not logout():
            sheet[f"E{row}"].value = f"{toast_msg} (Không thể đăng xuất sau 3 lần thử)"
            driver.get(login_url)  # Tải lại trang để thử test case tiếp theo

# Lưu file Excel
workbook.save(excel_path)

# Đóng trình duyệt
driver.quit()
print(f"\nHoàn thành {len(test_cases)} test case. Kết quả đã được ghi vào cột E của file Excel.")