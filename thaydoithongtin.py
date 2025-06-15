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
excel_path = r"C:\Users\Admin\Desktop\thaydoithongtin.xlsx"

# Tải file Excel
workbook = load_workbook(excel_path)
sheet = workbook.active

# Khởi tạo Selenium WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()

# URL trang đăng nhập và thông tin tài khoản
login_url = "https://my.thanhnien.vn/page/login.html?redirect_url=https://thanhnien.vn/"
profile_url = "https://my.thanhnien.vn/"


# Hàm đăng nhập
def login(email="kiemthu1@vomoto.com", password="123456"):
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


# Hàm xử lý giá trị ô
def process_cell_value(cell_value):
    if cell_value is None or not str(cell_value).strip():
        print("Ô Excel rỗng hoặc chỉ chứa khoảng trắng.")
        return "", "", "", "", ""

    cell_value = str(cell_value).strip()
    print(f"Debug - Giá trị ô: '{cell_value}'")

    # Regex để tách các trường, giữ nguyên dấu cách trong ngoặc kép
    name_pattern = r'Tên: "(.*?)"'
    gender_pattern = r'Giới tính: "(.*?)"'
    dob_pattern = r'Ngày sinh: "(.*?)"'
    phone_pattern = r'Điện thoại: "(.*?)"'
    address_pattern = r'Địa chỉ: "(.*?)"'

    name_match = re.search(name_pattern, cell_value, re.DOTALL)
    gender_match = re.search(gender_pattern, cell_value, re.DOTALL)
    dob_match = re.search(dob_pattern, cell_value, re.DOTALL)
    phone_match = re.search(phone_pattern, cell_value, re.DOTALL)
    address_match = re.search(address_pattern, cell_value, re.DOTALL)

    name = name_match.group(1) if name_match else ""
    gender = gender_match.group(1) if gender_match else ""
    dob = dob_match.group(1) if dob_match else ""
    phone = phone_match.group(1) if phone_match else ""
    address = address_match.group(1) if address_match else ""

    print(
        f"Debug - Sau khi tách: Tên='{name}', Giới tính='{gender}', Ngày sinh='{dob}', Điện thoại='{phone}', Địa chỉ='{address}'")

    if name.lower() == "(trống)":
        name = ""
    if gender.lower() == "(trống)":
        gender = ""
    if dob.lower() == "(trống)":
        dob = ""
    if phone.lower() == "(trống)":
        phone = ""
    if address.lower() == "(trống)":
        address = ""

    # Tách ngày sinh (chỉ nếu không phải "(giữ nguyên)" hoặc "(trống)")
    day, month, year = "", "", ""
    if dob and dob not in ["(giữ nguyên)", "(trống)"]:
        try:
            day, month, year = dob.split("/")
            print(f"Debug - Tách ngày sinh: Ngày='{day}', Tháng='{month}', Năm='{year}'")
        except:
            print("Định dạng ngày sinh không hợp lệ, bỏ qua.")

    return name, gender, (day, month, year), phone, address


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
        name, gender, dob, phone, address = process_cell_value(cell_value)
        test_cases.append((name, gender, dob, phone, address, row))
        row += 1

    # Thực hiện các test thay đổi thông tin
    for index, (name, gender, dob, phone, address, row) in enumerate(test_cases, start=1):
        day, month, year = dob
        print(
            f"\nChạy Test Case {index}: Tên='{name}', Giới tính='{gender}', Ngày sinh='{day}/{month}/{year}', Điện thoại='{phone}', Địa chỉ='{address}'")

        # Truy cập trang thông tin tài khoản
        driver.get(profile_url)
        time.sleep(0.5)

        # Nhập tên hiển thị
        if name != "(giữ nguyên)":
            try:
                name_input = driver.find_element(By.CSS_SELECTOR, "input[data-ng-model='userInfo.displayName']")
                name_input.clear()
                time.sleep(0.5)
                if name:
                    name_input.send_keys(name)
                time.sleep(0.5)
            except:
                print(f"Test Case {index}: Không tìm thấy ô Tên hiển thị.")
                sheet[f"E{row}"].value = "Lỗi: Không tìm thấy ô Tên hiển thị"

        # Chọn giới tính
        if gender != "(giữ nguyên)":
            try:
                if gender == "Nam":
                    gender_input = driver.find_element(By.ID, "nam")
                elif gender == "Nữ":
                    gender_input = driver.find_element(By.ID, "nu")
                elif gender == "Khác":
                    gender_input = driver.find_element(By.ID, "khac")
                else:
                    print(f"Test Case {index}: Giá trị giới tính '{gender}' không hợp lệ.")
                    sheet[f"E{row}"].value = f"Lỗi: Giá trị giới tính '{gender}' không hợp lệ"
                    continue
                gender_input.click()
                time.sleep(0.5)
            except:
                print(f"Test Case {index}: Không tìm thấy radio Giới tính.")
                sheet[f"E{row}"].value = "Lỗi: Không tìm thấy radio Giới tính"

        # Chọn ngày sinh
        if day and day != "(giữ nguyên)":
            try:
                print("Debug - Bắt đầu tìm dropdown ngày sinh...")
                day_select = driver.find_element(By.CSS_SELECTOR, "select[data-ng-model='userInfo.dayOfBirth']")
                day_select.find_element(By.CSS_SELECTOR, f"option[value='{day}']").click()
                time.sleep(0.5)

                month_select = driver.find_element(By.CSS_SELECTOR, "select[data-ng-model='userInfo.monthOfBirth']")
                month_select.find_element(By.CSS_SELECTOR, f"option[value='{month}']").click()
                time.sleep(0.5)

                year_select = driver.find_element(By.CSS_SELECTOR, "select[data-ng-model='userInfo.yearOfBirth']")
                year_select.find_element(By.CSS_SELECTOR, f"option[value='{year}']").click()
                time.sleep(0.5)
                print("Debug - Chọn ngày sinh thành công.")
            except:
                print(f"Test Case {index}: Không tìm thấy dropdown Ngày sinh.")
                sheet[f"E{row}"].value = "Lỗi: Không tìm thấy dropdown Ngày sinh"

        # Nhập số điện thoại
        if phone != "(giữ nguyên)":
            try:
                phone_input = driver.find_element(By.CSS_SELECTOR, "input[data-ng-model='userInfo.phone']")
                phone_input.clear()
                time.sleep(0.5)
                if phone:
                    phone_input.send_keys(phone)
                time.sleep(0.5)
            except:
                print(f"Test Case {index}: Không tìm thấy ô Điện thoại.")
                sheet[f"E{row}"].value = "Lỗi: Không tìm thấy ô Điện thoại"

        # Chọn địa chỉ
        if address != "(giữ nguyên)":
            try:
                address_select = driver.find_element(By.CSS_SELECTOR, "select[data-ng-model='userInfo.address']")
                address_select.find_element(By.CSS_SELECTOR, f"option[value='{address}']").click()
                time.sleep(0.5)
            except:
                print(f"Test Case {index}: Không tìm thấy dropdown Địa chỉ.")
                sheet[f"E{row}"].value = "Lỗi: Không tìm thấy dropdown Địa chỉ"

        # Nhấn nút lưu thay đổi
        try:
            save_button = driver.find_element(By.CSS_SELECTOR, "button.btn-save")
            save_button.click()
            time.sleep(0.5)
        except:
            print(f"Test Case {index}: Không tìm thấy nút Lưu thay đổi.")
            sheet[f"E{row}"].value = "Lỗi: Không tìm thấy nút Lưu thay đổi"
            continue

        # Bắt thông báo thành công hoặc thất bại
        result_msg = ""
        try:
            # Kiểm tra thông báo thành công
            success_msg = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "h4.alert-title"))
            )
            result_msg = success_msg.text
            print(f"Test Case {index}: Success message - {result_msg}")
        except:
            try:
                # Kiểm tra toast thất bại
                toast_msg = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "p.toast__msg"))
                )
                result_msg = toast_msg.text
                print(f"Test Case {index}: Toast message - {result_msg}")
            except:
                result_msg = "Không bắt được thông báo"
                print(f"Test Case {index}: {result_msg}")

        # Ghi kết quả vào cột E
        sheet[f"E{row}"].value = result_msg
        print(f"Debug - Ghi kết quả vào E{row}: '{result_msg}'")

        # Tải lại trang cho test case tiếp theo
        driver.get(profile_url)
        time.sleep(0.5)

# Lưu file Excel
workbook.save(excel_path)

# Đóng trình duyệt
driver.quit()
print(f"\nHoàn thành {len(test_cases)} test case. Kết quả đã được ghi vào cột E của file Excel.")