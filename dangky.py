from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


def test_register_account():
    # Khởi tạo driver
    driver = webdriver.Chrome()
    driver.maximize_window()

    try:
        # Truy cập trang
        driver.get("https://my.thanhnien.vn/page/login.html?redirect_url=https://thanhnien.vn/")
        time.sleep(0.5)  # Delay 2s sau khi load trang

        # Chờ và chọn tab Đăng ký
        wait = WebDriverWait(driver, 10)
        register_tab = wait.until(
            EC.element_to_be_clickable((By.XPATH, '//li[@class="tabs-login-item"]/a[@href="#menu_2"]')))
        register_tab.click()
        time.sleep(0.5)  # Delay 2s sau khi click tab

        # Chờ form đăng ký hiển thị
        form = wait.until(EC.visibility_of_element_located((By.ID, "menu_2")))
        assert form.is_displayed(), "Form đăng ký không hiển thị"
        time.sleep(0.5)  # Delay 2s sau khi kiểm tra form

        # Điền thông tin
        email_field = driver.find_element(By.CSS_SELECTOR, "input[data-ng-model='registration.email']")
        email_field.send_keys("testuser" + str(int(time.time())) + "@example.com")  # Email duy nhất
        time.sleep(0.5)  # Delay 2s sau khi điền email

        name_field = driver.find_element(By.CSS_SELECTOR, "input[data-ng-model='registration.displayName']")
        name_field.send_keys("Test User")
        time.sleep(0.5)  # Delay 2s sau khi điền tên

        password_field = driver.find_element(By.CSS_SELECTOR, "input[data-ng-model='registration.password']")
        password_field.send_keys("Test@123")
        time.sleep(0.5)  # Delay 2s sau khi điền mật khẩu

        confirm_password_field = driver.find_element(By.CSS_SELECTOR,
                                                     "input[data-ng-model='registration.confirmPassword']")
        confirm_password_field.send_keys("Test@123")
        time.sleep(0.5)  # Delay 2s sau khi điền xác nhận mật khẩu

        # Nhấn nút Đăng ký
        register_button = driver.find_element(By.CSS_SELECTOR, "a[data-ng-click='signUp()']")
        register_button.click()
        time.sleep(0.5)  # Delay 2s sau khi nhấn nút

        # Kiểm tra chuyển hướng đến https://thanhnien.vn/
        wait.until(EC.url_to_be("https://thanhnien.vn/"))
        assert driver.current_url == "https://thanhnien.vn/", "Chuyển hướng không đúng"
        time.sleep(0.5)  # Delay 2s sau khi kiểm tra URL

        print("Test case đăng ký tài khoản: PASS")

    except Exception as e:
        print(f"Test case thất bại: {str(e)}")

    finally:
        time.sleep(0.5)  # Delay 2s trước khi đóng
        driver.quit()


# Chạy test
test_register_account()