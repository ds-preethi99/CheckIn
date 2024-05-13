import send_check_in_alert
import send_issue_alert
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from time import sleep
from datetime import date


def export_attendance_excel(logger, user_id, password, webhook_url, lead_webhook_url, discord_user_id):
    try:     # Create a new driver instance and add the webdriver
        today = date.today()
        formatted_date = today.strftime("%d/%m/%Y")
        s = Service('chromedriver.exe')
        driver = webdriver.Chrome(service=s)
        driver.get('https://digitalseo.hrapp.co/auth')           # Open HRApp
        driver.maximize_window()              # Maximizing the chrome window
        sleep(10)
        driver.find_element(By.CSS_SELECTOR, '#signinWithEmailBtnText').click()         # click on the sign with email button
        driver.find_element(By.CSS_SELECTOR, '#formSigninEmailId').send_keys(user_id)   # Adjust based on actual element ID or another selector
        sleep(5)
        driver.find_element(By.CSS_SELECTOR, '#email-verification-button').click()      # click on the Verify button
        sleep(5)
        driver.find_element(By.ID, 'formSigninPassword').send_keys(password)  # Identify and interact with the password field
        sleep(5)
        driver.find_element(By.ID, 'email-password-submit-button').click()          # Identify and click the Submit button
        sleep(15)
        driver.get('https://digitalseo.hrapp.co/employees/attendance')
        sleep(10)
        driver.find_element(By.ID, 'filterAttendance').click()            # Click the filter button
        sleep(5)
        driver.find_element(By.ID, 'filterAttendanceDateBegin').click()           # inside the employee button we have a option called leave click the click button
        sleep(5)
        # Identify the datepicker input field (replace the following XPath with the actual locator for your datepicker)
        datepicker_input = driver.find_element(By.XPATH, '//*[@id="filterAttendanceDateBegin"]')
        # Clear the existing value (optional, if needed)
        datepicker_input.clear()
        # Enter the formatted date using send_keys()
        datepicker_input.send_keys(formatted_date)
        sleep(5)
        driver.find_element(By.ID, 'filterApplyAttendance').click()       # click apply button in filter
        sleep(5)
        driver.find_element(By.ID, 'exportAttendance').click()            # click export button
        sleep(10)
        logger.info('Attendance file downloaded')
        # download Today's Leave List
        driver.get('https://digitalseo.hrapp.co/employees/leaves')
        sleep(10)
        driver.find_element(By.ID, 'filterLeave').click()  # Click the filter button
        sleep(5)
        # Identify the datepicker input field (replace the following XPath with the actual locator for your datepicker)
        datepicker_input = driver.find_element(By.XPATH, '//*[@id="filterLeaveDateBegin"]')
        # Clear the existing value (optional, if needed)
        datepicker_input.clear()
        # Enter the formatted date using send_keys()
        datepicker_input.send_keys(formatted_date)
        sleep(5)
        driver.find_element(By.ID, 'filterApplyLeave').click()  # click apply button in filter
        sleep(5)
        driver.find_element(By.ID, 'exportLeave').click()  # click export button
        sleep(10)
        driver.quit()                      # quit the program
        logger.info('Leave Register file downloaded')
        send_check_in_alert.move_attendance_file(logger,webhook_url)
        send_check_in_alert.move_leave_register_file(logger,webhook_url)
        send_check_in_alert.get_not_check_in_list(logger,lead_webhook_url,discord_user_id)
    except Exception as e:
        response = send_issue_alert.send_discord_message_exception(webhook_url, f'Exception occurs in Check-in alert - check the log')
        # Check whether the response send or not
        if response.status_code == 204:
            print("Message sent successfully!")
        else:
            print(f"Failed to send message. Status code: {response.status_code} {response.text}")
