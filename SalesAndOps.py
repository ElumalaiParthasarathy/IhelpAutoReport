# -*- coding: utf-8 -*-
"""
Created on Mon Sep  2 21:12:48 2024

@author: elumalai.p2@piramal.com
"""


import configparser
from selenium.webdriver.chrome.service import Service
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import logging
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoSuchElementException, TimeoutException


# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class IhelpReport:
    new_count_xpath = "(//*[@x='38.25'])[2]"
    inprogress_count_xpath = "(//*[@x='191.25'])[2]"
    assigned_count_xpath = "(//*[@x='114.75'])[2]"
    pending_count_xpath = "(//*[@x='267.75'])[2]"
    pending_button_xpath = "(//*[text()='Pending'])[1]"
    to_load_xapth='//*[@id="BodyContentPlaceHolder_lblcsatResolved"]'
  
   

    def __init__(self):
         #chrome_options = Options()
        config = configparser.ConfigParser()
        config.read('config/config.ini')
        self.user = config.get('EMAIL', 'USER')
        self.app_password = config.get('EMAIL', 'APP_PASSWORD')
        self.to_mails = [email.strip() for email in config.get('EMAIL', 'to_mails').split(',')]
        service = Service()
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")

    # Other options to improve performance in headless mode
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

         # Initialize the WebDriver with the service and options
        self.driver = webdriver.Chrome(service=service, options=options)

    def wait_for_page_to_load(self, timeout=300):
        WebDriverWait(self.driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
    def wait_for_element(self, xpath, timeout=30):
        WebDriverWait(self.driver, timeout).until(
        EC.visibility_of_element_located((By.XPATH, xpath))
    )


    def send_email(self, subject, body):
        msg = MIMEMultipart()
        msg['From'] = self.user
        current_time = datetime.now().strftime("%Y-%m-%d %H-00")
        msg['Subject'] = f"{subject} - {current_time}"
        
        # Combine all recipients into a single string separated by commas
        msg['To'] = ', '.join(self.to_mails)
        
        # If you have BCC recipients, add them here
        
    
        # Attach the body with the msg instance
        msg.attach(MIMEText(body, 'plain'))
    
        # Send the email
        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(self.user, self.app_password)
            text = msg.as_string()
            
            # Combine To and BCC recipients into one list for sending
            all_recipients = self.to_mails
            server.sendmail(self.user, all_recipients, text)
            server.quit()
            logger.info("Emails sent successfully!")
        except Exception as e:
            logger.error(f"Failed to send email: {e}")



    def login_ihelp(self, website):
        try:
            logger.info("Browser Initiated")
            self.driver.get(website)
            self.driver.maximize_window()

            self.driver.find_element(By.ID, "txtLogin").send_keys("elumalai.p2")
            self.driver.find_element(By.ID, "txtPassword").send_keys("Seven@27")
            self.driver.find_element(By.ID, "butSubmit").click()
            logger.info("Login Successful")

            self.wait_for_page_to_load(10)

            iframes = self.driver.find_elements(By.ID, "SPopUp-frame")
            if iframes:
                self.driver.switch_to.frame(iframes[0])
                WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "ContentPanel_btnContinue"))
                ).click()
                self.driver.switch_to.default_content()

            # Get reports for both workgroups
            sales_central_report = self.get_workgroup_report("Sales Central")
            ops_central_report = self.get_workgroup_report("OPS Central")
            
            body = f"{sales_central_report}\n\n" \
                   f"{ops_central_report}"
            
            logger.info(body)
            self.send_email("Ihelp Sales and Ops Reports", body)

        finally:
            self.logout()
            self.driver.quit()

    def get_workgroup_report(self, workgroup):
        self.driver.find_element(By.XPATH, "//a[@id='filter']").click()
        self.driver.find_element(By.XPATH, "(//span[@class='select2-selection__arrow'])[2]").click()

        if workgroup.lower() == "sales central":
            option = self.driver.find_element(By.XPATH, "//ul[@class='select2-results__options']//li[text()='Sales Central']")
        else:
            option = self.driver.find_element(By.XPATH, "//ul[@class='select2-results__options']//li[text()='OPS Central']")

        option.click()
        self.driver.find_element(By.XPATH, "(//input[@value='Submit'])[1]").click()
        

        counts = self.take_count()
      
        pending_report = self.get_pending_report(workgroup)

        body = f"Auto Generated Incident Report for {workgroup}:\n\n" \
               f"{counts}\n" \
               f"{pending_report}"

        return body

    def take_count(self):
        self.wait_for_element(self.to_load_xapth,5)
        pending_count = new_count = assigned_count = inprogress_count = 0

        try:
            self.wait_for_element(self.new_count_xpath,5)
            new_counts = self.driver.find_element(By.XPATH, self.new_count_xpath)
            new_count = int(new_counts.text) if new_counts.text.isdigit() else 0
        except (NoSuchElementException, TimeoutException):
            new_count = 0

        try:
            self.wait_for_element(self.inprogress_count_xpath,5)
            inprogress = self.driver.find_element(By.XPATH, self.inprogress_count_xpath)
            inprogress_count = int(inprogress.text.strip()) if inprogress.text.strip().isdigit() else 0
        except (NoSuchElementException, TimeoutException):
            inprogress_count = 0

        try:
            self.wait_for_element(self.pending_count_xpath,5)
            pending = self.driver.find_element(By.XPATH, self.pending_count_xpath)
            pending_count = int(pending.text.strip()) if pending.text.strip().isdigit() else 0
        except (NoSuchElementException, TimeoutException):
            pending_count = 0

        try:
            self.wait_for_element(self.assigned_count_xpath,5)
            assigned = self.driver.find_element(By.XPATH, self.assigned_count_xpath)
            assigned_count = int(assigned.text.strip()) if assigned.text.strip().isdigit() else 0
        except (NoSuchElementException, TimeoutException):
            assigned_count = 0

        body = f"New Count: {new_count}\n" \
               f"Assigned Count: {assigned_count}\n" \
               f"Inprogress Count: {inprogress_count}\n" \
               f"Pending Count: {pending_count}\n\n"

        return body

    def get_pending_report(self,workgroup):
        self.driver.find_element(By.XPATH, self.pending_button_xpath).click()
    
        original_tab = self.driver.current_window_handle
        all_tabs = self.driver.window_handles
        for tab in all_tabs:
            if tab != original_tab:
                self.driver.switch_to.window(tab)
                break
    
        self.wait_for_page_to_load(60)
    
        report = f"Pending Incidents in {workgroup}:\n"
    
        try:
            # Check if the records container element is present
            self.driver.find_element(By.XPATH, "//*[@id='select2-BodyContentPlaceHolder_ddlRecords-container']").click()
            self.driver.find_element(By.XPATH, "//ul[@class='select2-results__options']//li[text()='100']").click()
            self.wait_for_page_to_load(120)
    
            # Extract the report data from the table
            table = self.driver.find_element(By.XPATH, '//table[@id="BodyContentPlaceHolder_gvMyTickets"]')
            rows = table.find_elements(By.TAG_NAME, "tr")
    
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) > 2:
                    pending_reason = cells[11].text.strip()  # Get and trim pending reason
                    incident_id = cells[1].text.strip()      # Get and trim incident ID
                    report += f"{incident_id:<15}|{pending_reason:<20}\n"
    
        except NoSuchElementException:
            report += "No pending incidents found.\n"
    
        # Close any tab that's not the original one
        for tab in all_tabs:
            if tab != original_tab:
                self.driver.switch_to.window(tab)
                self.driver.close()
    
        # Switch back to the original tab
        self.driver.switch_to.window(original_tab)
    
        return report

    def logout(self):
        self.driver.find_element(By.XPATH,'//div[@class="profile dropdown"]').click()
        self.driver.find_element(By.XPATH,'//a[@id="hrefLogout"]').click()
if __name__ == "__main__":
    ihelp_report = IhelpReport()
    ihelp_report.login_ihelp("https://mihelp.piramal.com/formLogin/#!")

