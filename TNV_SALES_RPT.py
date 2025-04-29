from playwright.sync_api import sync_playwright
from datetime import datetime, timedelta
import time
import lackey as lk
import numpy as np
import subprocess
import schedule

# Run JavaScript file
subprocess.run(["node", "script.js"], shell=True)


def run():
    with sync_playwright() as p:
        # Launch browser
        browser = p.chromium.launch(headless=False)  # Change to True to run headless
        context = browser.new_context()
        page = context.new_page()

        # Open the Crystal Report URL
        page.goto("http://192.168.50.32/branch/")
        page.set_viewport_size({"width": 1366, "height": 768})

        # Login
        page.fill("#txtUserName", "1110")
        page.fill("#txtPassword", "1110")
        page.click("xpath=/html/body/div[2]/div[1]/form/div[3]/div[4]/div/button")

        # Wait for page load
        page.wait_for_timeout(5000)

        # Navigate to Reports Menu
        page.click("xpath=/html/body/form/div[3]/div[1]/div/nav/div/ul/li[3]/a")
        page.wait_for_timeout(2000)

        # Click on Sales Report
        page.click("xpath=/html/body/form/div[3]/div[2]/div/div/div/div[2]/div[2]/div/div/div[3]/div/div/div[2]/div/div[1]/a")
        page.wait_for_timeout(3000)

        # # Select "From Date"
        page.wait_for_selector("#fromDate")
        page.click("#fromDate")
        time.sleep(1)  

        selected_date = page.input_value("#fromDate")  # Expected format: "dd/mm/yyyy"
        print(f"Selected Date: {selected_date}")
        
        date_obj = datetime.strptime(selected_date, "%d/%m/%Y")
        prev_date = date_obj - timedelta(days=1)
        prev_day = str(prev_date.day)
        prev_month_year = prev_date.strftime("%B %Y")


        while True:
                current_month_year = page.locator(".datepicker-switch").nth(0).text_content().strip()
                print(f"Current Month and Year: {current_month_year}")
                if current_month_year == prev_month_year:
                    break
                page.click(".prev")


        page.locator(f"//td[@class='day' and not(contains(@class, 'disabled')) and text()='{prev_day}']").click()
        page.click(".input-group-addon")  

        # Select "To Date"
        page.wait_for_selector("#toDate")
        page.click("#toDate")
        time.sleep(1) 

        # Get the currently selected date
        selected_date = page.input_value("#toDate")  # Expected format: "dd/mm/yyyy"
        print(f"Selected Date: {selected_date}")

        # Parse the selected date
        date_obj = datetime.strptime(selected_date, "%d/%m/%Y")

        # Calculate the previous day
        prev_date = date_obj - timedelta(days=1)
        prev_day = str(prev_date.day)
        prev_month_year = prev_date.strftime("%B %Y")

        # Navigate to the correct month
        while True:
            current_month_year = page.locator(".datepicker-switch").nth(0).text_content().strip()
            print(f"Current Month and Year: {current_month_year}")
            if current_month_year == prev_month_year:
                break
            page.click(".prev")  # Click the "previous" button to navigate to the previous month

        # Select the previous day (ensuring it’s enabled)
        page.locator(f"//td[@class='day' and not(contains(@class, 'disabled')) and text()='{prev_day}']").click()
        page.click(".input-group-addon")  


        # Select Group By
        page.click("#s2id_cmbGroupBy")
        page.fill("#s2id_autogen1_search", "Section Wise")
        page.wait_for_selector("#select2-results-1 > li")
        page.click("#select2-results-1 > li")
        time.sleep(1)

        # Select PR
        page.click("#s2id_cmbPR")
        page.fill("#s2id_autogen2_search", "PR")
        page.wait_for_selector("#select2-results-2 > li")
        page.click("#select2-results-2 > li")
        time.sleep(1)

        # Click Show Report Button
        page.click("#btnList")
        page.wait_for_timeout(3000)

        frame = page.frame_locator("#frmeBar")  # Change to your actual iframe ID
     
        export_button = frame.locator("#IconImg_CrystalReportViewer1_toptoolbar_export")  # Adjust ID
        export_button.click()
        print("Export button clicked.")
        page.wait_for_timeout(2000)
        
        
        rpt_page=page.frame_locator("#bobjid_1743675424641_page")
        print("frame located")
        # rpt_page.wait_for_timeout(2000)

        fil_form=page.frame_locator(".naviBarFrame naviFrame")   
        print("file format located")
        # fil_form.wait_for_timeout(2000)

        dropdown = frame.locator('xpath=/html/body/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/table/tbody/tr/td/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div')  # Correcting this
        dropdown.click()
        time.sleep(1)
        
        # Select PDF format
        pdf_option = frame.locator('xpath=/html/body/table[2]/tbody/tr/td/table/tbody/tr[2]/td[2]')
        pdf_option.click()
        time.sleep(1)

        # epxo_button=frame.locator('xpath=/html/body/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]')
        # epxo_button.click()
        # time.sleep(5)
        with page.expect_download() as download_info:
            epxo_button=frame.locator('xpath=/html/body/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]')
            epxo_button.click()
            time.sleep(5)

        # Get the download object
        download = download_info.value

        # Define custom path
        # save_path = r"H:\joedev\Myfolder\stock_rpt2.pdf"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # Format: YYYYMMDD_HHMMSS
        save_path = fr"H:\joedev\Myfolder\stock_rpt_{timestamp}.pdf"
        # Save the file to the specific folder
        download.save_as(save_path)

        print(f"File saved to: {save_path}")
        # Close browser  
        context.close()
        browser.close()           

schedule.every(24).hours.do(run)
if __name__ == "__main__":
    run()
while True:
        schedule.run_pending()  # Runs any pending scheduled tasks
        time.sleep(1)  