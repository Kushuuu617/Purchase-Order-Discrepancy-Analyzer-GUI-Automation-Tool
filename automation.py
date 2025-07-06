from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, WebDriverException

from excel_writer import create_excel_sheet
from utils import update_status
from tkinter import messagebox
import traceback

def start_automation(username, password, status_text):
    driver = None
    serial_no = 1

    try:
        options = Options()
        options.add_argument("--start-maximized")
        driver = webdriver.Chrome(options=options)
        update_status(status_text, "Browser launched successfully.")

        driver.get("https://induserp.industowers.com/OA_HTML/AppsLocalLogin.jsp")
        update_status(status_text, "Navigating to login page...")

        driver.find_element(By.ID, 'usernameField').send_keys(username)
        driver.find_element(By.ID, 'passwordField').send_keys(password)

        driver.find_element(By.CLASS_NAME, "OraButton").click()

        WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.XPATH, "//div[text()='India Local iSupplier']"))
        ).click()
        update_status(status_text, "Logged in successfully.")

        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//div[text()='Home Page']"))
        ).click()

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ILS_POS_HOME"))
        ).click()

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "PosHpOrdersTable:PosHpoPoNum:0"))
        ).click()

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "POS_PURCHASE_ORDERS"))
        ).click()

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "SrchBtn"))
        ).click()

        messagebox.showinfo("Manual Step", "Please apply your filters manually, then click 'Go' on the page. Once results load, press OK to continue.")

        try:
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, 'ResultRN.PosVpoPoList:Content'))
            )
            update_status(status_text, "Filtered results loaded. Starting data extraction...")
        except TimeoutException:
            messagebox.showerror("Timeout", "Result table did not load after you clicked 'Go'. Please retry.")
            return

        wb, ws = create_excel_sheet()

        while True:
            try:
                table = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.ID, 'ResultRN.PosVpoPoList:Content'))
                )
                rows = table.find_elements(By.TAG_NAME, 'tr')

                for i in range(len(rows)):
                    try:
                        doc_type = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, f"//span[contains(@id, 'PosDocumentType:{i}')]"))
                        ).text.strip()
                        if doc_type == "Global Blanket Agreement":
                            continue

                        po_link = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, f"//a[contains(@id, 'PosPoNumber:{i}')]"))
                        )
                        po_number = po_link.text

                        driver.execute_script("window.open(arguments[0], '_blank');", po_link.get_attribute('href'))
                        driver.switch_to.window(driver.window_handles[1])

                        try:
                            total_amount = float(WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, "TotalAmt"))
                            ).text.replace(',', ''))

                            billed_amount = float(WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, "AmtBilled"))
                            ).text.replace(',', ''))

                            if billed_amount < total_amount:
                                deviation = total_amount - billed_amount
                                update_status(status_text, f"Discrepancy found: {po_number}")
                                try:
                                    po_description = WebDriverWait(driver, 5).until(
                                        EC.presence_of_element_located((By.ID, "PosPoDescription"))
                                    ).text.strip()
                                except Exception:
                                    po_description = ""

                                site_id_elements = driver.find_elements(By.XPATH, "//span[contains(@id, 'PosOrderLines:SiteID')]")
                                seen_site_ids = set()

                                for site_id_el in site_id_elements:
                                    site_index = site_id_el.get_attribute("id").split(":")[-1]
                                    site_id = site_id_el.text.strip()
                                    if site_id in seen_site_ids or not site_id:
                                        continue
                                    seen_site_ids.add(site_id)

                                    try:
                                        site_name_el = driver.find_element(By.XPATH, f"//span[@id='PosOrderLines:SiteNameID:{site_index}']")
                                        site_name = site_name_el.text.strip()
                                    except NoSuchElementException:
                                        site_name = ""

                                    ws.append([serial_no, po_number, total_amount, billed_amount, deviation, site_id, site_name, po_description])
                                    serial_no += 1

                        except Exception as e:
                            update_status(status_text, f"Error reading PO details for {po_number}: {str(e)}")
                        finally:
                            driver.close()
                            driver.switch_to.window(driver.window_handles[0])
                            WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.ID, 'ResultRN.PosVpoPoList:Content'))
                            )

                    except Exception as e:
                        update_status(status_text, f"Error processing PO index {i}: {str(e)}")

                try:
                    next_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Next')]"))
                    )
                    next_button.click()
                    WebDriverWait(driver, 10).until(EC.staleness_of(table))
                except (TimeoutException, NoSuchElementException, ElementClickInterceptedException):
                    update_status(status_text, "✅ All pages processed.")
                    break

            except Exception as e:
                update_status(status_text, f"Error loading table: {str(e)}")
                break

        wb.save("POs_with_billed_less_than_total.xlsx")
        update_status(status_text, "✔ Excel file saved successfully.")

    except WebDriverException as we:
        update_status(status_text, f"Browser/driver issue: {we}")
    except Exception as e:
        update_status(status_text, "Fatal error occurred.\n" + traceback.format_exc())
    finally:
        if driver:
            try:
                driver.quit()
                update_status(status_text, "Browser closed.")
            except:
                update_status(status_text, "Failed to close browser properly.")