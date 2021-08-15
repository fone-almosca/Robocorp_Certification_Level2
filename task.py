"""Template robot with Python."""

import os
import csv
import time

from Browser import Browser
from Browser.utils.data_types import SelectAttribute
from RPA.Robocloud.Secrets import Secrets
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.Dialogs import Dialogs


#initialize variables
browser = Browser()
lib = Archive()
d = Dialogs()
http = HTTP()
pdf = PDF()
secrets = Secrets()


def download_the_csv(url):
    http.download(url)


def ask_for_url():
    d.add_heading("Please provide the URL where download the orders")
    d.add_text_input(name="url",label="URL")
    result = d.run_dialog()
    return result.url


def open_webapplication(): 
    browser.open_browser("https://robotsparebinindustries.com/#/robot-order")


def process_the_orders():
    # get the filename from the vault 
    orders_filename = secrets.get_secret("orders_parameters")["filename"]
    file_exist = False
    counter = 0
    while not file_exist:
        time.sleep(1)
        if (os.path.isfile(orders_filename)):
            file_exist = True
        else:
            counter += 1
        if counter>=10:
            break
    if file_exist:
        with open('orders.csv', newline='') as csvfile:
             reader = csv.DictReader(csvfile)
             for row in reader:
                fill_the_form(row)
                time.sleep(1)
             #rint("Process completed!")
    else: 
        print("file does not exist")


def remove_message():
    browser.click("text=OK")


def fill_the_form(row):
    preview_appeared = False
    attempt_1 = 1
    while not preview_appeared:
        try:
            # Insert the Head
            browser.select_options_by("xpath=//select[@id='head']",SelectAttribute["value"],str(row["Head"]))

            # Insert Body
            label_index = "id-body-" + str(row["Body"])
            browser.click("xpath=//input[@id='" + label_index + "']")

            # Insert Legs 
            browser.type_text("xpath=//input[@type='number']", str(row["Legs"]))

            # Insert Address
            browser.type_text("xpath=//input[@id='address']", str(row["Address"]))

            # Click on Preview
            browser.click("id=preview")

            # Extract the preview image
            time.sleep(2)
            preview_filename = f"{os.getcwd()}/output/preview_"+ str(row["Order number"]) + ".png"
            browser.take_screenshot(filename=preview_filename,
                                    selector="xpath=//div[@id='robot-preview-image']")
            preview_appeared = True
            # Click on Order
            order_complete = False
            attempts_2 = 1
            while not order_complete:
                browser.click("id=order")
                # Generate PDF 
                order_complete = generate_pdf(row["Order number"], preview_filename)
                if order_complete: 
                    insert_new_order()
                    #print("Order " + str(row["Order number"]) + " completed")
                if attempts_2 == 3:
                    #print("Order " + str(row["Order number"]) + " failed generating order")
                    break
                else:
                    attempts_2 += 1
        except: 
            #print("Order " + str(row["Order number"]) + " error while inserting the parameters")
            continue
        finally: 
            if (preview_appeared == True and order_complete == True) or attempt_1 == 3:
                break
            else:
                preview_appeared = False
                attempt_1 += 1


def generate_pdf(order_number, preview_filename):
    try:
        pdf_filename = f"{os.getcwd()}/output/receipt_"+ order_number+".pdf"
        receipt_html = browser.get_property(
            selector="xpath=//div[@id='receipt']", property="outerHTML")
        pdf.html_to_pdf(receipt_html, pdf_filename)
        # add image
        pdf.add_watermark_image_to_pdf(image_path=preview_filename,
                                       source_path=pdf_filename,
                                       output_path=pdf_filename)
        order_complete = True
    except: 
        order_complete = False
    return order_complete


def insert_new_order():
    browser.click("id=order-another")
    remove_message()


def create_zip_file():
    lib.archive_folder_with_zip(folder=f"{os.getcwd()}/output/", 
                            archive_name=f"{os.getcwd()}/output/pdf_receipts.zip",
                            include="*.pdf")


def close_browser(): 
    browser.playwright.close()


if __name__ == "__main__":  
    
    try:
        url = ask_for_url()
        download_the_csv(url)        
        open_webapplication()
        remove_message()        
        process_the_orders()
        create_zip_file()
    finally:
        close_browser()
