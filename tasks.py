from robocorp.tasks import task
from RPA.HTTP import HTTP
from robocorp import browser
from RPA.PDF import PDF
import pandas as pd
import os
import zipfile

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(slowmo=100)
    open_robot_order_website()
    web_response = get_orders()
    close_annoying_modal()
    
    orders_csv = pd.read_csv('output/orders.csv')
    for index, order in orders_csv.iterrows():
        fill_the_form(order)
    archive_receipts()

def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/")

def get_orders():
    http = HTTP()
    web_response = http.download("https://robotsparebinindustries.com/orders.csv",'output/orders.csv',overwrite=True)
    return web_response

def close_annoying_modal():
    #Click the button of order your robot
    page = browser.page()
    page.click('//*[@class="nav-link"]')
    #Click the button to close the pop up
    page.click('//*[@class="btn btn-danger"]')

def fill_the_form(order):
    page = browser.page()
    #We select the head of the robot
    page.select_option("#head",str(order["Head"]))

    # #We select the body of the robot
    body_number = str(order["Body"])
    page.click(f"#id-body-{body_number}")

    # #We select the legs of the robot
    page.fill("//*[@class='form-control']",str(order["Legs"]))

    # #We select the address
    page.fill("#address",order["Address"])

    # #We select the preview
    page.click("#preview")

    # #We create the order
    page.click("#order")

    # #We click again if there has been an error
    while page.is_visible("#order-another") != True:
        page.click("#order")
    
    image_path = screenshot_robot(order["Order number"])

    # #Next robot
    page.click("#order-another")

    # #Get rid of the pop up
    page.click('//*[@class="btn btn-danger"]')

    pdf_path = store_receipt_as_pdf(order["Order number"])

    embed_screenshot_to_receipt(image_path, pdf_path)


def store_receipt_as_pdf(order_number):
    pdf = PDF()
    pdf_path = ""
    html_content = f"<html><body><h1>Order number: {order_number}</h1></body></html>"
    pdf_path = f"output/Order{order_number}.pdf"
    pdf.html_to_pdf(html_content, pdf_path)
    return pdf_path

def screenshot_robot(order_number):
    page = browser.page()
    image_path = f"output/Order{order_number}.png"
    page.screenshot(path=image_path)
    return image_path

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    list_of_files = [
        pdf_file,  
        f'{screenshot}:align=center'  
    ]
    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document=pdf_file,
    )

def archive_receipts():
    output_path = './output/'
    files = os.listdir(output_path)
    pdfs = []
    for file in files:
        if file.endswith('.pdf'):
            pdf_path =  output_path + file
            pdfs.append(pdf_path)

    with zipfile.ZipFile('output/receipts.zip', 'w') as myzip:
        for pdf in pdfs:
            myzip.write(pdf)
