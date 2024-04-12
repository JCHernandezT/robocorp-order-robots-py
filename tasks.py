from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.FileSystem import FileSystem

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    orders = get_orders()
    open_robot_order_website()
    for order in orders:
        close_annoying_modal()
        fill_the_form(order)
        preview_robot()
        submit_robot()
        store_receipt_as_pdf(order['Order number'])
    archive_receipts()

    message = "finished"

def open_robot_order_website():
    """
    Opens the orders website.
    """
    browser.configure(
        slowmo=100,
    )
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """
    Downloads the orders CSV and return a table of orders.
    """
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"]
    )
    return orders

def close_annoying_modal():
    """
    Closes the annoying modal if visible.
    """
    page = browser.page()
    button_selector = "button.btn-dark"
    if page.is_visible(button_selector):
        page.click(button_selector)

def fill_the_form(order):
    """
    Fills and order form.
    """
    page = browser.page()
    page.select_option('//*[@id="head"]', str(order["Head"]))
    page.click(f'//*[@id="id-body-{order["Body"]}"]')
    page.fill('input[placeholder="Enter the part number for the legs"]', order["Legs"])
    page.fill('//*[@id="address"]', order["Address"])

def preview_robot():
    """
    Previews a robot.
    """
    page = browser.page()
    page.click('//*[@id="preview"]')
    locator = page.locator('//*[@id="robot-preview-image"]')
    locator.wait_for()

def submit_robot():
    """
    Submits a robot order.
    """
    page = browser.page()
    button_selector = '//*[@id="order"]'
    page.click(button_selector)
    
    retry_count = 0
    success = False
    MAX_RETRIES = 3 

    error_msg_selector = '.alert.alert-danger[role="alert"]'
    while retry_count < MAX_RETRIES and not success:
        if not page.is_visible(error_msg_selector):
            success = True
            # page.click('//*[@id="order-another"]')
            break
        page.click(button_selector)

def store_receipt_as_pdf(order_number):
    """
    Stores a receipt as a PDF file.
    """
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    page.locator("#robot-preview-image").screenshot(path="output/preview.png")
    pdf = PDF()
    pdf.html_to_pdf(receipt_html, f"output/receipts/{order_number}.pdf")
    list_of_files = [
        f"output/receipts/{order_number}.pdf",
        "output/preview.png",
    ]
    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document=f"output/receipts/{order_number}.pdf"
    )

    page.click('//*[@id="order-another"]')

def archive_receipts():
    """
    Creates a ZIP file with all the PDF orders.
    """
    archive = Archive()
    archive.archive_folder_with_zip(folder="output/receipts", archive_name='receipts.zip', recursive=True)
    
    file_system = FileSystem()
    file_system.remove_directory(path="output/receipts", recursive=True)
