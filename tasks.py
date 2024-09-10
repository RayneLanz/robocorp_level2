from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import time
@task
def minimal_task():
    '''Main'''
    browser.configure(
        slowmo=100
    )

    openLink()
    closeModal()
    downloadExcel()
    getOrders()

    ArchiveFiles()
    wait()



def openLink():
    '''Ordering Robots from spareBin'''
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    
def closeModal():
    '''Closes the Modal to the Website when first opened'''
    page = browser.page()
    page.click("button:text('I guess so...')")


def downloadExcel():
    '''Download the Excel'''
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def getOrders():
    '''Ordering using the Data from Excel'''
    temp = []
    table = Tables()
    orders = table.read_table_from_csv(
        "orders.csv",columns=["Order number","Head","Body","Legs","Address"]
    )
    for x in orders:
        
        #print(x["Head"])
        fillupForm(x["Head"],x["Body"],x["Legs"],x["Address"])
        
    
    #print(temp)



    #print(orders)

def fillupForm(Head,Body,Legs,Address):
    '''Start to fill up form'''
    page = browser.page()
    page.select_option('#head',Head)
    page.click(f"#id-body-{Body}")
    page.fill('input[placeholder="Enter the part number for the legs"]',Legs)
    page.fill('#address',Address)
    SubmitOrder()
    screenshotRobot(Address)
    getReceiptsAsPdf(Address)

    page.click("#order-another")
    time.sleep(1)
    closeModal()


def SubmitOrder(max_retries=5):
    '''Check if there are an error when submiting'''
    page = browser.page()
    flag = True
    count = 0
    while flag:
        page.click("#order")
        time.sleep(2)
        error_message = page.query_selector(".alert-danger")
        if count >=6:
            flag=False

        if error_message and error_message.is_visible():
            count+=1
        else:
            flag =False
            return True


def getReceiptsAsPdf(Address):
    '''This task to get the receipts and turn each one as pdfs'''
    #receipt
    page = browser.page()
    receipts = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(receipts,f"output/receipts/{Address}.pdf")

    pdf.add_files_to_pdf(
        
        files=[f"output/receipts/img/{Address}.png"],
        target_document=f"output/receipts/{Address}.pdf",
        append=True
    )

    
def screenshotRobot(Address):
    '''This will screenshot the ordered robot'''
    page = browser.page()
    robotPrev = page.locator("#robot-preview")
    robotPrev.screenshot(path=f"output/receipts/img/{Address}.png")

def ArchiveFiles():
    '''Archive the receipts'''
    archive = Archive()
    archive.archive_folder_with_zip(
        folder="output/receipts",
        archive_name="output/receipts.zip",
        include="*.pdf"
    )

def wait():
    
    time.sleep(2)
