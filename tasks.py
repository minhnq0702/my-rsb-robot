import logging
import pandas as pd

from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel import Files
from RPA.PDF import PDF

_logger = logging.getLogger(__name__)

SALE_FILE_NAME = "salesdata.xlsx"

@task
def robot_spare_bin_python():
    """Insert the sales data for the week and then export it as PDF"""
    browser.configure(
        slowmo=100,
    )
    open_the_intranet()
    log_in()
    download_excel_file()
    fill_and_submit_sale_form()
    collect_result()
    export_pdf()

    # logout after successful robot's task
    log_out()

def open_the_intranet():
    """Open browser and navigator to the URL"""
    browser.goto("https://robotsparebinindustries.com")
    pass

def log_in():
    """
    Login intranet
    Fill in the username and password and click on 'Login in' button"
    """
    page = browser.page()
    page.fill("id=username", "maria")
    page.fill("id=password", "thoushallnotpass")
    page.click("button:text('Log in')")

def fill_and_submit_sale_form():
    """Fill in the sales data and submit the form"""
    def _fill_form(sales_rep: dict):
        page.fill("#firstname", sales_rep["First Name"])
        page.fill("#lastname", sales_rep["Last Name"])
        page.select_option("#salestarget", str(sales_rep["Sales Target"]))
        page.fill("#salesresult", str(sales_rep["Sales"]))
        page.click("text=Submit")

    page = browser.page()
    excel = Files.Files()
    excel.open_workbook(SALE_FILE_NAME)
    ws = excel.read_worksheet_as_table(header=True)
    for row in ws:
        _fill_form(row)
    excel.close_workbook()
    

def download_excel_file():
    """Download the excel file"""
    http = HTTP()
    http.download("https://robotsparebinindustries.com/SalesData.xlsx", SALE_FILE_NAME, overwrite=True)

def collect_result():
    page = browser.page()
    page.screenshot(path="output/result.png")

def export_pdf():
    page = browser.page()
    sales_results_html = page.locator("#sales-results").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")

def log_out():
    """Logout from the intranet"""
    page = browser.page()
    page.click("text=Log out")
