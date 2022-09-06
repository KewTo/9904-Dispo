from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from glob import glob
from openpyxl import load_workbook
import xlwings as xw
import re
import win32com.client

# Automated 9904 disposition list for COAT Module. This is intended to grab the excel file from the Intel website, apply
# certain filters and formula in conjuncture with previous 9904 dispo list from the past. With the previous dispo list
# and VLOOKUP, it will apply which BLANKs belong to which engineer. It will sort and color code  names for easier view.


def grab_file():
    PATH = "C:\Program Files (x86)\chromedriver.exe"
    # From Chrome grab a list of COAT storage information
    driver = webdriver.Chrome(PATH)
    driver.get('https://imosc-ebiz.intel.com/imobi/module_stores.asp')
    # Select COAT module OP.9904 from drop down list and download Excel file
    driver.find_element(By.XPATH, "/html/body/div/div/div[2]/center/div[1]/select").click
    driver.find_element(By.XPATH, "/html/body/div/div/div[2]/center/div[1]/select/option[5]").click
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "element"))
        )
    except:
        driver.quit()


def recent_file():
    # Find latest Excel file
    filepath = r'C:\Users\kevinto\Downloads\Search by Module Stores*.xlsx'
    latest_file = max(glob(filepath), key=os.path.getmtime)
    return latest_file


wb = load_workbook(recent_file())
ws = wb.active
Dispo_Dimensions = ws.dimensions


def filter_headers():
    # Find all header titles column number for filtering
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == 'ENG_LOT_OWNER':
                ENG_LOT_OWNER1 = cell.coordinate
                ENG_LOT_OWNER_Full = ENG_LOT_OWNER1 + ':' + ENG_LOT_OWNER1[0] \
                                     + str(ws.max_row)
    return ENG_LOT_OWNER1


ENG_LOT_OWNER_Full = filter_headers() + ':' + filter_headers()[0] \
                     + str(ws.max_row)


def headers(x: str):
    header = []
    for i in range(1, ws.max_column):
        header.append(ws.cell(row=2, column=i).value)
    index = header.index(x) + 1
    return index


def apply_filter():
    # Filter various headers
    xw.sheets[0].api.Range(Dispo_Dimensions).AutoFilter(headers('LOT'), Criteria1 := 'BLNK*', Operator := 2, Criteria2 := 'BEUVF*')
    xw.sheets[0].api.Range(Dispo_Dimensions).AutoFilter(headers('Rack'), 'NOT-IN-FAB')
    xw.sheets[0].api.Range(Dispo_Dimensions).AutoFilter(headers('DAO'), Criteria1 := '<200')
    xw.sheets[0].api.Range(Dispo_Dimensions).AutoFilter(headers('GOLDEN_MASK'), 'Blanks')


def copy_info():
    # Copy previous 9904 dispo list to a new sheet
    info_workbook = sorted(glob(r'C:\Users\kevinto\Downloads\Search by Module Stores*.xlsx'), key=os.path.getmtime)[-2]
    wb1 = xw.Book(info_workbook)  # Initial file with data
    wb2 = xw.Book(recent_file())  # Target file
    ws1 = wb1.sheets[0]  # [1]
    ws1.api.Copy(After=wb2.sheets[0].api)
    return ws1


def vloop():
    # Apply VLOOKUP Formula on Status Column
    for x in range(2, len(xw.sheets[0].range('K1:K213').rows)):
        xw.sheets[0].range('K' + str(xw.sheets[0].range('K1:K213')[x].row)).value = '=VLOOKUP(C' + str(
            x + 1) + ', DISPO' + '!' + 'A:F, 6, FALSE)'
    # Copy vlookup on Status to ENG_LOT_OWNER
    head = []
    for i in range(2, len(xw.sheets[0].range('K1:K213').rows)):
        if xw.sheets[0].range('K1:K213')[i].value is not None:
            head.append('K' + str(xw.sheets[0].range('K1:K213')[i].row))
    for i in head:
        xw.sheets[0].range(i).value = xw.sheets[0].range(i).options(ndim=2).value
        xw.sheets[0].range(i).copy()
        xw.sheets[0].range('J' + str(i[1:])).paste()


def unsorted():
    # Sort ENG_LOT_OWNER by name
    excel = win32com.client.Dispatch("Excel.Application")
    wb_win32 = excel.Workbooks.Open(recent_file())
    ws_win32 = wb_win32.Worksheets('Sheet1')
    ws_win32.Range('J3:J213').Sort(Key1=ws_win32.Range('J1'), Order1=1, Orientation=1)


def resort():
    # Unsort and resort filters
    xw.sheets[0].api.AutoFilterMode = False
    xw.sheets[0].api.Range(Dispo_Dimensions).AutoFilter(headers('Rack'), 'NOT-IN-FAB')
    xw.sheets[0].api.Range(Dispo_Dimensions).AutoFilter(headers('DAO'), Criteria1 := '<200')


def owner_name():
    # Color code ENG_LOT_OWNER
    for index, elem in enumerate(xw.sheets[0].range('J2:J213').value):
        if re.match(r'^BWOLSON', str(elem)):
            xw.Range(filter_headers()[0] + str(index + 1)).color = (204, 0, 0)
        elif re.match(r'^GLUU', str(elem)):
            xw.Range(filter_headers()[0] + str(index + 1)).color = (204, 102, 0)
        elif re.match(r'^JABELARD', str(elem)):
            xw.Range(filter_headers()[0] + str(index + 1)).color = (204, 0, 102)
        elif re.match(r'^JKBOSWOR', str(elem)):
            xw.Range(filter_headers()[0] + str(index + 1)).color = (204, 204, 0)
        elif re.match(r'^JRNISKAL', str(elem)):
            xw.Range(filter_headers()[0] + str(index + 1)).color = (0, 204, 0)
        elif re.match(r'^MMARCINK', str(elem)):
            xw.Range(filter_headers()[0] + str(index + 1)).color = (0, 204, 204)
        elif re.match(r'^SCPRICE', str(elem)):
            xw.Range(filter_headers()[0] + str(index + 1)).color = (0, 0, 204)
        elif re.match(r'^YUNPINGF', str(elem)):
            xw.Range(filter_headers()[0] + str(index + 1)).color = (102, 0, 204)
        else:
            xw.Range(filter_headers()[0] + str(index + 1)).color = (204, 0, 204)


def delete_extra():
    # Delete more GOLDEN_MASK without Y
    for row in ws.rows:
        for cell in row:
            if cell.value in ("BLNK339600", "BLNK339601"):
                xw.sheets[0].range('A' + str(cell.row) + ':' + 'V' + str(cell.row)).delete()  # A97:V97
    # Delete unnecessary rows/columns
    xw.sheets[0].range('A1:V1').delete()
    xw.sheets[0].range('B1:B' + str(xw.sheets[0].range(copy_info().dimensions).current_region.last_cell.row)).delete()
    xw.sheets[0].range('A1:A' + str(xw.sheets[0].range(copy_info().dimensions).current_region.last_cell.row)).delete()


grab_file()
apply_filter()
vloop()
unsorted()
resort()
owner_name()
delete_extra()

if __name__ == '__main__':
    pass
