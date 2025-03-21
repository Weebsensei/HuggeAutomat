import time
import os
import openpyxl
import Var
from selenium import webdriver
from selenium.webdriver.support.select import Select

# Extremely basic automation script for filling out onlineform with data from excel
# Small sleep timers since onlineform was slow
# It is ugly, but it works :)
 
driver = webdriver.Chrome(options=Var.c_options)
url = driver.command_executor._url
session_id = driver.session_id

os.getcwd()

def nextPage():
    if Var.next_page:
        Var.next_page.click()
    else: # Navigation Button not found, wait and try again.
        time.sleep(0.2)
        nextPage()

def ID():
    if Var.id_element:
        Var.id_element.click()
    else: # Element not found, wait and try again.
        time.sleep(0.2)
        ID()

def dateYear(year_, quarter_):
    if Var.year and Var.quarter:
        Select(Var.year).select_by_visible_text(year_)
        Select(Var.quarter).select_by_visible_text(quarter_)
    else: # Element not found, wait and try again.
        time.sleep(0.2)
        dateYear(year_, quarter_)

import pandas as pd
data = pd.ExcelFile(Var.file)
ps = openpyxl.load_workbook(Var.file)
sheet = ps[Var.blad]

# Each iteration fills out one form
for row in range(780, sheet.max_row+1):
    year_ = str(sheet['A' + str(row)].value)
    quarter_ = ("Kvartal " + str(sheet['B' + str(row)].value))
    id_ = str(sheet['C' + str(row)].value)
    code_ = str(sheet['D' + str(row)].value)
    ref_ = str(sheet['E' + str(row)].value)
    bcode_ = str(sheet['F' + str(row)].value).replace(';', '')
    ucode_ = str(sheet['G' + str(row)].value)

    # Print row to see last excel row processed in case of crash :)
    print(row)

    Var.form

    # ID verification
    ID()
    nextPage()

    Var.confirm.click()
    nextPage()

    # Default static inputs
    Var.emailField.send_keys(Var.email)
    Var.teleField.send_keys(Var.tele)
    Var.orgField.send_keys(Var.org)
    Var.nameField.send_keys(Var.name)
    Var.emailConField.send_keys(Var.email)
    Var.teleConField.send_keys(Var.tele)
    nextPage()

    # Year and quarter XD
    dateYear(year_, quarter_)
    nextPage()

    Var.idField.send_keys(id_)
    Var.manualField.send_keys(code_)
    Var.amountField.send_keys("0")
    Var.refField.send_keys(ref_)
    Select(Var.methodOption).select_by_value(bcode_)
    time.sleep(2)
    Select(Var.methodCode).select_by_index(1)
    Var.amount2.send_keys("0")
    Var.amount3.send_keys("0")
    nextPage()
    time.sleep(0.5) # Seems to always need a bit of waiting
    nextPage()      # If page still has not loaded the function also waits
    time.sleep(0.5) 
    Var.formRefresh
