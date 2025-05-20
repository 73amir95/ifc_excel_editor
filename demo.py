import xlwings as xw
import ifcopenshell
from ifccsv import IfcCsv
import time

def is_excel_open():
    apps = xw.apps
    return len(apps) > 0

def import_and_update_model(model, csv_file):
    ifc_csv = IfcCsv()
    ifc_csv.Import(model, csv_file)
    model.write(file_path.split(".")[0] + "_updated.ifc")

file_path = "your_custom_IFC_file.ifc"
model = ifcopenshell.open(file_path)
elements = ifcopenshell.util.selector.Selector.parse(model, ".IfcElement")
attributes = ["Name", "Description", "PredefinedType"]

ifc_csv = IfcCsv()
ifc_csv.export(model, elements, attributes, output="file.csv", format="csv", delimiter=",", null="-")

excel_app = xw.App(visible=True)          # opening excel
wb = excel_app.books.open("file.csv")     # opens the csv file in it


while is_excel_open():                    # monitoring Excel status and wait for it to close
    time.sleep(1)                         # waiting before checking again

import_and_update_model(model, 'file.csv')

excel_app.quit()
