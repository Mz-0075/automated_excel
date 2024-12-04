from openpyxl import load_workbook  # import a class from openpyxl package

# Paths
# path_raw_report = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/test.xlsx'  # path of raw report for automation
import os
import tkinter as tk
from tkinter import filedialog
# Initial file path
path_raw_report = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/test.xlsx'

# Check if file exists
if not os.path.exists(path_raw_report):
    # File doesn't exist, ask user to select the file using tkinter
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    path_raw_report = filedialog.askopenfilename(
        title="Select the Excel file",
        filetypes=[("Excel files", "*.xlsx;*.xls")],
    )

    # If the user selects a file, path_raw_report will be updated
    if path_raw_report:
        print(f"File selected: {path_raw_report}")
    else:
        print("No file selected.")
else:
    print(f"File exists at: {path_raw_report}")
path_save = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/4G_L1800_SCFT_REPORT_MAL_EDAV12_OUT.xlsx'  # path to save automated report
path_db = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/4g_db.xlsx'  # path of database should be latest
path_gis = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/RS Spectra_SCFT_4G GIS.xlsx'
sec1_azimuth = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/SEC1/AZIMUTH.jpg'
sec1_clutter = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/SEC1/CLUTTER.jpg'
sec1_mt = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/SEC1/MT.jpg'
sec1_et = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/SEC1/ET.jpg'
sec1_label = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/SEC1/LABEL.jpg'
sec2_azimuth = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/SEC2/AZIMUTH.jpg'
sec2_clutter = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/SEC2/CLUTTER.jpg'
sec2_mt = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/SEC2/MT.jpg'
sec2_et = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/SEC2/ET.jpg'
sec2_label = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/SEC2/label.jpg'
sec3_azimuth = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/SEC3/AZIMUTH.jpg'
sec3_clutter = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/SEC3/CLUTTER.jpg'
sec3_mt = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/SEC3/MT.jpg'
sec3_et = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/SEC3/ET.jpg'
sec3_label = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/SEC3/lABEL.jpg'
sec1_tape = "D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/TOWER+ANTENNA/SEC1 AH.jpg"
sec2_tape = "D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/TOWER+ANTENNA/SEC2 AH.jpg"
sec3_tape = "D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/TOWER+ANTENNA/SEC3 AH.jpg"
tower = "D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/TOWER+ANTENNA/TP.jpg"
degree_0 = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/PANORAMIC/0.jpg'
degree_30 = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/PANORAMIC/30.jpg'
degree_60 = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/PANORAMIC/60.jpg'
degree_90 = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/PANORAMIC/90.jpg'
degree_120 = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/PANORAMIC/120.jpg'
degree_150 = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/PANORAMIC/150.jpg'
degree_180 = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/PANORAMIC/180.jpg'
degree_210 = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/PANORAMIC/210.jpg'
degree_240 = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/PANORAMIC/240.jpg'
degree_270 = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/PANORAMIC/270.jpg'
degree_300 = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/PANORAMIC/300.jpg'
degree_330 = 'D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/reference reports/BADA05/PANORAMIC/330.jpg'

header_list = ['SITE', 'CELL', 'TECHNOLOGY', 'TOWER TYPE', 'TOWER HEIGHT', 'BUILDING HEIGHT', 'ANTINA HEIGHT',
               'PRE AZIMUTH', 'M TILT', 'E TILT', 'POST AZIMUTH', 'M TILT',
               'E TILT']  # heading values for audit details sheet

wb = load_workbook(path_raw_report)  # creating an object to raw report
telecom_sheet = wb['Telecom']  # assign telecom sheet from the raw report to telecom_sheet class
main_sheet = wb['Main Sheet']  # assign Main sheet  from the raw report to main_sheet class
VOLTE_SCFT_NEW_INT_sheet = wb[
    'VOLTE_SCFT_NEW_INT']  # assign VOLTE_SCFT_NEW_INT from the raw report to VOLTE_SCFT_NEW_INT_sheet class
wb.remove(wb['Clutter_Photos'])  # for deleting existing sheet because cell space not correct
wb.create_sheet('Clutter Photos', index=17)  # creating new clutter photos sheet
Clutter_Photos_sheet = wb['Clutter Photos']  # assign Clutter_Photos from the raw report to class
wb.create_sheet('Panoramic Photos', index=18)  # creating new sheet for panoramic photos
Panoramic_photos_sheet = wb['Panoramic Photos']  # create an object for panoramic photos sheet
wb.create_sheet("Audit Details", index=19)  # create a sheet for audit details
Audit_details_sheet = wb['Audit Details']  # create an object for audit details sheet
scft_sheet = wb["SCFT"]  # assign scft from raw report
SCFT_INT_sheet = wb['SCFT_INT']  # assign SCFT_INT from raw report
