import attributes as att


# ____________*VALIDATION*_____________
def validation():
    # checking site id correct or not with cell file
    if att.scft_sheet['B2'].value == att.main_sheet['B2'].value:
        print("Site id ok")
    else:
        print("No site id/miss match, please add 'site_id' in cell file.")
    # checking site name correct or not with cell file
    if att.scft_sheet['D2'].value == att.main_sheet['C2'].value:
        print("Site name is ok")
    else:
        print("Site name is not correct")
    # checking cell id is present or not
    if att.scft_sheet['F2'].value is None:
        print("please add cell id in cell file")
    else:
        print("Cell id is ok")
    # checking erfcn with main sheet and scft sheet
    if att.scft_sheet['D9'].value == 265:  # checking 265 or not
        if att.main_sheet['O2'].value == '265':
            print('erfcn ok')
            for i in ['D50', 'E50', 'F50']:
                att.scft_sheet[i].value = "FDD"
        else:  # when miss match with scft sheet assign scft sheet value to main sheet
            att.main_sheet['O2'].value = att.main_sheet['O3'].value = att.main_sheet['O4'].value = '265'
            print("erfcn changed to 265")
            for i in ['D50', 'E50', 'F50']:
                att.scft_sheet[i].value = "FDD"
    if att.scft_sheet['D9'].value == 1551:  # checking 1551 or not
        if att.main_sheet['O2'].value == '1551':
            print('erfcn ok')
            for i in ['D50', 'E50', 'F50']:
                att.scft_sheet[i].value = "FDD"
        else:  # when miss match with scft sheet assign scft sheet value to main sheet
            att.main_sheet['O2'].value = att.main_sheet['O3'].value = att.main_sheet['O4'].value = '1551'
            print("erfcn changed to 1551")
            for i in ['D50', 'E50', 'F50']:
                att.scft_sheet[i].value = "FDD"
    if att.scft_sheet['D9'].value == 39150:  # checking 39150 or not
        if att.main_sheet['O2'].value == '39150':
            print('erfcn ok')
            for i in ['D50', 'E50', 'F50']:
                att.scft_sheet[i].value = "TDD"
        else:  # when miss match with scft sheet assign scft sheet value to main sheet
            att.main_sheet['O2'].value = att.main_sheet['O3'].value = att.main_sheet['O4'].value = '39150'
            print("erfcn changed to 39150")
            for i in ['D50', 'E50', 'F50']:
                att.scft_sheet[i].value = "TDD"

    # inter and intra handover YES
    for i in ['D21', 'E21', 'F21', 'D22', 'E22', 'F22']:
        att.scft_sheet[i].value = "YES"
    for i in ['D23', 'E23', 'F23']:
        att.scft_sheet[i].value = "NA"

    #  Checking SCFT_INT failed kpi
    for row in att.SCFT_INT_sheet.iter_rows(min_row=3, max_row=14, min_col=7, max_col=9):
        for cell in row:
            if cell.value == 'Fail':
                print(f"failed kpi is {cell}")
