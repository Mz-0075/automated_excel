import methods as md  # importing methods package renames as md
import attributes as att  # importing attributes or value package renamed as att
import random
import validation as vd


def main():
    print(att.wb.sheetnames)  # print the all sheets in Excel
    # Telecom sheet__
    source_cell_b5 = att.telecom_sheet['D5']  # assign source cell for formate copied from and cell D5 assign
    target_cell_d11 = att.telecom_sheet['D11']  # assign target cell for formate copied to and cell D11 assign
    source_cell_volte = att.VOLTE_SCFT_NEW_INT_sheet['D3']
    # ___recode_start___
    # load raw report main sheet to a single variable
    # v-lookup function for taking district
    # taking v-lookup value from the main sheet cell B3
    lookup_value = att.main_sheet['B2'].value
    print(f'shine{lookup_value}')
    # taking v-lookup value from the main sheet for each sector ex azi
    lookup_value_sec_a = att.main_sheet['F2'].value
    print(lookup_value_sec_a)
    # taking v-lookup value from the main sheet for each sector ex azi
    lookup_value_sec_b = att.main_sheet['F3'].value
    print(lookup_value_sec_b)
    # taking v-lookup value from the main sheet for each sector ex azi
    lookup_value_sec_c = att.main_sheet['F3'].value
    print(lookup_value_sec_c)
    # ___recode_end___

    district_name = md.vlookup(
        lookup_value)  # calling function and the function return the name of district or not found (nan)
    target_cell_f11 = att.telecom_sheet['F11']  # assign f11 cell to target_cell_f11
    target_cell_h11 = att.telecom_sheet['H11']  # assign f11 cell to target_cell_f11
    # same as(same values) above and F11 and H11 so assign value of D11

    # ___recode_start___
    target_cell_d11.value = target_cell_f11.value = target_cell_h11.value = f"{source_cell_b5.value},{district_name}"  # combine the value of source cell and district name and assign to target cell
    # target_cell_f11.value = target_cell_d11.value  # assign d11 value to f11 cell
    # target_cell_h11.value = target_cell_d11.value  # assign d11 value to h11 cell because both are same value
    # ___recode_end___

    # volte sheet__
    att.wb.remove(att.wb['Volte-SCFT_new_script'])  # for deleting unwanted sheet
    # ___recode_start___

    # md.add_values_to_volte_scft('D')  # adding value to vo-lte scft sheet column D
    # md.add_values_to_volte_scft('E')  # adding value to vo-lte scft sheet column E
    # md.add_values_to_volte_scft('F')  # adding value to vo-lte scft sheet column F
    md.add_values_to_volte_scft(['D', 'E', 'F'])
    att.VOLTE_SCFT_NEW_INT_sheet['D7'].value = round(random.uniform(20, 23), 5)
    att.VOLTE_SCFT_NEW_INT_sheet['E7'].value = round(random.uniform(20, 23), 5)
    att.VOLTE_SCFT_NEW_INT_sheet['f7'].value = round(random.uniform(20, 23), 5)
    att.VOLTE_SCFT_NEW_INT_sheet['D8'].value = round(random.uniform(20, 23), 5)
    att.VOLTE_SCFT_NEW_INT_sheet['E8'].value = round(random.uniform(20, 23), 5)
    att.VOLTE_SCFT_NEW_INT_sheet['F8'].value = round(random.uniform(20, 23), 5)
    # ___recode_end___

    md.add_pass_to_volte_scft()
    md.format_to_volte_scft(
        source_cell_volte)  # apply format volte scft sheet cells from  VOLTE_SCFT_NEW_INT_sheet['D3']
    md.copy_cell_format(source_cell_b5, target_cell_d11)  # calling the function for format copy to d11 from d5
    md.copy_cell_format(source_cell_b5, target_cell_f11)  # calling the function for format copy to f11 from d5
    md.copy_cell_format(source_cell_b5, target_cell_h11)  # calling the function for format copy to h11 from d5

    # clutter sheet__
    # format and rearrange the cells
    md.sheet_cell_format('B2', 'F2', 'Sector1 Clutter', att.Clutter_Photos_sheet)
    md.sheet_cell_format('I2', 'M2', 'Sector2 Clutter', att.Clutter_Photos_sheet)
    md.sheet_cell_format('P2', 'T2', 'Sector3 Clutter', att.Clutter_Photos_sheet)

    md.sheet_cell_format('B23', 'F23', 'Sector1 MT', att.Clutter_Photos_sheet)
    md.sheet_cell_format('I23', 'M23', 'Sector2 MT', att.Clutter_Photos_sheet)
    md.sheet_cell_format('P23', 'T23', 'Sector3 MT', att.Clutter_Photos_sheet)

    md.sheet_cell_format('B44', 'F44', 'Sector1 ET', att.Clutter_Photos_sheet)
    md.sheet_cell_format('I44', 'M44', 'Sector2 ET', att.Clutter_Photos_sheet)
    md.sheet_cell_format('P44', 'T44', 'Sector3 ET', att.Clutter_Photos_sheet)

    md.sheet_cell_format('B65', 'F65', 'Sector1 Label', att.Clutter_Photos_sheet)
    md.sheet_cell_format('I65', 'M65', 'Sector2 Label', att.Clutter_Photos_sheet)
    md.sheet_cell_format('P65', 'T65', 'Sector3 Label', att.Clutter_Photos_sheet)

    md.sheet_cell_format('B86', 'F86', 'Sector1 Tape', att.Clutter_Photos_sheet)
    md.sheet_cell_format('I86', 'M86', 'Sector2 Tape', att.Clutter_Photos_sheet)
    md.sheet_cell_format('P86', 'T86', 'Sector3 Tape', att.Clutter_Photos_sheet)

    md.sheet_cell_format('B107', 'F107', 'Tower', att.Clutter_Photos_sheet)
    # adding Images to each cell like clutter,azimuth etc..
    md.sheet_add_image(att.sec1_clutter, 'B4', att.Clutter_Photos_sheet)
    md.sheet_add_image(att.sec1_mt, 'B25', att.Clutter_Photos_sheet)
    md.sheet_add_image(att.sec1_et, 'B46', att.Clutter_Photos_sheet)
    md.sheet_add_image(att.sec1_label, 'B67', att.Clutter_Photos_sheet)
    md.sheet_add_image(att.sec2_clutter, 'I4', att.Clutter_Photos_sheet)
    md.sheet_add_image(att.sec2_mt, 'I25', att.Clutter_Photos_sheet)
    md.sheet_add_image(att.sec2_et, 'I46', att.Clutter_Photos_sheet)
    md.sheet_add_image(att.sec2_label, 'I67', att.Clutter_Photos_sheet)
    md.sheet_add_image(att.sec3_clutter, 'P4', att.Clutter_Photos_sheet)
    md.sheet_add_image(att.sec3_mt, 'P25', att.Clutter_Photos_sheet)
    md.sheet_add_image(att.sec3_et, 'P46', att.Clutter_Photos_sheet)
    md.sheet_add_image(att.sec3_label, 'P67', att.Clutter_Photos_sheet)
    md.sheet_add_image(att.sec1_tape, 'B88', att.Clutter_Photos_sheet)
    md.sheet_add_image(att.sec2_tape, 'I88', att.Clutter_Photos_sheet)
    md.sheet_add_image(att.sec3_tape, 'P88', att.Clutter_Photos_sheet)
    md.sheet_add_image(att.tower, 'B109', att.Clutter_Photos_sheet)

    # Panoramic sheet__
    # format and rearrange the cells
    md.sheet_cell_format('B1', 'F1', 'Panoramic Photos', att.Panoramic_photos_sheet)
    md.sheet_cell_format('B3', 'F3', '0 Degree', att.Panoramic_photos_sheet)
    md.sheet_cell_format('H3', 'L3', '30 Degree', att.Panoramic_photos_sheet)
    md.sheet_cell_format('N3', 'R3', '60 Degree', att.Panoramic_photos_sheet)
    md.sheet_cell_format('T3', 'X3', '90 Degree', att.Panoramic_photos_sheet)
    md.sheet_cell_format('B24', 'F24', '120 Degree', att.Panoramic_photos_sheet)
    md.sheet_cell_format('H24', 'L24', '150 Degree', att.Panoramic_photos_sheet)
    md.sheet_cell_format('N24', 'R24', '180 Degree', att.Panoramic_photos_sheet)
    md.sheet_cell_format('T24', 'X24', '210 Degree', att.Panoramic_photos_sheet)
    md.sheet_cell_format('B45', 'F45', '240 Degree', att.Panoramic_photos_sheet)
    md.sheet_cell_format('H45', 'L45', '270 Degree', att.Panoramic_photos_sheet)
    md.sheet_cell_format('N45', 'R45', '300 Degree', att.Panoramic_photos_sheet)
    md.sheet_cell_format('T45', 'X45', '330 Degree', att.Panoramic_photos_sheet)
    # Adding images to panoramic sheet
    md.sheet_add_image(att.degree_0, 'B5', att.Panoramic_photos_sheet)
    md.sheet_add_image(att.degree_30, 'H5', att.Panoramic_photos_sheet)
    md.sheet_add_image(att.degree_60, 'N5', att.Panoramic_photos_sheet)
    md.sheet_add_image(att.degree_90, 'T5', att.Panoramic_photos_sheet)
    md.sheet_add_image(att.degree_120, 'B26', att.Panoramic_photos_sheet)
    md.sheet_add_image(att.degree_150, 'H26', att.Panoramic_photos_sheet)
    md.sheet_add_image(att.degree_180, 'N26', att.Panoramic_photos_sheet)
    md.sheet_add_image(att.degree_210, 'T26', att.Panoramic_photos_sheet)
    md.sheet_add_image(att.degree_240, 'B47', att.Panoramic_photos_sheet)
    md.sheet_add_image(att.degree_270, 'H47', att.Panoramic_photos_sheet)
    md.sheet_add_image(att.degree_300, 'N47', att.Panoramic_photos_sheet)
    md.sheet_add_image(att.degree_330, 'T47', att.Panoramic_photos_sheet)

    # Audit details sheet__
    # heading
    md.add_header_value(source_cell_b5)  # calling function for adding heading and format the header cell
    att.Audit_details_sheet['A3'].value = att.main_sheet['B2'].value
    att.Audit_details_sheet['A4'].value = att.main_sheet['B2'].value
    att.Audit_details_sheet['A5'].value = att.main_sheet['B2'].value
    att.Audit_details_sheet['B3'].value = att.main_sheet['F2'].value
    att.Audit_details_sheet['B4'].value = att.main_sheet['F3'].value
    att.Audit_details_sheet['B5'].value = att.main_sheet['F4'].value
    att.Audit_details_sheet['C3'].value = att.main_sheet['A2'].value
    att.Audit_details_sheet['C4'].value = att.main_sheet['A2'].value
    att.Audit_details_sheet['C5'].value = att.main_sheet['A2'].value
    # tower type
    # ___recode_start___
    tower_type = md.tower_type(lookup_value)  # function for vlookup tower type from gis need to recode
    att.Audit_details_sheet['D3'].value = tower_type  # adding tower type to each cell to recode
    att.Audit_details_sheet['D4'].value = tower_type  # adding tower type to each cell
    att.Audit_details_sheet['D5'].value = tower_type  # adding tower type to each cell
    # tower height
    tower_height = md.tower_height(lookup_value)  # function for vlookup tower height form gis need to recode
    att.Audit_details_sheet['E3'].value = tower_height  # adding tower height to each cell need to recode
    att.Audit_details_sheet['E4'].value = tower_height  # adding tower height to each cell
    att.Audit_details_sheet['E5'].value = tower_height  # adding tower height to each cell
    if tower_type == 'RTT':  # checking rtt or not to add building height
        print('calling rtt')
        building_height = md.tower_type_rtt(lookup_value)  # function for building height
        att.Audit_details_sheet['F3'].value = building_height  # adding building height to each cell need to recode
        att.Audit_details_sheet['F4'].value = building_height  # adding building height to each cell
        att.Audit_details_sheet['F5'].value = building_height  # adding building height to each cell
    # antina height
    att.Audit_details_sheet['G3'].value = att.main_sheet['J2'].value
    att.Audit_details_sheet['G4'].value = att.main_sheet['J3'].value
    att.Audit_details_sheet['G5'].value = att.main_sheet['J4'].value
    # pre azimuth
    azi_pre_a = md.azi_pre(lookup_value_sec_a)  # function for finding pre azimuth for first sec
    att.Audit_details_sheet['H3'].value = azi_pre_a
    azi_pre_b = md.azi_pre(lookup_value_sec_b)  # function for finding pre azimuth for second sec
    att.Audit_details_sheet['H4'].value = azi_pre_b
    azi_pre_c = md.azi_pre(lookup_value_sec_c)  # function for finding pre azimuth for third sec
    att.Audit_details_sheet['H5'].value = azi_pre_c
    # post azimuth
    azi_post_a = md.azi_post(lookup_value_sec_a)  # function for finding pre azimuth for first sec
    att.Audit_details_sheet['K3'].value = azi_post_a
    azi_post_b = md.azi_post(lookup_value_sec_b)  # function for finding pre azimuth for second sec
    att.Audit_details_sheet['K4'].value = azi_post_b
    azi_post_c = md.azi_post(lookup_value_sec_c)  # function for finding pre azimuth for third sec
    att.Audit_details_sheet['K5'].value = azi_post_c
    # mt
    att.Audit_details_sheet['I3'].value = att.Audit_details_sheet['L3'].value = att.main_sheet['I2'].value
    att.Audit_details_sheet['I4'].value = att.Audit_details_sheet['L4'].value = att.main_sheet['I3'].value
    att.Audit_details_sheet['I5'].value = att.Audit_details_sheet['L5'].value = att.main_sheet['I4'].value
    # et
    att.Audit_details_sheet['J3'].value = att.Audit_details_sheet['M3'].value = att.main_sheet['H2'].value
    att.Audit_details_sheet['J4'].value = att.Audit_details_sheet['M4'].value = att.main_sheet['H3'].value
    att.Audit_details_sheet['J5'].value = att.Audit_details_sheet['M5'].value = att.main_sheet['H4'].value
    print(target_cell_d11.value)  # print the cell value of telecom sheet cell is D11

    vd.validation()
    att.wb.save(att.path_save)  # saving final automated report


# if __name__ == '__main__':
#     main()
