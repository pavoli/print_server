__author__ = 'p.olifer'

import oraclenv
import excel.edit_excel as xls
from printers.gru_barcode import create_barcode
from PIL import Image
from openpyxl.drawing.image import Image as xls_image
from app_config import EXCEL_IMAGE_FOLDER_ROOT_PATH


REPORT_NAME = ('shipping_mark')
ARROW_UP_PATH = EXCEL_IMAGE_FOLDER_ROOT_PATH.format('arrow_up.jpg')


def get_caseno_data(case_no, loc_tp):
    '''
    outside_data = {
        'ORIGINAL LIST',
        'HFGSJB1498YC8639',
        '90 Kg',
        'JS',
        '2020-05-16',
        'pavoli (Pavel Olifer)',
        'C',
        '5',
        '60',
        '15'
    }

    detail_data = [
        '1', '87610-H0330MZH', 'MIRROR', 'SL-08-31', 10, 7, 1,
        '2', '87610-H0330RMZ', 'MIRROR', 'SL-08-21', 10, 7, 2,
        '3', '87610-H0330RHM', 'MIRROR', 'SL-08-25', 10, 7, 4,
        '4', '87610-H0330MZH', 'MIRROR', 'SL-10-41', 10, 5, 2,
        '5', '87610-H0330MZW', 'MIRROR', 'SL-03-21', 10, 7, 3,
        '6', '87610-H0330MZT', 'MIRROR', 'SL-08-31', 10, 6, 3,
    ]
    '''
    con = oraclenv.create_connection()
    cursor = con.cursor()

    outside_cur = cursor.var(oraclenv.REF_CURSOR)
    detail_cur = cursor.var(oraclenv.REF_CURSOR)

    params = [case_no, loc_tp, outside_cur, detail_cur]

    answer = cursor.callproc('gcs_inventory.get_shipping_mark_for_xls', params)

    outside_data = answer[2]
    detail_data = answer[3]

    # con.commit()
    # cursor.close()
    # con.close()

    return [outside_data, detail_data]


def write_caseno_data(printer_no=15):
    to_be_deleted = []
    answer = []

    data = get_caseno_data(case_no='HFGSJB1498YC8639', loc_tp='CY')
    report = xls.EditExcelTemplate(REPORT_NAME)
    to_be_deleted.append(report.report_path)

    # parse data
    main_data, detail_data = data

    for case_type, case_no, case_weight, material_group, print_date, emp_name, shift_cd, part_no_qty, part_qty, box_qty in main_data:
        # get an barcode image
        barcode_image_path = create_barcode(
            case_no=case_no, label_type='shipping_mark', code='code128')
        # input image into cell in excel
        img = xls_image(barcode_image_path)
        report.insert_image(img, 'P15')

        arrow_up = xls_image(ARROW_UP_PATH)
        report.insert_image(arrow_up, 'R02')

        # write data
        report.write_workbook(3, 10, case_type)
        report.write_workbook(6, 2, case_no)
        report.write_workbook(15, 5, case_weight)
        report.write_workbook(16, 10, material_group)
        report.write_workbook(21, 5, print_date)
        report.write_workbook(24, 5, emp_name)
        report.write_workbook(24, 17, shift_cd)
        report.write_workbook(48, 3, part_no_qty)
        report.write_workbook(48, 16, part_qty)
        report.write_workbook(48, 18, box_qty)

        to_be_deleted.append(barcode_image_path)

    start_row = 27
    for row_num, part_no, part_name, loc_no, part_qty, qty_in_box, box_qty in detail_data:
        report.write_workbook(start_row, 2, row_num)
        report.write_workbook(start_row, 3, part_no)
        report.write_workbook(start_row, 8, part_name)
        report.write_workbook(start_row, 14, loc_no)
        report.write_workbook(start_row, 16, part_qty)
        report.write_workbook(start_row, 17, qty_in_box)
        report.write_workbook(start_row, 18, box_qty)

        start_row += 1

    report.save_excel()
    report.print_excel(printer_no=printer_no)

    for i in to_be_deleted:
        report.delete_file(i)
        
    report.set_answer('printer_name', printer_no)

    return answer


if __name__ == "__main__":
    write_caseno_data()
