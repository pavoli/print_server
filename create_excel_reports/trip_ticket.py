# -*- coding: utf-8 -*-
# ! /usr/local/bin/python3
__author__ = 'i.ishmukhametov'

import oraclenv
import excel.edit_excel as xls
import app_config as config
from printers.gru_barcode import create_barcode
from openpyxl.drawing.image import Image as xls_image
from printers.printers import set_duplex


TRIP_TICKET = ('trip_ticket')


def get_trip_ticket_list(reg_emp, date):
    con = oraclenv.create_connection()
    cursor = con.cursor()
    params = [date, reg_emp]
    answer = cursor.callfunc('get_trip_ticket', oraclenv.REF_CURSOR, params)

    return answer


def print_trip_ticket(reg_emp_list, date, printer_no):
    delete_files = []
    set_duplex(printer_no, 2)
    for reg_emp in str.split(reg_emp_list, ','):
        if reg_emp == '':
            continue
        report = xls.EditExcelTemplate(TRIP_TICKET)
        data = get_trip_ticket_list(reg_emp, date)
        row = 12
        # parse data
        for emp_no, local_nm, check_date, area, start_dt, end_dt in data:
            if row == 12:
                report.write_workbook(7, 5, local_nm)
                barcode_image_path = create_barcode(case_no=emp_no, code='code39', label_type='user_id')
                img = xls_image(barcode_image_path)
                report.insert_image(img, 'S2')
                img1 = xls_image(barcode_image_path)
                report.insert_image(img1, 'S68')
            report.write_workbook(row, 1, check_date)
            report.write_workbook(row, 5, area)
            report.write_workbook(row, 16, start_dt)
            report.write_workbook(row, 20, end_dt)
            row += 1
            if row == 26:
                row = 40
        report.save_excel()
        report.print_excel(printer_no=printer_no)
        delete_files.append(report.report_path)
        delete_files.append(barcode_image_path)
    # delete all files
    for x in delete_files:
        report.delete_file(x)

    report.set_answer('printer_name', config.PRINTER_NAMES[printer_no])

    return report.get_answer()


if __name__ == '__main__':
    pass

