# -*- coding: utf-8 -*-
__author__ = 'p.olifer'

import oraclenv
import excel.edit_excel as xls


REPORT_NAME = ('waybill')


def get_waybill(from_date, to_date, truck_no):
    con = oraclenv.create_connection()
    cursor = con.cursor()

    header_cur = cursor.var(oraclenv.REF_CURSOR)
    driver_cur = cursor.var(oraclenv.REF_CURSOR)
    fuel_cur = cursor.var(oraclenv.REF_CURSOR)

    params = [from_date, to_date, truck_no, header_cur, driver_cur, fuel_cur]

    answer = cursor.callproc('gcs_equipment.get_check_list_truck_for_xls', params)

    header_info = answer[3]
    driver_info = answer[4]
    fuel_info = answer[5]

    return [header_info, driver_info, fuel_info]


def write_waybill(from_date, to_date, truck_no, printer_no):
    header_info, driver_info, fuel_info = get_waybill(from_date, to_date, truck_no)

    report = xls.EditExcelTemplate(REPORT_NAME)

    # write header
    for i in header_info:
        report.write_workbook(row_dest=5, column_dest=6, value=i[0])
        report.write_workbook(row_dest=7, column_dest=2, value=i[1])
        report.write_workbook(row_dest=17, column_dest=4, value=i[2])
        report.write_workbook(row_dest=19, column_dest=4, value=i[3])

    # write driver information
    row = 31
    row_to_jump = [
        64,
        103,
    ]

    for i in driver_info:
        if row in row_to_jump:
            row += 6
        report.write_workbook(row_dest=row, column_dest=3, value=i[2])
        report.write_workbook(row_dest=row, column_dest=5, value=i[3])

        row += 1

    # write fuel information
    row = 31
    for i in fuel_info:
        if row in row_to_jump:
            row += 6

        report.write_workbook(row_dest=row, column_dest=2, value=i[0])
        report.write_workbook(row_dest=row, column_dest=6, value=i[1])
        report.write_workbook(row_dest=row, column_dest=7, value=i[2])
        report.write_workbook(row_dest=row, column_dest=8, value=i[3])
        report.write_workbook(row_dest=row, column_dest=9, value=i[4])

        row += 3

    report.save_excel()
    report.print_excel_file(printer_name=printer_no)
    report.delete_file(report.report_path)
    report.set_answer('printer_name', printer_no)

    return report.get_answer()


if __name__ == '__main__':
    # data_1, data_2 = get_waybill(from_date='20190814', to_date='20190814', truck_no='TT044')
    # write_waybill(from_date='20191001', to_date='20191031', truck_no='TT044', printer_no='\\\\grusafeq02\\RUPRN_CC_UNPACKING01')
    pass
