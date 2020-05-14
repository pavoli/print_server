# -*- coding: utf-8 -*-
# ! /usr/local/bin/python3
__author__ = 's.ilnitskiy'

import oraclenv
import excel.edit_excel as xls
import app_config as config

FUEL_CONTROL_REPORT_1 = ('fuel_control_report_1')
FUEL_CONTROL_REPORT_2 = ('fuel_control_report_2')
FUEL_CONTROL_REPORT_3 = ('fuel_control_report_3')


def get_daily_info_list(move_seq):
    con = oraclenv.create_connection()
    cursor = con.cursor()
    params = [move_seq]
    answer_1 = cursor.callfunc('gcs_export.get_fuel_report_1', oraclenv.REF_CURSOR, params)
    answer_2 = cursor.callfunc('gcs_export.get_fuel_report_2', oraclenv.REF_CURSOR, params)

    return [answer_1, answer_2]


def get_monthly_info_list(month):
    con = oraclenv.create_connection()
    cursor = con.cursor()
    params = [month]
    answer = cursor.callfunc('gcs_export.get_fuel_report_3', oraclenv.REF_CURSOR, params)

    return answer


def print_daily_fuel_report(move_seq, printer_no):
    data_1, data_2 = get_daily_info_list(move_seq)
    report_1 = xls.EditExcelTemplate(FUEL_CONTROL_REPORT_1)
    report_2 = xls.EditExcelTemplate(FUEL_CONTROL_REPORT_2)

    # parse data
    for i in data_1:
        report_1.write_workbook(4, 1, i[0])
        report_1.write_workbook(7, 2, i[1])
        report_1.write_workbook(8, 2, i[2])
        report_1.write_workbook(9, 2, i[3])
        report_1.write_workbook(16, 1, i[4])
        report_1.write_workbook(16, 3, i[5])
        report_1.write_workbook(20, 3, i[6])
        report_1.write_workbook(21, 3, i[7])

    for i in data_2:
        report_2.write_workbook(5, 4, i[0])
        report_2.write_workbook(9, 4, i[1])
        report_2.write_workbook(10, 4, i[2])
        report_2.write_workbook(11, 4, i[3])
        report_2.write_workbook(21, 4, i[4])
        report_2.write_workbook(7, 7, i[5])

    report_1.save_excel()
    report_1.print_excel(printer_no=printer_no)
    report_1.delete_file(report_1.report_path)
    report_1.set_answer('printer_name', config.PRINTER_NAMES[printer_no])

    report_2.save_excel()
    report_2.print_excel(printer_no=printer_no)
    report_2.delete_file(report_2.report_path)
    report_2.set_answer('printer_name', config.PRINTER_NAMES[printer_no])

    return [report_1.get_answer(), report_2.get_answer()]


def print_monthly_fuel_report(month, printer_no):
    data = get_monthly_info_list(month)
    report = xls.EditExcelTemplate(FUEL_CONTROL_REPORT_3)

    # parse data
    for i in data:
        report.write_workbook(10, 5, i[0])
        report.write_workbook(10, 6, i[1])
        report.write_workbook(10, 7, i[2])
        report.write_workbook(12, 1, i[3])

        report.write_workbook(20, 3, i[4])
        report.write_workbook(20, 5, i[5])
        report.write_workbook(20, 6, i[6])
        report.write_workbook(20, 7, i[5]+i[6]-i[7])
        report.write_workbook(20, 8, i[7])
        report.write_workbook(21, 3, i[8])
        report.write_workbook(21, 5, i[9])
        report.write_workbook(21, 6, i[10])
        report.write_workbook(21, 7, i[9]+i[10]-i[11])
        report.write_workbook(21, 8, i[11])
        report.write_workbook(22, 3, i[12])
        report.write_workbook(22, 5, i[13])
        report.write_workbook(22, 6, i[14])
        report.write_workbook(22, 7, i[13]+i[14]-i[15])
        report.write_workbook(22, 8, i[15])
        report.write_workbook(23, 3, i[16])
        report.write_workbook(23, 5, i[17])
        report.write_workbook(23, 6, i[18])
        report.write_workbook(23, 7, i[17]+i[18]-i[19])
        report.write_workbook(23, 8, i[19])
        report.write_workbook(24, 3, i[20])
        report.write_workbook(24, 5, i[21])
        report.write_workbook(24, 6, i[22])
        report.write_workbook(24, 7, i[21]+i[22]-i[23])
        report.write_workbook(24, 8, i[23])
        report.write_workbook(25, 3, i[24])
        report.write_workbook(25, 5, i[25])
        report.write_workbook(25, 6, i[26])
        report.write_workbook(25, 7, i[25]+i[26]-i[27])
        report.write_workbook(25, 8, i[27])

        report.write_workbook(27, 1, i[28])
        report.write_workbook(27, 7, i[29])
        report.write_workbook(32, 1, i[30])
        report.write_workbook(32, 7, i[31])
        report.write_workbook(35, 1, i[32])
        report.write_workbook(35, 7, i[33])
        report.write_workbook(38, 1, i[34])
        report.write_workbook(38, 7, i[35])

    report.save_excel()
    report.print_excel(printer_no=printer_no)
    report.delete_file(report.report_path)
    report.set_answer('printer_name', config.PRINTER_NAMES[printer_no])

    return report.get_answer()


if __name__ == '__main__':
    pass
