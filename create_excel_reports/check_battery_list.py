# -*- coding: utf-8 -*-
# ! /usr/local/bin/python3
__author__ = 'p.olifer'

import oraclenv
import excel.edit_excel as xls


REPORT_NAME = ('check_battery_list')


def get_battery_list(from_date='20190521', to_date='20190521'):
    con = oraclenv.create_connection()
    cursor = con.cursor()
    params = [from_date, to_date]
    answer = cursor.callfunc('gcs_equipment.get_battery_charge_record_xls', oraclenv.REF_CURSOR, params)

    # con.commit()
    # cursor.close()
    # con.close()

    return answer


def write_battery_list(from_date, to_date, printer_no):
    row = 7

    data = get_battery_list(from_date, to_date)
    report = xls.EditExcelTemplate(REPORT_NAME)

    # copy style for all rows in report
    style = report.ws['A7']._style

    # parse data
    for i in data:
        # write data from 1 till 11 columns
        for j in range(11):
            report.write_workbook_style(row, j+1, i[j], style)
        row += 1
        report.add_row(row)

    # write header info
    import datetime
    current_date = datetime.date.today()
    report.write_workbook(3, 3, current_date)

    report.save_excel()
    report.print_excel_file(printer_name=printer_no)
    report.delete_file(report.report_path)

    report.set_answer('printer_name', printer_no)

    return report.get_answer()


if __name__ == '__main__':
    # print(write_battery_list())
    pass
