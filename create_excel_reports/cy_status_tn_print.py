# -*- coding: utf-8 -*-
# ! /usr/local/bin/python3
__author__ = 'p.olifer'

import oraclenv
import excel.edit_excel as xls

REPORT_NAME = ('cy_status_tn_print')

def get_tn_print(print_date, consignee_id,
                 cargo_delivery_id, cargo_receipt_id,
                 idx, in_cont_no):
    con = oraclenv.create_connection()
    cursor = con.cursor()
    params = [print_date,
              consignee_id,
              cargo_delivery_id,
              cargo_receipt_id,
              idx, in_cont_no
              ]
    answer = cursor.callfunc('gcs_export.get_tn_print_xls',
                             oraclenv.REF_CURSOR,
                             params)

    return answer


def write_tn_print(print_date, consignee_id, cargo_delivery_id,
                   cargo_receipt_id, idx, in_cont_no, printer_no):
    data = get_tn_print(print_date, consignee_id,
                        cargo_delivery_id, cargo_receipt_id,
                        idx, in_cont_no)
    report = xls.EditExcelTemplate(REPORT_NAME)
    answer = []

    # parse data
    for report_date, cargo_receipt, consignor, consignee, cargo_delivery, \
        out_ct, in_date, out_date, out_date_sign, \
        carrier_name, carrier, carrier_contract, cargo_receipt_name, \
        operator_name, in_drv_nm, out_drv_nm, out_cmnt, \
        truck_brand, truck_reg_no in data:

        row_shift = 119
        for i in range(4):
            report.write_workbook(4 + row_shift * i, 23, report_date)
            report.write_workbook(7 + row_shift * i, 4, consignor)
            report.write_workbook(7 + row_shift * i, 19, consignee)
            report.write_workbook(12 + row_shift * i, 20, out_ct)
            report.write_workbook(26 + row_shift * i, 20, truck_brand)
            report.write_workbook(35 + row_shift * i, 3, cargo_receipt)
            report.write_workbook(35 + row_shift * i, 20, cargo_delivery)
            report.write_workbook(37 + row_shift * i, 3, in_date)
            report.write_workbook(37 + row_shift * i, 20, report_date)
            report.write_workbook(39 + row_shift * i, 3, in_date)
            report.write_workbook(39 + row_shift * i, 11, out_date)
            report.write_workbook(39 + row_shift * i, 20, report_date)
            report.write_workbook(39 + row_shift * i, 32, report_date)
            report.write_workbook(46 + row_shift * i, 6, operator_name)
            report.write_workbook(49 + row_shift * i, 3, out_drv_nm)
            report.write_workbook(49 + row_shift * i, 20, out_drv_nm)
            report.write_workbook(65 + row_shift * i, 3, report_date)
            report.write_workbook(65 + row_shift * i, 11, carrier_name)
            report.write_workbook(72 + row_shift * i, 5, carrier)
            report.write_workbook(74 + row_shift * i, 5, out_drv_nm)
            report.write_workbook(78 + row_shift * i, 12, truck_brand)
            report.write_workbook(78 + row_shift * i, 26, truck_reg_no)
            report.write_workbook(100 + row_shift * i, 5, carrier_contract)
            report.write_workbook(111 + row_shift * i, 10, report_date)
            report.write_workbook(111 + row_shift * i, 22, carrier_name)
            report.write_workbook(111 + row_shift * i, 35, report_date)
            report.write_workbook(112 + row_shift * i, 5, operator_name)
            report.write_workbook(112 + row_shift * i, 22, out_drv_nm)

            if out_cmnt:
                report.write_workbook(117 + row_shift * i, 5, out_cmnt)
                report.write_workbook(117 + row_shift * i, 31, out_date_sign)

        answer.append(out_ct)
        answer.append(cargo_delivery)
        answer.append(truck_reg_no)
        answer.append(truck_brand)
        answer.append(out_drv_nm)
        answer.append(carrier_name)
        answer.append(out_cmnt)

    report.save_excel()
    report.print_excel_file(printer_name=printer_no)
    report.delete_file(report.report_path)

    return answer


def write_empty_tn_print(print_date, consignor, consignee, cargo_name_1,
                         cargo_name_2, cargo_reception, cargo_delivery,
                         driver_name, power_of_attorney, carrier_name,
                         truck_brand, truck_reg_no, printer_no):
    report = xls.EditExcelTemplate(REPORT_NAME)
    answer = list()
    answer.append(cargo_name_2)
    answer.append(cargo_delivery)
    answer.append(truck_reg_no)
    answer.append(truck_brand)
    answer.append(driver_name)
    answer.append(carrier_name)
    answer.append('')

    p_date = '{year}-{month}-{day}'.format(year=print_date[:4],
                                           month=print_date[4:6],
                                           day=print_date[6:])

    row_shift = 119
    for i in range(4):
        report.write_workbook(4 + row_shift * i, 23, p_date)
        report.write_workbook(7 + row_shift * i, 4, consignor)
        report.write_workbook(7 + row_shift * i, 19, consignee)
        report.write_workbook(12 + row_shift * i, 5, cargo_name_1)
        report.write_workbook(12 + row_shift * i, 20, cargo_name_2)
        report.write_workbook(26 + row_shift * i, 20, truck_brand)
        report.write_workbook(35 + row_shift * i, 3, cargo_reception)
        report.write_workbook(35 + row_shift * i, 20, cargo_delivery)
        # report.write_workbook(37 + row_shift * i, 3, in_date)
        report.write_workbook(37 + row_shift * i, 20, p_date)
        # report.write_workbook(39 + row_shift * i, 3, in_date)
        #report.write_workbook(39 + row_shift * i, 11, out_date)
        report.write_workbook(39 + row_shift * i, 20, p_date)
        report.write_workbook(39 + row_shift * i, 32, p_date)
        # report.write_workbook(46 + row_shift * i, 6, power_of_attorney)
        report.write_workbook(49 + row_shift * i, 3, driver_name)
        report.write_workbook(49 + row_shift * i, 20, driver_name)
        report.write_workbook(65 + row_shift * i, 3, p_date)
        report.write_workbook(65 + row_shift * i, 11, carrier_name)
        report.write_workbook(72 + row_shift * i, 5, carrier_name)
        report.write_workbook(74 + row_shift * i, 5, driver_name)
        report.write_workbook(78 + row_shift * i, 12, truck_brand)
        report.write_workbook(78 + row_shift * i, 26, truck_reg_no)
        # report.write_workbook(100 + row_shift * i, 5, carrier_contract)
        report.write_workbook(111 + row_shift * i, 10, p_date)
        report.write_workbook(111 + row_shift * i, 22, carrier_name)
        report.write_workbook(111 + row_shift * i, 35, p_date)
        # report.write_workbook(112 + row_shift * i, 5, operator_name)
        report.write_workbook(112 + row_shift * i, 22, driver_name)
    report.save_excel()
    report.print_excel_file(printer_name=printer_no)
    report.delete_file(report.report_path)
    
    return answer


if __name__ == '__main__':
    pass
