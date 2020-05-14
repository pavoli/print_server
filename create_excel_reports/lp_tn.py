# -*- coding: utf-8 -*-
# ! /usr/local/bin/python3
__author__ = 'i.ishmukhametov'

import oraclenv
import excel.edit_excel as xls
import app_config as config
from printers.printers import set_duplex


LP_TN = 'lp_tn'


def get_lp_tn_list(pool_seq, print_date, vendor, transporter, truck, driver, qty, bqty):
    con = oraclenv.create_connection()
    cursor = con.cursor()
    params = [pool_seq, print_date, vendor, transporter, truck, driver, qty, bqty]
    answer = cursor.callfunc('get_lp_tn', oraclenv.REF_CURSOR, params)

    return answer

def print_lp_tn(pool_seq, print_date, vendor, transporter, truck, driver, qty, bqty, printer_no):
    set_duplex(printer_no, 2)
    report = xls.EditExcelTemplate(LP_TN)
    copies = 0
    data = get_lp_tn_list(pool_seq, print_date, vendor, transporter, truck, driver, qty, bqty)
    for reg_dt, vend_name, vend_addr, tel_no, vend_tel_no, qty, bqty, vend_addr2, drv_nm, drv_contract, trpr_nm, address, contract, brand, ton, truck_cd, vend_contract, pay_det, paper_qty in data:
        copies = paper_qty
        report.write_workbook(3, 61, reg_dt)
        report.write_workbook(7, 2, 'ООО “ГЛОВИС РУС“,197374,г. Санкт-Петербург, ул. Савушкина, д.126, лит.Б, пом.62-Н, ИНН 7805464287, КПП 781401001 по поручению ' + vend_name + ' ' + vend_addr)
        report.write_workbook(7, 53, vend_name + ', ' + vend_addr)
        report.write_workbook(9, 2, tel_no)
        report.write_workbook(9, 53, vend_tel_no)
        text_for_boxes = ''
        if (qty != '0') and (bqty != '0'):
            text_for_boxes = qty + ' транспортировочных тележек и ящики пластиковые ' + bqty + ' штук'
        elif (qty != '0'):
            text_for_boxes = qty + ' транспортировочных тележек'
        else:
            text_for_boxes = 'Ящики пластиковые ' + bqty + ' штук'
        report.write_workbook(14, 2, text_for_boxes)
        report.write_workbook(33, 53, vend_name + ' ' + vend_addr2)
        report.write_workbook(35, 2, reg_dt)
        report.write_workbook(35, 53, reg_dt)
        report.write_workbook(37, 2, reg_dt)
        report.write_workbook(37, 53, reg_dt)
        report.write_workbook(39, 2, reg_dt)
        report.write_workbook(39, 53, reg_dt)
        report.write_workbook(41, 2, text_for_boxes)
        report.write_workbook(41, 53, text_for_boxes)
        report.write_workbook(43, 28, drv_nm)
        report.write_workbook(43, 87, drv_nm)
        report.write_workbook(59, 2 , reg_dt)
        report.write_workbook(59, 16, drv_nm)
        report.write_workbook(59, 72, drv_contract)
        text_for_transporter = ''
        if contract != ' ':
            text_for_transporter = trpr_nm + ' ' + address + ' по поручению ООО “ГЛОВИС РУС“,197374,г. Санкт-Петербург, ул. Савушкина, д.126, лит.Б, пом.62-Н, ИНН 7805464287, КПП 781401001'
        else:
            text_for_transporter = trpr_nm + ' ' + address
        report.write_workbook(62, 2, text_for_transporter)
        report.write_workbook(62, 53, drv_nm)
        report.write_workbook(71, 4, brand)
        report.write_workbook(71, 41, ton)
        report.write_workbook(71, 79, truck_cd)
        if contract != ' ':
            report.write_workbook(91, 2, '"На основании договора ' + contract + ' между ' + trpr_nm + ' и ООО “ГЛОВИС РУС" и на основании договора ' + vend_contract + ' между ' + vend_name + ' и ООО “ГЛОВИС РУС”')
        report.write_workbook(99, 2, vend_name + ' ' + vend_addr + ' ' + pay_det)
        report.write_workbook(102, 53, trpr_nm)
        report.write_workbook(103, 62, drv_nm)
        report.write_workbook(104, 19, reg_dt)
        report.write_workbook(104, 53, drv_contract)
        report.write_workbook(104, 99, reg_dt)
    report.save_excel()
    # print report
    for x in range(copies):
        report.print_excel(printer_no)
    report.delete_file(report.report_path)
    report.set_answer('printer_name', config.PRINTER_NAMES[printer_no])
    return report.get_answer()


if __name__ == '__main__':
    pass
