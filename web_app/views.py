# -*- coding: utf-8 -*-
# ! /usr/local/bin/python3


__author__ = 'p.olifer'

import os
from flask import render_template, request, make_response, redirect
from web_app import app
from create_excel_reports import *


@app.errorhandler(404)
def not_found_error(error):
    return render_template('404.html'), 404


@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('home.html', title='')


@app.route('/excel_reports/check_battery_list', methods=['GET'])
def check_battery_list_xl():
    if request.method == 'GET':
        from_date = request.args.get('from_date', type=str)
        to_date = request.args.get('to_date', type=str)
        printer_no = request.args.get('printer_no', type=str)

    answer = check_battery_list.write_battery_list(from_date, to_date, printer_no)
    print('Answer: {}'.format(answer))

    return render_template('report.html', title='', answer=answer)


@app.route('/excel_reports/fuel_control_reports', methods=['GET'])
def fuel_control_reports_xl():
    if request.method == 'GET':
        move_seq = request.args.get('move_seq', type=str)
        month = request.args.get('month', type=str)
        printer_no = request.args.get('printer_no', type=int)

        if move_seq == '0':
            answer = fuel_control_reports.print_monthly_fuel_report(month, printer_no)
            print('Answer: {}'.format(answer))

            return render_template('report.html', title='', answer=answer)
        else:
            answer_1, answer_2 = fuel_control_reports.print_daily_fuel_report(move_seq, printer_no)
            print('Answer: {}'.format(answer_1))
            print('Answer: {}'.format(answer_2))

            return render_template('report.html', title='', answer=answer_1)


@app.route('/excel_reports/waybill', methods=['GET'])
def waybill_xl():
    if request.method == 'GET':
        from_date = request.args.get('from_date', type=str)
        to_date = request.args.get('to_date', type=str)
        truck_no = request.args.get('truck_no', type=str)
        printer_no = request.args.get('printer_no', type=str)

        answer = waybill.write_waybill(from_date, to_date, truck_no, printer_no)
        print('Answer: {}'.format(answer))

        return render_template('report.html', title='', answer=answer)


@app.route('/excel_reports/cy_status_tn_print', methods=['GET'])
# ?idx=201910300403&in_cont_no=TEMU6562485&print_date=20191101&printer_no=15&consignee_id=RW104&cargo_receipt_id=P06_1&cargo_delivery_id=RW100
def cy_status_tn_print_xl():
    if request.method == 'GET':
        idx = request.args.get('idx', type=str)
        cont_no = request.args.get('in_cont_no', type=str)
        print_date = request.args.get('print_date', type=str)
        consignee = request.args.get('consignee', type=str)
        cargo_receipt = request.args.get('cargo_receipt', type=str)
        cargo_delivery = request.args.get('cargo_delivery', type=str)
        printer_name = request.args.get('printer_no', type=str)

        answer = cy_status_tn_print.write_tn_print(
            print_date=print_date,
            consignee_id=consignee,
            cargo_delivery_id=cargo_delivery,
            cargo_receipt_id=cargo_receipt,
            idx=idx,
            in_cont_no=cont_no,
            printer_no=printer_name
        )

        return render_template('report_tn_print.html', title='', answer=answer)


@app.route('/excel_reports/cy_status_empty_tn_print', methods=['GET'])
# ?idx=201910300403&in_cont_no=TEMU6562485&print_date=20191101&printer_no=15&consignee_id=RW104&cargo_receipt_id=P06_1&cargo_delivery_id=RW100
def cy_status_empty_tn_print_xl():
    if request.method == 'GET':
        print_date = request.args.get('print_date', type=str)
        consignor = request.args.get('consignor', type=str)
        consignee = request.args.get('consignee', type=str)
        cargo_name_1 = request.args.get('cargo_name_1', type=str)
        cargo_name_2 = request.args.get('cargo_name_2', type=str)
        cargo_reception = request.args.get('cargo_reception', type=str)
        cargo_delivery = request.args.get('cargo_delivery', type=str)
        driver_name = request.args.get('driver_name', type=str)
        power_of_attorney = request.args.get('power_of_attorney', type=str)
        carrier_name = request.args.get('carrier_name', type=str)
        truck_brand = request.args.get('truck_brand', type=str)
        truck_reg_no = request.args.get('truck_reg_no', type=str)
        printer_no = request.args.get('printer_no', type=str)

        answer = cy_status_tn_print.write_empty_tn_print(
            print_date=print_date,
            consignor=consignor,
            consignee=consignee,
            cargo_name_1=cargo_name_1,
            cargo_name_2=cargo_name_2,
            cargo_reception=cargo_reception,
            cargo_delivery=cargo_delivery,
            driver_name=driver_name,
            power_of_attorney=power_of_attorney,
            carrier_name=carrier_name,
            truck_brand=truck_brand,
            truck_reg_no=truck_reg_no,
            printer_no=printer_no
        )

        return render_template('report_tn_print.html', title='', answer=answer)


@app.route('/excel_reports/trip_ticket', methods=['GET'])
def check_trip_ticket_xl():
    if request.method == 'GET':
        reg_emp = request.args.get('reg_emp', type=str)
        date = request.args.get('date', type=str)
        printer_no = request.args.get('printer_no', type=int)

        answer = trip_ticket.print_trip_ticket(reg_emp, date, printer_no)
        print('Answer: {}'.format(answer))

        return render_template('report.html', title='', answer=answer)


@app.route('/show_pdf/<int:id>', methods=['GET'])
def show_pdf(id):
    if request.method == 'GET':
        filepath = ''

        if id == 1:
            filepath = '../ftp/reports/pdf/shipping_mark_2020420_14488.pdf'
        elif id == 2:
            filepath = '../ftp/reports/pdf/trip_ticket_2020423_91518.pdf'
        elif id == 3:
            filepath = '../ftp/reports/pdf/cy_status_tn_print_2020422_132651.pdf'
        else:
            return redirect('404.html')
        
        with open(filepath, 'rb') as f:
            file_content = f.read()

        response = make_response(file_content, 200)
        response.headers['Content-type'] = 'application/pdf'
        response.headers['Content-disposition'] = 'inline;'

        return response


@app.route('/excel_reports/lp_tn_seq', methods=['GET'])
def lp_tn_xl():
    if request.method == 'GET':
        pool_seq = request.args.get('pool_seq', type=str)
        printer_no = request.args.get('printer_no', type=int)
        print_date = request.args.get('vDate', type=str)
        vendor = request.args.get('vendorCd', type=str)
        transporter = request.args.get('transpCd', type=str)
        truck = request.args.get('truckCd', type=str)
        driver = request.args.get('driverCd', type=str)
        #tare = request.args.get('tareCd', type=str)
        qty = request.args.get('qty', type=str)
        bqty = request.args.get('bqty', type=str)
        #printer_no = int(printer_no[-2:])
        answer = lp_tn.print_lp_tn(pool_seq, print_date, vendor, transporter, truck, driver, qty, bqty, printer_no)
        print('Answer: {}'.format(answer))

        return render_template('report.html', title='', answer=answer)
