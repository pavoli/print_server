# -*- coding: utf-8 -*-
# ! /usr/local/bin/python3

__author__ = 'p.olifer'

import win32print
import win32api
import os
import time
import app_config


def printers_list():
    printers = win32print.EnumPrinters(5)
    for p in printers:
        print(p[2])


def set_default_printer(printer_no=15):
    win32print.SetDefaultPrinter(app_config.PRINTER_NAMES[printer_no])


def get_default_printer():
    printer_default_name = win32print.GetDefaultPrinter()
    print(printer_default_name)


# 1 = no flip, 2 = flip up, 3 = flip over
def set_duplex(printer_no, duplex):
    print_defaults = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
    handle = win32print.OpenPrinter(app_config.PRINTER_NAMES[printer_no], print_defaults)
    level = 2
    attributes = win32print.GetPrinter(handle, level)
    attributes['pDevMode'].Duplex = duplex
    try:
        win32print.SetPrinter(handle, level, attributes, 0)
    except Exception as e:
        print('Failed to set duplex')
        print(e)
    finally:
        win32print.ClosePrinter(handle)


def print_excel(printer_no, path_to_file):
    printer_name = app_config.PRINTER_NAMES[printer_no]

    win32api.ShellExecute(
        1,
        'printto',
        path_to_file,
        '{}'.format(printer_name),
        '.',
        0
    )


def print_excel_file(printer_name, path_to_file):
    win32api.ShellExecute(
        1,
        'printto',
        path_to_file,
        '{}'.format(printer_name),
        '.',
        0
    )


def delete_file(path_to_file, try_count=1):
    if os.path.exists(path=path_to_file):
        file_name = path_to_file.split('\\')[-1]
        try:
            os.remove(path_to_file)
            print('File {} deleted!'.format(file_name))
        except PermissionError:
            print('Can not delete file {}'.format(file_name))
            if try_count < 60:
                time.sleep(1.0)
                try_count += 1
                delete_file(path_to_file, try_count)


if __name__ == '__main__':
    # delete_file('E:\\git\\toolApps\\print_service\\ftp\\templates\\excel\\fuel_control_report_1.xlsx')
    pass
