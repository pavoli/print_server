# -*- coding: utf-8 -*-
# ! /usr/local/bin/python3
__author__ = 'p.olifer'

import cx_Oracle
import os
import app_config as config

REF_CURSOR = cx_Oracle.CURSOR


def create_connection():
    os.environ['NLS_LANG'] = '.AL32UTF8'
    con = cx_Oracle.connect(config.GCS_DB_PARAMETER)
    return con


if __name__ == '__main__':
    # check_battery_list()
    pass
