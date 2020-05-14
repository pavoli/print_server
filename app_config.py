# -*- coding: utf-8 -*-
# ! /usr/local/bin/python3

import os


# GCS_DB_PARAMETER = '*****'
GCS_DB_PARAMETER = '*****'

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
# excel paths
EXCEL_TEMPLATE_PATH = os.path.join(PROJECT_ROOT, 'ftp', 'templates', 'excel', '{}.xlsx')
EXCEL_REPORT_PATH = os.path.join(PROJECT_ROOT, 'ftp', 'reports', 'excel', '{}.xlsx')
EXCEL_FOLDER_PATH = os.path.join(PROJECT_ROOT, 'ftp', 'reports', 'excel')
# pdf paths
PDF_REPORT_PATH = os.path.join(PROJECT_ROOT, 'ftp', 'reports', 'pdf', '{}.pdf')
# barcode path
BARCODE_FOLDER_ROOT_PATH = os.path.join(PROJECT_ROOT, 'ftp', 'reports', 'barcode', '{}')
# images path
EXCEL_IMAGE_FOLDER_ROOT_PATH = os.path.join(PROJECT_ROOT, 'ftp', 'templates', 'image', '{}')
