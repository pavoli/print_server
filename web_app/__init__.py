# -*- coding: utf-8 -*-
# ! /usr/local/bin/python3
__author__ = 'p.olifer'


'''

'''

from flask import Flask

app = Flask(__name__)
# web_app.config.from_object('config')
from web_app import views, models
