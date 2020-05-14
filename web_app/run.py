# -*- coding: utf-8 -*-
# ! /usr/local/bin/python3
__author__ = 'p.olifer'


'''

'''

from web_app import app
app.run(host='0.0.0.0', port=5000, debug=True, threaded=True)
