# -*- coding: utf-8 -*-
"""
Created on Mon Dec 28 18:46:14 2015

@author: zhaoyi
"""

import os
from app import create_app

os.chdir('/Users/chenshunliang/wepiaocode/home/zhaoyi/python3/investmentApp_v2_handOver')
app = create_app(os.getenv('FLASK_CONFIG') or 'default')

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=8080, debug=True)
