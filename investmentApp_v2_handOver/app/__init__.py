# -*- coding: utf-8 -*-
"""
Created on Mon Dec 28 18:48:07 2015

@author: zhaoyi
"""

# -*- coding: utf-8 -*-
"""
Created on Tue Dec  8 19:04:47 2015

@author: zhaoyi
"""

from flask import Flask
from flask.ext.bootstrap import Bootstrap
from flask.ext.mail import Mail
from flask.ext.moment import Moment
from flask.ext.sqlalchemy import SQLAlchemy
from flask.ext.login import LoginManager
from flask.ext.pagedown import PageDown
from flask.ext.cache import Cache
from config import config

bootstrap = Bootstrap()
mail = Mail()
moment = Moment()
db = SQLAlchemy()
pagedown = PageDown()
cache = Cache()


def create_app(config_name):
    app = Flask(__name__)
    app.config.from_object(config[config_name])
    config[config_name].init_app(app)

    bootstrap.init_app(app)
    mail.init_app(app)
    moment.init_app(app)
    db.init_app(app)
    pagedown.init_app(app)
    cache.init_app(app)

    from .calculator import calculator as calculator_blueprint
    app.register_blueprint(calculator_blueprint)

    return app
