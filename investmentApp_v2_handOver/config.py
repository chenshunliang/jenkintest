# -*- coding: utf-8 -*-
"""
Created on Mon Dec 28 18:49:46 2015

@author: zhaoyi
"""

import os
basedir = os.path.abspath(os.path.dirname(__file__))


class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'hard to guess string'
    SQLALCHEMY_COMMIT_ON_TEARDOWN = True
    SQLALCHEMY_RECORD_QUERIES = True
    FLASKY_COMMENTS_PER_PAGE = 30
    FLASKY_SLOW_DB_QUERY_TIME=0.5
    CACHE_TYPE = "null" #clearing the cache
    FLASKY_PROGRAMS_PER_PAGE=10

    @staticmethod
    def init_app(app):
        pass


class DevelopmentConfig(Config):
    DEBUG = True

config = {
    'default': DevelopmentConfig
}

print 'aaa'
