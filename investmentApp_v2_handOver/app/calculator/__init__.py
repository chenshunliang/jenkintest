from flask import Blueprint

calculator = Blueprint('calculator', __name__)

from . import views
