import imp
from src.handlers.handler import hello
import json

def docGen(event, context):
    hello(event, context)
