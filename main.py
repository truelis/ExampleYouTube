from flask import Flask, request, redirect, send_file
import pandas as pd
import io

app = Flask(__name__)

@app.route("/")
def hello():
    name = request.args.get('name', 'World')
    return 'Hello' + name + '!'
