from flask import Flask, render_template, request, send_file
import os
import cohere
import markdown2
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib import colors
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time
from flask import send_from_directory, redirect

from VBA_Documentation import *

COHERE_API_KEY_TEXT = "cfsUYJQURZuHCThMSSUEUWWv3LVdTpcdP7YXH8RS"
os.environ["COHERE_API_KEY"] = COHERE_API_KEY_TEXT

app = Flask(__name__)


files_list = os.listdir('files')

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/open_document', methods=['GET'])
def open_document():
    return send_from_directory(directory=os.getcwd(),
                           path='Documentation.pdf',
                           mimetype='application/pdf')

@app.route('/open_logic', methods=['GET'])
def open_logic():
    return send_from_directory(directory=os.getcwd(),
                           path='Functional_Logic.pdf',
                           mimetype='application/pdf')

@app.route('/open_chart', methods=['GET'])
def open_chart():
    return send_from_directory(directory=os.getcwd(),
                           path='flowchart.html',
                           mimetype='application/html')

@app.route('/open_file', methods=['GET'])
def open_file():
    return send_from_directory(directory=os.getcwd(),
                           path='Modernized_Code.py',
                           mimetype='application/py')

@app.route('/open_txt', methods=['GET'])
def open_txt():
    return send_from_directory(directory=os.getcwd(),
                           path='Optimized_Code.txt',
                           mimetype='application/txt')

@app.route('/document', methods=['GET','POST'])
def document():
  if request.method == 'POST':
        text = request.form.get('textarea')
        auto_documentation(text)
        return redirect('/open_document')
  return render_template('document.html') 

@app.route('/logic', methods=['GET','POST'])
def logic():
    if request.method == 'POST':
        text = request.form.get('textarea')
        functional_logic_extractor(text)
        return redirect('/open_logic')
    return render_template('logic.html')

@app.route('/flowchart', methods=['GET','POST'])
def flowchart():
    if request.method == 'POST':
        text = request.form.get('textarea')
        create_flowchart(text)
        return redirect('/open_chart')
    return render_template('flowchart.html')

@app.route('/modernize', methods=['GET','POST'])
def modernize():
    if request.method == 'POST':
        text = request.form.get('textarea')
        modernize_code(text)
        return redirect('/open_file')
    return render_template('modernize.html')

@app.route('/optimize', methods=['GET','POST'])
def optimize():
    if request.method == 'POST':
        text = request.form.get('textarea')
        optimize_code(text)
        return redirect('/open_txt')
    return render_template('optimize.html')

if __name__ == '__main__':
    app.run(debug=True)
