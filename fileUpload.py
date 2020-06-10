# -*- coding: utf-8 -*-
"""
tiny_example.py
:copyright: (c) 2015 by C. W.
:license: GPL v3 or BSD
"""
from flask import Flask, request, jsonify
import flask_excel as excel
from openpyxl import load_workbook
import pandas as pd

app = Flask(__name__)

@app.route("/upload", methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        f = request.files['file']
#        return (f.filename)
        return process(f.filename)
#        data_xls = pd.read_excel(f)
#        return data_xls.to_html()
    return '''
    <!doctype html>
    <title>Upload an excel file</title>
    <h1>Excel file upload (csv, tsv, csvz, tsvz only)</h1>
    <form action="" method=post enctype=multipart/form-data><p>
    <input type=file name=file><input type=submit value=Upload>
    </form>
    '''
def process(filename):

    routename = ['ZYAA', 'ZYBB', 'ZYCC']
    supervisors = ['X', 'Y', 'Z']
    workbook = load_workbook(filename)
    worksheet = workbook.active
    worksheet.column_dimensions.group('A', 'B', hidden=True)
    routes = worksheet.columns[16]
    i = 16
    worksheet['D1'] = 'Supervisor'
    for route in routes:
        if route.value in routename:
            pos = routes.index(route)
            worksheet['D' + str(i)].value = supervisors[pos]
            print (route.value)
            i += 1

    workbook.save(filename)

@app.route("/download", methods=['GET'])
def download_file():
    return excel.make_response_from_array([[1, 2], [3, 4]], "csv")


@app.route("/export", methods=['GET'])
def export_records():
    return excel.make_response_from_array([[1, 2], [3, 4]], "csv",
                                          file_name="export_data")


@app.route("/download_file_named_in_unicode", methods=['GET'])
def download_file_named_in_unicode():
    return excel.make_response_from_array([[1, 2], [3, 4]], "csv",
                                          file_name=u"中文文件名")


# insert database related code here
if __name__ == "__main__":
    excel.init_excel(app)
    app.run()