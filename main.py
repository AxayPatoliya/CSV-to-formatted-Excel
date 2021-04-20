from flask import Flask, render_template, request, current_app, send_from_directory, redirect, send_file
from werkzeug.utils import secure_filename
import os

import pandas as pd
import openpyxl
import csv

UPLOAD_FOLDER = 'static/csv'
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'static/csv'

# rendering the index.html
@app.route("/")
def home():
    return render_template("index.html")

# file uploader and storing at static/csv/
@app.route('/uploader', methods = ['GET', 'POST'])
def upload_file():
    # uploader
   if request.method == 'POST':
      file = request.files['file']
      filename = secure_filename(file.filename)
      file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

      # converting the file into customize form
      destination = openpyxl.load_workbook("fill-3.xlsx")
      sheet_destination = destination.active
    #   os.rename(file %i, r"static/csv/source.csv")
      with open ("static/csv/" + filename, 'r') as f:
          files = csv.reader(f)
          lst = list(files)
      for row in lst:
          sheet_destination.append(row)

      destination.save("final.xlsx")

      #return 'file uploaded successfully'
      return render_template("download.html")

@app.route('/uploads/<path:filename>', methods=['GET', 'POST'])
def download(filename):
    # Appending app path to upload folder path within app root folder
    uploads = os.path.join(current_app.root_path, app.config['UPLOAD_FOLDER'])
    # Returning file from appended path
    return send_from_directory(directory=uploads, filename=filename)

# removes the csv files from static/csv folder
@app.route('/remove_all_csv')
def remove():
    import os
    import shutil
    shutil.rmtree('static/csv')
    os.mkdir('static/csv')
    # os.remove("final_with_header.xlsx")
    return redirect("/")

# removes the excel files from static/csv folder
@app.route("/remove_exl")
def remove_exl():
    os.remove("final.xlsx")
    return redirect("/")

# downlods the customized excel file
@app.route('/download')
def downloadFile():
    path = "final.xlsx"
    return send_file(path, as_attachment=True, cache_timeout=0)

if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0')
