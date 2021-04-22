from flask import Flask, render_template, request, current_app, send_from_directory, redirect, send_file, session
from werkzeug.utils import secure_filename
import os

import pandas as pd
import openpyxl
import csv

UPLOAD_FOLDER = 'static/csv'
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'static/csv'
app.secret_key = "super super secret key"

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
    #   passing the filename from one session to another
      session['file-name'] = filename
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
    shutil.rmtree('static/img')
    os.mkdir('static/img')
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

@app.route('/plot')
def plot():
    import pandas as pd
    import matplotlib.pyplot as plt
    from matplotlib.figure import Figure
    file_name = session.get('file-name')
    df = pd.read_csv('static/csv/'+file_name)
    df.drop('Seq#', axis=1, inplace=True)
    df.drop('V.THD. %', axis=1, inplace=True)
    df.drop('I.THD. %', axis=1, inplace=True)
    df.drop('P.F.', axis=1, inplace=True)

    ref = []
    for i in range(len(df)):
        ref.append(i)

    test = df['TEST']
    product = df['PRODUCT']
    ac_voltage = df['AC VOLTAGE']
    ac_current = df['AC CURRENT']
    ac_watt = df['AC WATT']
    dc_voltage = df['DC VOLTAGE']
    dc_current = df['DC CURRENT']
    dc_power = df['DC POWER']
    efficiency = df['EFFICIENCY']
    pcb_temp = df['PCB TEMP']
    junc_temp = df['JUNC TEMP']
    housing_temp = df['HOUSING TEMP']
    xmer_temp = df["X'MER TEMP"]
    room_temp = df['ROOM TEMP']
    tap = df['TAP']
    recipe = df['RECIPE']


    plt.plot(ref, test)
    plt.plot(ref, product)
    plt.plot(ref, ac_voltage)
    plt.plot(ref, ac_current)
    plt.plot(ref, ac_watt)
    plt.plot(ref, dc_voltage)
    plt.plot(ref, dc_current)
    plt.plot(ref, dc_power)
    plt.plot(ref, efficiency)
    plt.plot(ref, pcb_temp)
    plt.plot(ref, junc_temp)
    plt.plot(ref, housing_temp)
    plt.plot(ref, xmer_temp)
    plt.plot(ref, room_temp)
    plt.plot(ref, tap)
    plt.plot(ref, recipe)


    plt.legend(['Test', 'Product', 'AC Voltage', 'AC Current', 'AC Watt', 'DC Voltage', 
            'DC Current', 'DC Power', 'Efficiency', 'PCB Temp.', 'Junc Temp', 'Housing  Temp', 'Xmer Temp', 'Room Temp',
            'Tap', 'Recipe'], fontsize="xx-small")
    plt.plot()
    plt.savefig('static/img/plot.png')

    return render_template('plot.html', url='static/img/plot.png')

if __name__ == "__main__":
    app.run(debug=True)
