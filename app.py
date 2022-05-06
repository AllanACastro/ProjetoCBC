import os
import time

from flask import Flask, render_template, request, send_from_directory
from projeto import *


app = Flask(__name__)

@app.route("/", methods=['GET', 'POST']) 
def send():

    if request.method == 'POST':
        AU = request.form['AU']
        f = request.files['file']
        f.save(f.filename)

        organizaTabela(f.filename, AU)
        os.remove(f.filename)
        return render_template('success.html')

    return render_template('main.html')



@app.route("/success")
def getPlotCSV():
    name_arquivo = os.listdir("Arquivos")[0]
    diretorio = 'Arquivos'
    return send_from_directory(diretorio, name_arquivo, as_attachment=True)




if __name__ == "__main__":
    app.run(debug= True)