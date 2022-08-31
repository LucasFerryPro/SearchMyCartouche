import os
import glob
from pandas import *
from flask import Flask, flash, request, redirect, url_for, render_template
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = '/home/SearchMyCartouche/SearchMyCartouche/static/excel/'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['SECRET_KEY'] = 'some random string'
app.config['SESSION_TYPE'] = 'filesystem'

listeFinal = []
keys=[]
pwd = 'InkBI123emPT'

canHome = False
canUpload = False

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def dataSorter(xls):
    data = xls.parse(xls.sheet_names[0])
    liste = list(data.to_dict().values())
    listeFinal = []
    listeInter = {}

    for i in range(2, len(liste[0])):
        for j in range(len(liste)):
            listeInter[liste[j][1]] = liste[j][i]
        listeFinal.append(listeInter)
        listeInter = {}

    return listeFinal, list(listeFinal[0].keys())

@app.route('/', methods=['GET','POST'])
def password():
    global canHome
    global canUpload
    canHome = False
    canUpload = False
    if request.method == 'POST':
        if request.form.get('password') == pwd:
            canHome = True
            canUpload = True
            try:
                return redirect(url_for('home'))
            except:
                return redirect(url_for('upload_file'))
    return render_template('password.html')

@app.route('/home', methods=['GET','POST'])
def home():
    if not canHome:
        return redirect(url_for('password'))
    try:
        find = []
        filename = glob.glob('/home/SearchMyCartouche/SearchMyCartouche/static/excel/*')[0][55:]
        listeFinal = dataSorter(ExcelFile(os.path.join(app.config['UPLOAD_FOLDER'], filename)))[0]
        keys = dataSorter(ExcelFile(os.path.join(app.config['UPLOAD_FOLDER'], filename)))[1]
        if request.method == 'POST':
            for element in listeFinal:
                if request.form.get("texte") in element[keys[1]]:
                    find.append(listeFinal.index(element))
            for index in find:
                inter = listeFinal[index]
                del listeFinal[index]
                listeFinal.insert(0,inter)
        return render_template('home.html', listeFinal=listeFinal, keys=keys, find=find)
    except:
        return redirect(url_for('upload_file'))



@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if not canUpload:
        return redirect(url_for('password'))
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            try:
                return redirect(url_for('home'))
            except:
                return redirect(url_for('upload_file'))
        if file and allowed_file(file.filename):
            files = glob.glob('/home/SearchMyCartouche/SearchMyCartouche/static/excel/*')
            for f in files: os.remove(f)
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            return redirect(url_for('home'))
    return render_template('upload.html')


if __name__ == '__main__':


    app.run()
