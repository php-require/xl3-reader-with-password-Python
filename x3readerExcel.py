import os
from flask import Flask, flash, request, redirect, jsonify
import re
import pandas as pd
import numpy as np
# объясняется ниже
from werkzeug.utils import secure_filename

# папка для сохранения загруженных файлов
UPLOAD_FOLDER = './uploads'
# расширения файлов, которые разрешено загружать
ALLOWED_EXTENSIONS = {'xlsx'}

# создаем экземпляр приложения
app = Flask(__name__)
# конфигурируем
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# проверка колонок excel
def strip_character(dataCol):
    #r = re.compile(r'[^a-zA-Z !@#$%&*_+-=|\:";<>,./()[\]{}\']')
    r = re.compile(r'[^0-9!@#$%&*_+=|\:";<>,/()[\]{}\']')
    return r.sub('', dataCol)

def parse_excel(name):
    # читаем файл
    df = pd.read_excel(name)
    #df = pd.read_excel("nopass.xlsx")    
    df = df.set_index("№")

    # ищем ошибки
    fio = (df["Ф.И.О."].apply(strip_character).str.len() > 0) | (df["Ф.И.О."].isna())
    sex = ~df["Пол"].str.upper().isin(["М", "Ж"])
    birth, material = [
        pd.to_datetime(df[col], errors="coerce").isna()
        for col in ["Дата рождения", "Дата отбора материала"]
    ]
    sick = (
        (pd.to_datetime(df["Дата заболевания"], errors="coerce").isna())
        & (df["Дата заболевания"].notna())
    )
    errors = pd.DataFrame([fio, sex, birth, material, sick]).T
    
    # список записей с ошибками
    keys = df[np.any([fio, sex, birth, material, sick], 0)].index
    # сохраняем всё в dictionary
    error_dict = {}
    for k in keys:
        error_dict[int(k)] = {col: df.loc[k, col] if errors.loc[k, col] else None for col in errors.columns}
    
    return error_dict
 
def allowed_file(filename):
    """ Функция проверки расширения файла """
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # проверим, передается ли в запросе файл 
        if 'file' not in request.files:
            # После перенаправления на страницу загрузки
            # покажем сообщение пользователю 
            flash('Не могу прочитать файл')
            return redirect(request.url)
        file = request.files['file']
        # Если файл не выбран, то браузер может
        # отправить пустой файл без имени.
        if file.filename == '':
            flash('Нет выбранного файла')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            # безопасно извлекаем оригинальное имя файла
            filename = secure_filename(file.filename)
            # сохраняем файл
            print(os.path.abspath(app.config['UPLOAD_FOLDER']))
            file.save(os.path.join(os.getcwd(), app.config['UPLOAD_FOLDER'], filename))
            # если все прошло успешно, проверяем файл  
            error_dict = parse_excel(filename)
            #return jsonify(error_dict)
            return str(error_dict)
            #return filename

    return '''
    <!doctype html>
    <title>Загрузить новый файл</title>
    <h1>Загрузить новый файл</h1>
    <form method=post enctype=multipart/form-data>
      <input type=file name=file>
      <input type=submit value=Upload>
    </form>
    </html>
    '''
    
