import sys
import re
import json
import numpy as np
import pandas as pd
import xlwings as xw

def strip_character(dataCol):
    #r = re.compile(r'[^a-zA-Z !@#$%&*_+-=|\:";<>,./()[\]{}\']')
    r = re.compile(r'[^0-9!@#$%&*_+=|\:";<>,/()[\]{}\']')
    return r.sub('', dataCol)

if __name__ == "__main__":
	# читаем файл
	name, password = sys.argv[1:]
	#wb = xw.Book("asdasd_04.10.2021.xlsx", password="i2dnby2020111")
	wb = xw.Book(name, password=password)
	sheet = wb.sheets[0]
	df = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
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

	# конвертируем в json строку
	error_json = json.dumps(error_dict)
	print(error_json)
