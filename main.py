import pandas as pd
import pymorphy3
from tkinter import *
from tkinter import filedialog as fd

morph = pymorphy3.MorphAnalyzer()

marks = '''!()-[]{};?@#$%:'"\,./^&amp;*_'''


def select_file():
    file_name = str(fd.askopenfilename())
    dir = file_name[0:file_name.rfind('/') + 1]
    FILE_ = file_name
    file_name_resul = dir + 'result.xlsx'
    l1['text'] = "Берем данные из файла:"
    l2['text'] = file_name
    b1['state'] = 'disabled'
    l3['text'] = "Программа выполняется..."
    root.update()
    excel_data = pd.read_excel(FILE_, usecols=[1])

    col_B = []
    col_C = []
    col_D = []
    col_E = []
    col_F = []

    for item in excel_data[:].values:
        value_NOUN = None
        value_ADJF = None
        value_VERB = None
        line_str = str(item[0]).strip()
        col_B.append(line_str)
        single_row = line_str.split()
        print(single_row)

        for word in single_row:
            for x in word:
                if x in marks:
                    word = word.replace(x, '')
            part_of_speech = morph.parse(word)[0].tag.POS
            print(part_of_speech)

            if part_of_speech == 'NOUN':  # существительное
                value_NOUN = 1
            elif part_of_speech in ['ADJF', 'ADJS']:  # прилагательное
                value_ADJF = 1
            elif part_of_speech in ['VERB', 'INFN']:  # глагол
                value_VERB = 1

        if value_NOUN == 1:
            col_D.append(1)
        else:
            col_D.append('')

        if value_ADJF == 1:
            col_E.append(1)
        else:
            col_E.append('')

        if value_VERB == 1:
            col_F.append(1)
        else:
            col_F.append('')

        if value_NOUN is None and value_ADJF is None and value_VERB is None:
            col_C.append('нд')
        else:
            if value_NOUN == 1 and value_ADJF is None and value_VERB is None:
                col_C.append('существительное')
            elif value_NOUN is None and value_ADJF == 1 and value_VERB is None:
                col_C.append('прилагательное')
            elif value_NOUN is None and value_ADJF is None and value_VERB == 1:
                col_C.append('глагол')
            else:
                col_C.append('')

    df = pd.DataFrame({
        'Исходный текст': col_B,
        'Тип (глагол/существительно/прилагательное)': col_C,
        'Существительное': col_D,
        'прилагательное': col_E,
        'глагол': col_F,
    })

    df.to_excel(file_name_resul, sheet_name='дата', index=False)

    l3['text'] = "Готово! Результат здесь:"
    l4['text'] = file_name_resul
    b1['state'] = 'normal'


if __name__ == '__main__':
    root = Tk()
    root.title('Парсер граммем')
    root.geometry('500x150')
    root.resizable(width=False, height=False)

    b1 = Button(text="Открыть", command=select_file)
    l1 = Label(root, width=450)
    l2 = Label(root, text="", width=450)
    l3 = Label(root, text="", width=450)
    l4 = Label(root, text="", width=450)

    l0 = Label(root, text="Выберите файл")
    l0.pack()
    b1.pack()
    l1.pack()
    l2.pack()
    l3.pack()
    l4.pack()

    root.mainloop()
