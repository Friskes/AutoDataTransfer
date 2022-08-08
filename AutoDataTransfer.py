import openpyxl
from tkinter import Tk, filedialog
from SpinCursor import SpinCursor
from string import punctuation

spin = SpinCursor(msg="Ожидайте выполнения программы",
                  del_msg_after_stop=True,
                  maxspin=500,
                  minspin=0,
                  speed=2,
                  animType='sticks')

def get_file_names():
    print('AutoDataTransfer by Friskes. v1.4\n')
    print('Сначало выберите путь к файлу бурения.')
    print('Затем выберите путь к файлу сводки.\n')
    input('Нажмите ENTER для запуска программы.\n')

    window = Tk()
    screenwidth = window.winfo_screenwidth() // 2 - 540 # влево вправо
    screenheight = window.winfo_screenheight() // 2 - 330 # вверх вниз
    window.geometry('0x0+{}+{}'.format(screenwidth, screenheight))

    filePath = filedialog.askopenfilename(filetypes=(("Text files", "*.xlsx"), ("all files", "*.*")), title='Выберите путь к файлу бурения.')
    filePath2 = filedialog.askopenfilename(filetypes=(("Text files", "*.xlsx"), ("all files", "*.*")), title='Выберите путь к файлу сводки.')

    window.destroy()

    if filePath == '' or filePath2 == '':
        print('Не правильно выбран путь.\n')
        input('Нажмите ENTER для закрытия программы.\n')
        return None, None

    elif filePath == filePath2:
        print('Вы выбрали два одинаковых пути.\n')
        input('Нажмите ENTER для закрытия программы.\n')
        return None, None

    elif 'бурение' in filePath2.lower() and 'сводка' in filePath.lower():
        print('Вы перепутали пути местами.\n')
        input('Нажмите ENTER для закрытия программы.\n')
        return None, None

    print(f'Путь к файлу бурения:\n{filePath}')
    print(f'\nПуть к файлу сводки:\n{filePath2}\n')

    return filePath, filePath2


def parse_and_make_dict(filePath):
    spin.start()

    book = openpyxl.open(filePath, read_only=True)
    sheet = book.worksheets[0]

    zaboi = {}
    sostoyanie = {}

    for row in range(1, sheet.max_row + 1):
    # for row in range(1, 150):

        s3 = sheet[row][3-1].value
        # получили индексы строк с именами в столбце 'мастер (бригада)'
        if s3 != None and s3.isalpha() and s3.isupper():
            # print(row, s3)

            s2 = sheet[row][2-1].value
            # проверка на наличие цифр в столбце '№ п/п'
            if s2 != None and str(s2).isdigit():
                # print(row, s2)

                s5 = sheet[row][5-1].value
                s6 = sheet[row][6-1].value
                # проверяем содержится ли информация в столбцах 'М/Р' и '№ КУСТ'
                # если содержится то запоминаем информацию из этих строк в качестве ключей
                if s5 != None and s6 != None:
                    # print(row, s3, s5, s6)

                    s25 = sheet[row][25-1].value
                    s27 = sheet[row][27-1].value

                    # в качестве ключей используем склееные строки 'М/Р' и '№ КУСТ'
                    zaboi[str(s5) + str (s6)] = s25
                    sostoyanie[str(s5) + str (s6)] = s27

                    # print('zaboi:', str(s5) + str (s6), s25)
                    # print('sostoyanie:', str(s5) + str (s6), s27)
    spin.stop()
    spin.join() # подождём полного завершения работы модуля анимации

    return zaboi, sostoyanie

# print('zaboi:', zaboi)
# print('sostoyanie:', sostoyanie)


def write_xlsx_file(filePath2, zaboi, sostoyanie):
    print('—'*99)
    print()

    book = openpyxl.open(filePath2)
    sheet = book.worksheets[0]

    for row in range(1, sheet.max_row + 1):
    # for row in range(1, 150):

        s4 = sheet[row][4-1].value
        s5 = sheet[row][5-1].value

        if s4 != None and s5 != None:

            sostoyanie_key = str(s4).lower() + str(s5).lower()
            sostoyanie_key = sostoyanie_key.translate(str.maketrans('', '', punctuation + ' '))

            for k, val in zaboi.items():

                k = k.lower()
                k = k.translate(str.maketrans('', '', punctuation + ' '))

                if k == sostoyanie_key:
                    # print(k, str(s4) + str(s5))
                    if val != sheet[row][13-1].value:
                        print(f'По ключу "{k}" в текущий забой записываю "{val}" вместо "{sheet[row][13-1].value}"')
                        sheet[row][13-1].value = val

            for k, val in sostoyanie.items():

                k = k.lower()
                k = k.translate(str.maketrans('', '', punctuation + ' '))

                if k == sostoyanie_key:
                    # print(k, str(s4) + str(s5))
                    if val != sheet[row][14-1].value:
                        print(f'По ключу "{k}" в состояние записываю "{val}" вместо "{sheet[row][14-1].value}"')
                        print()
                        sheet[row][14-1].value = val

    book.save(filePath2)
    book.close()

    print('—'*99)
    print('\nВыполнение программы завершено.\n')
    input('Нажмите ENTER для закрытия программы.\n')


def main():
    filePath, filePath2 = get_file_names()
    if filePath == None or filePath2 == None:
        return
    zaboi, sostoyanie = parse_and_make_dict(filePath)
    write_xlsx_file(filePath2, zaboi, sostoyanie)

if __name__ == '__main__':
    main()
