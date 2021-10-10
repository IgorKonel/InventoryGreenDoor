from tkinter import *

import variables


def get_window():

    def click():
        """
        Команды для обработки нажатия кнопки 'Принять'
        :return: None
        """
        # Label о получении ссылки
        lbl_success = Label(wnd, text="Ссылка получена! Нажмите Закрыть или введите новую ссылку",
                            font=("Arial Bold", 12))
        lbl_success.place(x=100, y=50)

        variables.SHEET_ID = txt.get()
        # Обновить окно
        wnd.update()

    variables.SHEET_ID = ""

    # Создание окна
    wnd = Tk()

    # Заголовок окна
    wnd.title("Инвентура")

    # Геометрические размеры окна
    wnd.geometry("750x220")

    # Окно постоянных размеров
    wnd.resizable(False, False)

    # Шрифты
    fnt_2 = ("Arial", 9, "italic")

    # Создание объекта метки:
    lbl = Label(wnd, text="Введите ссылку на файл Google Sheets этого месяца:", font=("Arial Bold", 19))

    # Размещение метки в окне:
    lbl.place(x=65, y=10)

    # Создание объекта поля вода:
    txt = Entry(master=wnd, width=90)

    # Шрифт для поля ввода:
    txt.configure(font=fnt_2)

    txt.place(x=65, y=120)

    # Объект кнопки:
    btn = Button(wnd, text="Принять", font=("Courier New Bold", 13), command=click)

    # Размещение объекта кнопки в окне:
    btn.place(x=310, y=160, width=160, height=30)

    # Кнопка закрытия
    btn_close = Button(wnd, text="Закрыть", font=("Courier New Bold", 13), command=wnd.destroy)

    # Расположение и размер кнопки Закрыть
    btn_close.place(x=500, y=160, width=170, height=30)

    # Показать окно на экране
    wnd.mainloop()
