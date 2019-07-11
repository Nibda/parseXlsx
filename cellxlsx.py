# Парсинг та масова заміна даних в таблиці Excel
# Джерела
# https://habr.com/company/otus/blog/331998/
# https://openpyxl.readthedocs.io/en/stable/usage.html
# http://qaru.site/questions/11190/find-all-files-in-a-directory-with-extension-txt-in-python

import os
from openpyxl import load_workbook


def renew(fop, text, cell):
    for root, dirs, files in os.walk("d:\\python\\Nakladni\\{}\\".format(fop)):
        for file in files:
            if file.endswith(".xlsx"):
                print(os.path.join(root, file))
                # Відкриваємо файл
                wb = load_workbook(os.path.join(root, file))
                #  читаємо таблицю файлу
                sheet = wb['Sheet1']
                print(sheet[cell].value)
                # робимо заміну тексту в ячейкі
                sheet[cell] = text
                print(sheet[cell].value)
                # записуємо назад у файл
                wb.save(os.path.join(root, file))


if __name__ == '__main__':
    renew('kirch', 'Договір комісії №19-02/2 від 19.02.18р.', 'c11')
    renew('vengl', 'Договір комісії №19-02/1 від 19.02.18р.', 'c11')
    renew('kirch_serp\\', 'Договір комісії №19-02/2 від 19.02.18р.', 'c12')
    renew('vengl_serp\\', 'Договір комісії №19-02/1 від 19.02.18р.', 'c12')
