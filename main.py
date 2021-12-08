import sqlite3
import openpyxl


def main():
    book = openpyxl.load_workbook(filename="one.xlsx")  # open xlsx file
    sheet = book.worksheets[0]
    counter_postamat = 0
    db = sqlite3.connect('map.db')  # open data_base
    cursor = db.cursor()

    for row in range(7, sheet.max_row):
        city = sheet[row][5].value
        print(row)
        for value_db in cursor.execute("SELECT * FROM prices"):
            if (value_db[0] in city) == True:
                if sheet[row][4].value != "Постамат":
                    counter_postamat = 0
                    price_release(sheet[row], value_db, 0)
                elif ("Итог" in sheet[row][5].value) == True:
                    if counter_postamat <= 10:
                        price_release(sheet[row], value_db, 0)
                        sheet[row][14].value = 200
                        counter_postamat = 0
                    else:
                        counter_postamat -= 10
                        price_release(sheet[row], value_db, 0)
                        sheet[row][14].value =200 + counter_postamat * 10
                        counter_postamat = 0
                else:
                    counter_postamat += 1
                if sheet[row][11].value is not None:
                    if sheet[row][12].value == "Карта":
                        sheet[row][15].value = round(float(sheet[row][11].value) / 100 * 1.75,2)
                    else:
                        sheet[row][15].value = round(float(sheet[row][11].value) / 100 * 0.5,2)
    book.save("one.xlsx")


def price_release(sheet, value_db, price):

    size_pos_rec = float('{:.2f}'.format(float(sheet[9].value)))  # 14.129999 fix to 14.13
    size_pos_real = float('{:.2f}'.format(float(sheet[10].value)))
    max_size = max(size_pos_rec, size_pos_real)  # max size in file
    if max_size <= 5:
        price += value_db[1]
    elif max_size <= 20:
        price += value_db[2]
    else:
        price += value_db[2]
        while(max_size - 20 > 0):
            price += value_db[3]
            max_size -= 1
    sheet[13].value = round(price)
        #print(sheet[0].value, max_size, value_db,price)


if __name__ == '__main__':
    main()
