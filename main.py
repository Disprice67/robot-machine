from math import ceil
from os import remove
from time import sleep

from config import *
from db import Database
from ebay import Ebay
from excel import Excel, File, Writer
from mail import Email, InfoMessage

"""
In the main.py file, the main logic of program interaction is grouped, 
working with an SQL query, processing and collecting information and more.
-----
* First, the main constants are presented.

* method category() determines the category of party member

* method economy() performs economic calculation of labor costs

* method find() trying to find the part number in the archive
to remember his last line and contract

* method search() the main method that forms queries to the database, 
queries can be viewed in .env

* method add() adds the collected info to a certain pair number in the dictionary

* method compilate() collects the main method 
for the formation of information for the part number
"""

MESSAGE = ['Хотим купить под: {atr}. По цене: {pricegrabe}$', 
           'P/N совпал без букв: {part}',
           'Шасси! БП - {part_bp}, FAN - {part_fan}, Комментарий - {comment}']

CATEGORY_EXCEPTION = ('LIC-1', 'SOFT-1', 'MSCL')

SEARCH = [('Свод', 'ЗИП'), ('Закупка1', 'ЗИП'), 
          ('Закупка2', 'ЗИП'), ('Архив', '№ЗАПРОСА'), 
          ('АрхивСтр', '№ ПОСЛЕДНЕГО ЗАПРОСА ИЗ АРХИВА'), ('Шасси', 'PN')]

# sql_request
SQL_REQUEST = {
                  'Архив': [SQL_ARCHIVE, SQL_ARCHIVE_IN, SQL_ARCHIVE_OUT],
                  'Коллизии': SQL_COLLISION,
                  'Категория': SQL_CATEGORY,
                  'Свод': [SQL_CORPUSE, SQL_CORPUSE_IN, SQL_CORPUSE_OUT],
                  'Закупка1': [SQL_PURCHASE, SQL_PURCHASE_IN, SQL_PURCHASE_OUT],
                  'Закупка2': [SQL_PURCHASE_TWO, SQL_PURCHASE_TWO_IN, SQL_PURCHASE_TWO_OUT],
                  'Шасси': SQL_SHASSIS_IN,
                  'АрхивСтр': ARCHIVE
    }


# const_dir
SUP_DIR = ['Статусы', 'Закупка1', 'Закупка2', 
           'Категоризатор1', 'Категоризатор2', 
           'Коллизии', 'Свод', 'Архив', 'Шасси']


# const_column_value
PARAM = [{
            '№ ЗАПРОСА': 'TEXT, ',
            'СТАТУС': 'TEXT',
    },
            {
            'АРТИКУЛ': 'TEXT, ',
            'СУММА ПРОДАЖИ': 'REAL NOT NULL, ',
            'КЛИЕНТ': 'TEXT, ',
            'НАЗНАЧЕНИЕ': 'TEXT, ',
            'MAIN_KEY': 'TEXT, ',
            'PK': 'ID INTEGER PRIMARY KEY'
    },
            {
            'PN': 'TEXT, ',
            'КЛИЕНТЫ': 'TEXT, ',
            'ЗАКУПАЕМ ПОД ЗАКАЗЧИКА': 'TEXT, ',
            'СУММА СОВМЕСТНОЙ ЗАКУПКИ': 'REAL, ',
            'МАГАЗИН': 'TEXT, ',
            'ОЦЕНОЧНАЯ СТОИМОСТЬ': 'TEXT, ',
            'MAIN_KEY': 'TEXT, ',
            'PK': 'ID INTEGER PRIMARY KEY'
    },
            {
            'КАТЕГОРИЯ': 'TEXT, ',
            'ТЗ': 'REAL, ',
            'РЕМОНТЫ': 'INTEGER'
    },
            {
            'МОДЕЛЬ НАЧИНАЕТСЯ С…': 'TEXT, ',
            'КАТЕГОРИЯ СЛОЖНОСТИ ТЗ': 'TEXT, ',
            'РЕМОНТ': 'INTEGER, ',
            'ТРУДОЗАТРАТЫ': 'REAL, ',
            'MAIN_KEY': 'TEXT, ',
            'PK': 'ID INTEGER PRIMARY KEY'
    },
            {
            'ОПИСАНИЕ ВКЛЮЧАЕТ': 'TEXT, ',
            'КАТЕГОРИЯ СЛОЖНОСТИ ТЗ': 'TEXT, ',
            'РЕМОНТ': 'INTEGER, ',
            'ТРУДОЗАТРАТЫ': 'REAL'

    },
            {
            'PART #': 'TEXT, ',
            'НАЗНАЧЕНИЕ': 'TEXT, ',
            'ЛОГИЧЕСКИЙ УЧЕТ': 'TEXT, ',
            'MAIN_KEY': 'TEXT, ',
            'PK': 'ID INTEGER PRIMARY KEY'
    },
            {
            'PN': 'TEXT, ',
            'СТОИМОСТЬ ЗАКУПКИ ЗИП': 'TEXT, ',
            'ЗИП': 'TEXT, ',
            'ДТК СЕРВИС': 'TEXT, ',
            'НАЗНАЧЕНИЕ': 'TEXT, ',
            'КОЛ-ВО': 'INTEGER, ',
            '№ ЗАПРОСА': 'TEXT, ',
            'MAIN_KEY': 'TEXT, ',
            'PK': 'ID INTEGER PRIMARY KEY'
    },
            {
            'PN': 'TEXT, ',
            'БП': 'TEXT, ',
            'FAN': 'TEXT, ',
            'КОММЕНТАРИИ': 'TEXT, ',
            'PK': 'ID INTEGER PRIMARY KEY'
    }]

SUBJECT_RULE = (
        'ОБНОВЛЕНИЕ ШАБЛОНА'
    )

SAMPLE_XLSX = {
        'ЗАКАЗЧИК': 'TEXT',
        'ВЕНДОР': 'TEXT',
        'P/N': 'TEXT',
        'ОПИСАНИЕ': 'TEXT',
        'КОЛИЧЕСТВО': 'INTEGER'
    }

FIND_KEY = None


def category(db: Database, data, amount, 
             excel: Excel, main_key, comment):
    """Add_category."""

    value = ''
    if comment:
        # category_collision
        filter_comment = excel.filterkey(comment.upper(), 'PN')
        collision = db.takeinfo(SQL_REQUEST['Коллизии'].format(comment=f"{filter_comment}"))
        if collision:

            value = collision[0]['КАТЕГОРИЯ СЛОЖНОСТИ ТЗ']
            if value:
                repair = collision[0]['РЕМОНТ']
                work_time = collision[0]['ТРУДОЗАТРАТЫ']

    if not value:
        # category_normal
        category = db.takeinfo(SQL_REQUEST['Категория'].format(part=f"{excel.filterkey(main_key, 'PN')}"))
        if category:
            value = category[0]['КАТЕГОРИЯ СЛОЖНОСТИ ТЗ']
        
            if value:
                repair = category[0]['РЕМОНТ']
                work_time = category[0]['ТРУДОЗАТРАТЫ']
        
    if not value:
        # none_result
        value = 'None'
        repair = 6001
        work_time = 4
    
    if repair == 'None':
        repair = 0

    if work_time == 'None':
        work_time = 0  

    data[main_key]['КАТЕГОРИЯ'] = value
    economy(repair, work_time, amount, data, main_key)

def economy(repair_one, work_time_one, amount, data, main_key):
    """Calculation_of_repairs_and_abor_costs."""

    low_factor_archive = 1
    low_factor_amount = 1
    one_unit_repair = 0
    one_unit_work = 0
    work = 0
    repair = 0
    try:
        if data[main_key]['QTY ИЗ АРХИВОВ'] > 100:
            low_factor_archive = 0.75
    except:
        pass

    if amount > 20:
        low_factor_amount = 0.5
    elif amount > 10:
        low_factor_amount = 0.2
    
    low_factor = low_factor_archive * low_factor_amount

    if amount >= 11 and amount <= 20:
        repair = 10 * repair_one
    elif amount < 11:
        repair = amount * repair_one
    elif amount > 20:
        repair = amount * repair_one * low_factor

    if amount < 10:
        work = amount * work_time_one
    elif amount > 10:
        work = amount * work_time_one * low_factor

    if amount:
        one_unit_repair = ceil(repair / amount)
        one_unit_work = ceil(work / amount)

    data[main_key]['РЕМОНТЫ ЗА 1ЕД/РУБ'] = one_unit_repair
    data[main_key]['ТРУДОЗАТРАТЫ ЗА 1ЕД/HOURS'] = one_unit_work
    data[main_key]['РУБ, СТОИМОСТЬ ПОДДЕРЖКИ'] = ceil(repair)
    data[main_key]['HOURS'] = ceil(work)

def find(db: Database, key, data, excel: Excel, atr: tuple):
    """Search_archive."""
    name, col = atr
    request = db.takeinfo(SQL_REQUEST[name].format(part=f"{excel.filterkey(key, 'PN')}"))
    if request:
        value = request[0][col]
        if value:
            if name == 'АрхивСтр':
                number_str = request[0]['№ ПОСЛЕДНЕГО ЗАПРОСА ИЗ АРХИВА']
                request[0]['НОМЕР СТРОКИ ИЗ АРХИВА'] = number_str + 1
                request[0]['№ ПОСЛЕДНЕГО ЗАПРОСА ИЗ АРХИВА'] = request[0]['№ЗАПРОСА']

            elif name == 'Шасси':
                part_bp = request[0]['БП']
                part_fan = request[0]['FAN']
                comment = request[0]['КОММЕНТАРИИ']
                request[0] = {}
                request[0]['ЗИП'] = MESSAGE[2].format(part_bp=part_bp, part_fan=part_fan, comment=comment)
            data[key].update(request[0])

def search(db: Database, keys, 
           data, excel: Excel, atribute: tuple):
    """Add_directory_item."""

    global FIND_KEY
    value = ''
    flag = False
    name, col = atribute
    for item in keys:
        request = db.takeinfo(SQL_REQUEST[name][0].format(part=f"{excel.filterkey(item, 'PN')}"))
        if request:
            value = request[0][col]
            if value:
                FIND_KEY = item
                value = item
                break

    if not value:
        request = db.takeinfo(SQL_REQUEST[name][1].format(part=f"{excel.filterkey(keys[0], 'PN')}"))
        if request:
            value = request[0][col]
            if value:
                FIND_KEY = request[0]['MAIN_KEY']
                value = MESSAGE[1].format(part=FIND_KEY)

        else:
            request = db.takeinfo(SQL_REQUEST[name][2].format(part=f"{excel.filterkey(keys[0], 'PN')}"))
            if request:
                value = request[0][col]
                if value:
                    FIND_KEY = request[0]['MAIN_KEY']
                    value = MESSAGE[1].format(part=FIND_KEY)
    
    if value:

        flag = True
        add(data, name, request[0], value, keys[0])

    return flag

def add(data, name, request, value, main_key):
    """Add_item_in_data."""

    exc = (0, '0', None, 'None')
    if name == 'Свод':
        if 'ЗИП' in data[main_key]:
            request['ЗИП'] = value + '\n' + data[main_key]['ЗИП']
        else:
            request['ЗИП'] = value
        data[main_key].update(request)


    if name == 'Закупка1':
        client = request['ДТК СЕРВИС (КОММЕНТАРИИ ИНЖЕНЕРОВ)']
        request['ДТК СЕРВИС (КОММЕНТАРИИ ИНЖЕНЕРОВ)'] = f'Закупается под: {client}'
        if 'ЗИП' in data[main_key]:
            request['ЗИП'] = value + '\n' + data[main_key]['ЗИП']
        else:
            request['ЗИП'] = value
        data[main_key].update(request)

    if name == 'Закупка2':
        pn = request['PN']
        client = request['КЛИЕНТЫ']
        grade = request['ОЦЕНОЧНАЯ СТОИМОСТЬ']
        atr = request['ЗАКУПАЕМ ПОД ЗАКАЗЧИКА']
        if atr in exc:
            data[main_key]['ДТК СЕРВИС (КОММЕНТАРИИ ИНЖЕНЕРОВ)'] = MESSAGE[0].format(atr=client, pricegrabe=grade)
        else:
            data[main_key]['ДТК СЕРВИС (КОММЕНТАРИИ ИНЖЕНЕРОВ)'] = MESSAGE[0].format(atr=atr, pricegrabe=grade)
        del request['КЛИЕНТЫ']
        del request['ОЦЕНОЧНАЯ СТОИМОСТЬ']
        del request['ЗАКУПАЕМ ПОД ЗАКАЗЧИКА']
        if 'ЗИП' in data[main_key]:
            request['ЗИП'] = pn + '\n' + data[main_key]['ЗИП']
        else:
            request['ЗИП'] = pn
        data[main_key].update(request)

    if name == 'Архив':
        number_str = request['№ ПОСЛЕДНЕГО ЗАПРОСА ИЗ АРХИВА']
        request['НОМЕР СТРОКИ ИЗ АРХИВА'] = number_str + 1
        request['№ ПОСЛЕДНЕГО ЗАПРОСА ИЗ АРХИВА'] = request['№ЗАПРОСА']
        zip = request['ЗИП']
        if 'совпал без' in value:
            request['ЗИП'] = value + ': ' + zip   
        if 'ЗИП' in data[main_key]:
            if request['ЗИП'] != 'None':
                request['ЗИП'] = data[main_key]['ЗИП'] + '\n' + request['ЗИП']
            else:
                request['ЗИП'] = value + '\n' + data[main_key]['ЗИП']
        data[main_key].update(request)

    data[main_key]['ГДЕ НАШЛИ'] = name

def compilate(item: dict, db: Database, 
              excel: Excel, ebay: Ebay):
    """Create_final_collection."""

    global FIND_KEY
    vendor = ''
    comment = ''
    cup = ','
    for key in item:
        print(key)
        if key == 'SHEET':
            break

        FIND_KEY = key
        keys: list = []

        if 'ОПИСАНИЕ' in item[key]:
            comment = item[key]['ОПИСАНИЕ']
        if 'ВЕНДОР' in item[key]:
            vendor = item[key]['ВЕНДОР'].upper()

        amount = item[key]['КОЛИЧЕСТВО']
        keys = excel.exceptions(key, vendor)
        print(vendor, "VENDOR")
        print(keys, "EXCEPT")

        # add_huawei_element
        if vendor == 'HUAWEI':
            element = cup.join(keys[1:])
            item[key]['MODEL/PN'] = element
        
        # check chassis
        result = find(db, key, item, excel, SEARCH[5])

        # add_arch_book
        result = search(db, keys, item, excel, SEARCH[0])
        
        if not result:
            # add_purchase_one
            result = search(db, keys, item, excel, SEARCH[1])
    
        if not result:
            # add_purchase_two
            result = search(db, keys, item, excel, SEARCH[2])
        
        if not result:
            # add_archive
            result = search(db, keys, item, excel, SEARCH[3])
        
        # add_qty
        if '№ ПОСЛЕДНЕГО ЗАПРОСА ИЗ АРХИВА' not in item[key]:
            find(db, key, item, excel, SEARCH[4])

        # add_category
        category(db, item, amount, excel, key, comment)

        # add_ebay_search
        if item[key]['КАТЕГОРИЯ'] not in CATEGORY_EXCEPTION:
            ebay.searchebay(FIND_KEY, item, excel, vendor, key)    

    return item

def create_db(table: list, param: list, 
              db: Database, file: File, excel: Excel):
    """Create_db."""

    format_xlsm = ('Закупка1', 'Закупка2', 'Свод')
    try:
        for name in table:
                ind = table.index(name)
                local_dir = ROOT_DIR + "\\" + name
                print(local_dir)

                # main_object
                if name in format_xlsm:
                    root = file.find_file(local_dir, format=".xlsm")
                else:
                    root = file.find_file(local_dir)
                xlsx_list = db.filing(param[ind], root[0], excel)
                db.create(name, xlsx_list, param[ind])
    except:
        raise 'Error create Database.'

def main():    
    """Main_func."""

    # main_object
    global A
    db = Database(DATABASE)
    write = Writer()
    ebay = Ebay(API_KEY, CERT_ID, DEV_ID, TOKEN)
    excel = Excel()
    filee = File()

    # # create_database
    # db_point = input("Create Database?").upper()
    # if db_point == "YES":
    #     create_db(SUP_DIR, PARAM, db, filee, excel)

    while(True):

        sleep(15)
        mail = Email(MAIL_SERVER, USERNAME_GMAIL, PASSWORD_GMAIL)
        email_id = mail.check_folder('ALL')

        if email_id:
            email_message = mail.check_message(email_id)

            if email_message:
                files = mail.get_attachments(email_message, filee)

                if files:
                    if type(files) is not list:
                        files = [files]

                    flag = False
                    message = InfoMessage(subject=mail.SUBJECT)
                    for file in files:
                        xlsx_list = excel.find_item(SAMPLE_XLSX, file[0])
                        remove(file[0])

                        if not xlsx_list:
                            continue
                        
                        filee.FILENAME = file[1]
                        message.filename = filee.FILENAME
                        flag = True
                        for item in xlsx_list:
                        
                            data = compilate(item, db, excel, ebay)
                            write.writeinfo(data)
                            message.sheetname = item['SHEET']
                            mail.send_email(message.finalbody(write),
                                            write.OUTPUT_FILE,
                                            filee.FILENAME)

                            remove(write.OUTPUT_FILE)
                    if not flag:
                        mail.send_email(message.get_message(3))

"""Main_driver"""
if __name__ == '__main__':
    main()

