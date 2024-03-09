import os
from re import findall
from random import choices
from glob import iglob
from time import sleep
from xlrd import open_workbook
import pyodbc
import sqlite3 as sq
import threading
import win32api
import win32com.client as client
from colorama import init, Fore, Back, Style
from xml.etree import ElementTree as et


def main():
    def request_check_block_test(index_list_test):
        """Проверка блокировки теста"""

        nonlocal block_test

        def check_block_test():
            """Проверка блокировки теста"""

            return cursor.execute(f"""SELECT DISTINCT
                                        [IDtest],
                                        [RegionCode],
                                        [DateFrom],
                                        [DateTo],
                                        [InformationLetter]
                                    FROM [RInStopTests]
                                    WHERE [IDtest] =  {list_test_code[index_list_test]}   """).fetchone()

        block_test = check_block_test()
        if block_test:
            list_header_email.append('Блокировка исполнения исследования')

    def request_check_available_test(index_list_test):
        """Проверка доступности исследования"""

        nonlocal available_test

        def check_available_test():
            return cursor.execute(f"select [dbo].[getIdByShortName3] "
                                  f"('{sender}' , {list_test_code[index_list_test]} )").fetchone()

        test = check_available_test()
        if not test[0]:
            available_test = 'Исследование недоступно для отправителя'
            list_header_email.append('Исследование недоступно для отправителя')
        else:
            available_test = 'Исследование доступно, необходимо доназначить'
            list_header_email.append('Исследование доступно, необходимо доназначить')

    def request_info_person():
        """Поиск данных по № заявки"""

        def connect_drive():
            """Если SQL база не вернула результат, то подключаем сетевой диск для поиска XML файла"""

            if not os.path.exists('x:/'):
                os.system('net use x: \\\\path /persistent:yes')

        def search_file():
            """Поиск файла XML в указанном каталоге"""

            nonlocal found_file

            if file_name:
                search_name_file = file_name
            else:
                search_name_file = number

            for find_file in iglob(f'X:\\{folder}\\*.*'):
                if search_name_file in find_file:
                    found_file = find_file
                    os.startfile(find_file)
                    break

        def scan_files():
            """Визуализация поиска файла"""

            chars = ('|  ', '/  ', '-- ', '\\  ', '|\\ ', '|/ ', '/| ', '\\| ', '__ ')
            print(Fore.BLUE + Style.BRIGHT, end='')
            while threading.active_count() > 1:
                print('\r', ' ', *choices(chars, k=5), ' ', end='')
                sleep(0.3)
            print(Style.RESET_ALL, '\r', sep='', end='')

        def info_person():
            """Информация по заявке"""

            person = f"""
            SELECT [OrderNumber], [IDsender], [dbo].[lc_orders].[DateIns], [PatientLastname], [PatientName], [PatientDoB], [PatientSex]
            FROM [dbo].[lc_orders] full join [dbo].[lc_patients]
            on [dbo].[lc_orders].IDpatient = [dbo].[lc_patients].IDpatient
            where OrderNumber in ({number})
            """
            return cursor.execute(person).fetchone()

        print(Fore.GREEN + 'Данные файла заявки:')
        print('Номер заявки\t\tКод отправителя\t\tДата создания заказа\t\tФамилия\tИмя-Отчество\tДата рождения\tПол')
        print(Fore.RESET, end='')

        data_person = info_person()

        if data_person:  # Если база SQL вернула данные по заявке
            for index_data_person in range(len(data_person)):
                if index_data_person == 2:  # Удаление тысячных секунд в "Дате создания заказа"
                    print(str(data_person[index_data_person]).split('.')[0], end='\t\t')
                elif index_data_person == 5:  # Удаление времени (0:00:00) в "Дата рождения"
                    print(str(data_person[index_data_person]).split()[0], end='\t\t')
                else:
                    print(data_person[index_data_person], end='\t\t')
        else:  # Если база SQL ничего не вернула
            connect_drive()

            found_file = ''
            if '_' in folder:
                threading.Thread(target=search_file).start()
                scan_files()

                # Если XML файл найден, то производится выборка данных по тегам
                if found_file:
                    try:
                        # Поиск номера заявки
                        # Если файл по типу \r_42KM7794_20220425_134257.XML выходит исключение
                        tree = et.parse(found_file)
                        print(tree.find('Order').attrib['OrderID'], end='\t')

                        # Поиск кода отправителя, даты и время создания заказа
                        header = tree.find('Header')
                        print(header.find('ClinicID').text, end='\t')
                        print(header.find('FileDate').text, end='\t')
                        print(header.find('FileTime').text, end='\t')

                        # Поиск фамилии, имя-отчества, даты рождения и пол
                        order = tree.find('Order')
                        patient = order.find('Patient')
                        print(patient.find('LastName').text, end='\t')
                        print(patient.find('FirstMiddleName').text, end='\t')
                        print(patient.find('DOB').text, end='\t')
                        print(patient.find('Sex').text, end='')
                    except AttributeError as attr:
                        print(Fore.RED + 'Tag error! Tag was not found.' + Fore.RESET, end='')
                    except Exception as ex:
                        print(Fore.RED + str(ex).capitalize() + Fore.RESET, end='')
                else:
                    print(Fore.RED + f'File {number} in folder {folder} not found!' + Fore.RESET, end='')
            else:
                print(Fore.RED + f'Incorrect name folder < {folder} >' + Fore.RESET, end='')

        print()

    def request_get_form(index_list_code):
        def get_form():
            """Получение формуляра по коду теста"""

            request_form = f"""
            SELECT b.*
            FROM [OrdersFromCACHE].[dbo].[lc_Blanks] b
            LEFT JOIN [OrdersFromCACHE].[dbo].[lc_BlanksHeader] bh ON b.Code=bh.Code
            WHERE b.IDTest = {list_test_code[index_list_code]}                                    
            ORDER BY Code   
            """
            return cursor.execute(request_form).fetchone()

        form = get_form()
        try:
            print(form[5], 'Причина: ', end='')
        except TypeError:
            print(Fore.RED + 'Получение формуляра по коду теста, БД ничего не вернула!' + Fore.RESET)
        else:
            if block_test:
                print(Fore.RED, end='')
                print('Заблокирован с', block_test[2])
                print(Fore.RESET, end='')
            else:
                print(Fore.MAGENTA + Style.BRIGHT + available_test + Style.RESET_ALL)

                # Если тест не доступен для отправителя, то добавляем в список unavailable_test
                if available_test == 'Исследование недоступно для отправителя':
                    unavailable_test.append(form[5])

            print('Название:', form[6])

    def request_get_form_code(n):
        """Получение формуляра по коду теста с удалением копий"""

        index_list_unavailable = n

        def get_form_code():
            sql_request = f"""
            SELECT DISTINCT lc_Blanks.Code, lc_Blanks.DBRegion
            FROM [OrdersFromCACHE].[dbo].[lc_Blanks],
             [OrdersFromCACHE].[dbo].[lc_BlanksHeader]
            WHERE lc_BlanksHeader.Code = lc_Blanks.Code
            --AND lc_BlanksHeader.IsActive='1'
            AND lc_Blanks.IDtest = '{unavailable_test[index_list_unavailable]}'
            """
            return cursor.execute(sql_request).fetchall()

        print(Fore.YELLOW + f'Бланки для - {unavailable_test[index_list_unavailable]}' + Fore.RESET)

        form_code = get_form_code()
        for row in form_code:
            print(row[0])

    def check_number_outlook():
        """Проверка номера заявки в отправленных письмах Outlook"""

        def search_number_email_folder(digit, count):
            sent_name_folder = namespace.GetDefaultFolder(digit)

            # Количество писем в выбранной папке
            count_email = sent_name_folder.Items.Count - 1

            # Проверка номера заявки на совпадение в последних письмах. В папках "Отправленные"
            for i in range(count_email, count_email - count, -1):
                email = sent_name_folder.Items[i]
                if number in email.Body:
                    print(Fore.RED + 'Письмо уже было отправлено!')
                    print('Кем:', email.SenderName)
                    print('Тема письма:', email.Subject + Fore.RESET)
                    return True

        # Создать объект outlook
        outlook = client.Dispatch('Outlook.Application')

        # Метод экземпляра Outlook
        namespace = outlook.GetNameSpace("MAPI")

        # Проверка номера заявки в папке "Отправленные", последние 50 писем
        if search_number_email_folder(5, 50):
            return True

    def emile_title():
        """Заполнение темы письма"""

        print(Fore.GREEN + 'Тема письма:' + Fore.RESET)
        print(Fore.YELLOW + f'Состав заказа {number} от отправителя {sender} ::: ', end='')
        print(*set(list_header_email), sep=' ::: ')
        print(Fore.RESET)

    def xls_senders():
        """Создание кортежа с отправителями из xls фала"""
        nonlocal tuple_rows_sender, xls_sheet_clients

        def find_client_database():
            """Проверяем наличие подключенного диска V: с файлом Клиентской базы"""

            def check_file_database():
                """Проверка файла Клиентской базы на диске V:"""
                nonlocal path_client_database
                for path in os.listdir("V:"):
                    find_file = findall(r'\AКлиентская база', path)
                    if find_file:
                        path_client_database = os.path.join('V:\\', path)
                        return True

            if os.path.exists("V:"):
                if check_file_database():
                    return True
            else:
                os.system(r'NET USE V: "\\path" /PERSISTENT:YES')
                print(Fore.GREEN + r'Connected network drive V: \\path' + Fore.RESET)
                if check_file_database():
                    return True

        path_client_database = ''

        # Создание объекта book для парсинга
        if find_client_database():
            book = open_workbook(filename=path_client_database)

            # Получить объект листа по индексу
            xls_sheet_clients = book.sheet_by_index(2)

            # Проверяем количество строк в листе
            sheet2_rows = xls_sheet_clients.nrows

            # Кортеж с отправителями
            tuple_rows_sender = tuple(map(str, xls_sheet_clients.col_values(2, 0, sheet2_rows)))
        else:
            print(Fore.RED + '\nФайл "V:\Клиентская база*.xls" не найден!' + Fore.RESET)
            print(Fore.RED + 'Поиск менеджера в клиентской базе производиться не будет!')

    def emile():
        """Поиск Менеджера отдела продаж по отправителю"""

        # Заполнение шапки письма для региональных отправителей и списка исключений
        def fill_email(to_massage):
            outlook = client.Dispatch('Outlook.Application')
            massage = outlook.CreateItem(0)
            massage.Display()
            massage.To = to_massage
            massage.CC = 'mails'
            massage.Subject = f'Состав заказа {number} от отправителя {sender} ::: ' + \
                              ' ::: '.join(set(list_header_email))

        def get_mail_manager(list_name_manager):
            """Возвращает почтовый ящик менеджера, если есть в БД, или пустой список, если менеджера нет в БД"""
            mail = cursor_sq.execute(f"""
                SELECT mail FROM users_mail WHERE name LIKE '%{" ".join(list_name_manager)}%'
            """).fetchone()
            return mail

        # Заполнение шапки письма для московских отравителей
        def fill_email_moscow_sender(to_massage):
            outlook = client.Dispatch('Outlook.Application')
            massage = outlook.CreateItem(0)
            massage.Display()

            if ufa:
                massage.To = f"mails; {to_massage}"
            else:
                massage.To = f"mail; {to_massage}"

            massage.CC = 'mails'
            massage.Subject = f'Состав заказа {number} от отправителя {sender} ::: ' + \
                              ' ::: '.join(set(list_header_email))

        flag = False
        ufa = False
        sender_exception = {'36VR11531': 'mail',
                            '42KM7794': 'mails',
                            '42KM9549': 'mails',
                            '42BA6694': 'mails',
                            '54KY6347': 'mails',
                            '77MS9989': 'Тестовый отправитель! Писать письмо не нужно, можно сразу закрывать заявку.',
                            '77MS9999': 'Тестовый отправитель! Писать письмо не нужно, можно сразу закрывать заявку.'}

        regions = {'KR': 'mails',
                   'PR': 'mails',
                   'KA': 'mails',
                   'EK': 'mails',
                   'VL': 'mails',
                   'RS': 'mails',
                   'AS': 'mails',
                   'SA': 'mails',
                   'TN': 'mails',
                   'OM': 'mails',
                   'NS': 'mails',
                   'NK': 'mails',
                   'UF': 'mails'}

        # Редактирование отправителя
        edit_sender = sender.split('_')[1]

        # Проверка исключений для отправителей
        for exception in sender_exception:
            if edit_sender == exception:
                print(Fore.GREEN + 'Кому:' + Fore.RESET)
                print(Fore.YELLOW + sender_exception[exception] + Fore.RESET, '\n')
                fill_email(sender_exception[exception])
                flag = True
                break

        # Проверка региональных отправителей
        if not flag:
            for region in regions:
                if region in edit_sender:
                    if region in 'UF':
                        ufa = True
                        break
                    print(Fore.GREEN + 'Кому:' + Fore.RESET)
                    print(Fore.YELLOW + regions[region] + Fore.RESET, '\n')
                    fill_email(regions[region])
                    flag = True
                    break

        # Проверка московских отправителей
        if not flag:
            for index_tuple, value in enumerate(tuple_rows_sender):
                if edit_sender in value:
                    print(Fore.GREEN + 'Кому:' + Fore.RESET)

                    if ufa:
                        print(Fore.YELLOW + 'mails', end='; ')
                    else:
                        print(Fore.YELLOW + 'mails', end='; ')

                    list_manager = findall(r'[А-ЯЁ][а-яё]+', str(xls_sheet_clients.cell(index_tuple, 13).value))
                    edit_list_manager = list(filter(lambda i: 'Моб' not in i, list_manager))

                    # Поиск почтового ящика менеджера в БД
                    if edit_list_manager:
                        mail_manager = get_mail_manager(edit_list_manager)

                        if mail_manager:
                            print(*mail_manager)
                            fill_email_moscow_sender(mail_manager[0])
                        else:
                            print(*list_manager)
                            fill_email_moscow_sender(' ')
                        print(Fore.RESET)
                        break
            else:
                print(Fore.GREEN + 'Кому:' + Fore.RESET)
                print(Fore.YELLOW + 'mail' + Fore.RESET,
                      Fore.RED + 'Manager not found in "Клиентская база"' + Fore.RESET, '\n')

    def input_text():
        """Обработка данных из письма"""

        def first_string(string):
            """Сбор данных из первой строки письма (номер заявки, каталог с файлом .xml и отправитель)"""

            nonlocal number, folder, sender, file_name, incorrect_sender

            # Find order number and save in variable number
            list_number = findall(r'[19]\d{9}', string)
            if list_number:
                number = list_number[0]
            else:
                list_number = findall(r'[19]\d{9}', ' '.join(email_text))
                if list_number:
                    number = list_number[0]
                    file_name = findall(r'(?i)\s(\w+)\.xml', string)[0]

            # Find folder and save in variable folder
            list_folder = findall(r'\\(\w+?)\\', string)
            if list_folder:
                folder = list_folder[0]

            list_sender = findall(r'\s(r_\w+)\s', string)
            if list_sender:
                sender = list_sender[0]
            else:
                incorrect_sender = True

        def input_email_text():
            """Пользовательский ввод содержимого письма и сохранение в список email_text"""

            print(Fore.LIGHTBLUE_EX + '\nВставьте содержимое письма и нажмите Enter:' + Fore.RESET)

            print(Fore.LIGHTYELLOW_EX, end='')
            text = list(iter(input, ''))
            print(Fore.RESET, end='')

            if text and text[0].startswith('При обработке'):
                for string in text:
                    if string not in email_text:
                        email_text.append(string)
            else:
                print(Fore.RED + 'Неверный ввод!\nТекст письма должен начинаться со слов "При обработке файла ..."')
                print(Fore.RESET, end='')

        # Пользовательский ввод с обращением к функции input_emile_text
        email_text = []
        while not email_text:
            input_email_text()

        # Вывод текста в консоль, из списка email_text, и сбор данных из письма
        print(Fore.GREEN + 'Добрый день!' + Fore.RESET)

        for index_string in range(len(email_text)):
            print(email_text[index_string])

            # Поиск № заявки, каталога и отправителя в первой строке
            if index_string == 0:
                first_string(email_text[index_string])
            else:
                # Поиск тестов и добавление в список list_test_code
                test_code = findall(r' (\'?)([А-ЯA-Z0-9]+\.[А-ЯA-Z0-9.]+)\1', email_text[index_string])
                if test_code:
                    for test in test_code:
                        edit_test = "'" + test[1].strip('.') + "'"
                        if edit_test not in list_test_code:
                            list_test_code.append(edit_test)

    # Начало программы
    numbers = []
    tuple_rows_sender = ()
    xls_sheet_clients = None

    # Извлечение отправителей из файла "V:\Клиентская база*.xls"
    xls_senders()

    while True:
        block_test = None
        available_test = ''
        number = folder = sender = file_name = ''
        list_test_code = []
        incorrect_sender = False
        unavailable_test = []
        list_header_email = []

        # Ввод и вывод в консоль содержимого письма
        input_text()

        # Проверка номера заявки
        if not number:
            print(Fore.RED + f'\nНомер заявки не найден!'
                             '\nНомер заявки должен состоять из 10 цифр.' + Fore.RESET)
            continue

        # Проверка отправителя
        if incorrect_sender:
            print(Fore.RED + f'\nПроверьте правильность написания отправителя!' + Fore.RESET)
            print(Fore.RED + 'Отправитель должен начинаться с "r_"')
            continue
        print()

        # Проверка списка с тестами
        if not list_test_code:
            print(Fore.RED + 'Тест не обнаружен!')
            print('Проверьте содержимое письма. Возможно текст выстроился в одну строку '
                  'или тесты не указаны в письме.' + Fore.RESET)
            continue

        # Проверка на дубли номеров заказа. Если письмо с полученным номером заказа уже обрабатывалось,
        # то пропустить все дальнейшие действия.
        if number not in numbers:
            numbers.append(number)
        else:
            print(Fore.RED + f'Письмо с номером заявки {number} уже было!' + Fore.RESET)
            for index in range(len(numbers)):
                if numbers[index] == number:
                    print(Fore.RED + numbers[index] + Fore.RESET, end=' ')
                else:
                    print(numbers[index], end=' ')

            print('\n')
            continue

        # Проверка номера заявки в отправленных письмах Outlook. Папки "Отправленные" и "Ханин Александр Михайлович"
        if check_number_outlook():
            continue

        # Вывод данных файла заявки
        try:
            request_info_person()
        except pyodbc.ProgrammingError:
            print(Fore.RED + f'Error input data! Check correct file number < {number} >' + Fore.RESET)
        print()

        print(Fore.GREEN + 'Назначения не вошедшие в заказ:' + Fore.RESET)
        for index in range(len(list_test_code)):
            if not list_test_code[index].isascii():
                print(list_test_code[index].strip("'"), 'Причина:',
                      Fore.MAGENTA + Style.BRIGHT + 'Русская буква в коде теста!' + Style.RESET_ALL)
                list_header_email.append('Русская буква в коде теста!')
            else:
                request_check_block_test(index)
                if not block_test:
                    request_check_available_test(index)
                request_get_form(index)

        # Если есть недоступные тесты и не заблокированные, то требуется вывести бланки для подключения
        if unavailable_test:
            print(Fore.GREEN)
            print("Для тестов, отмеченных как недоступные для отправителя необходимо подключение бланка\n"
                  "Если тест для МЦ должен выполняться, по условию договора, то: Требуется доназначить\n"
                  "Необходимо выбрать нужный бланк, из списка ниже, и составить заявку на SD для подключения бланка:")
            print(Fore.RESET, end='')
            for list_index in range(len(unavailable_test)):
                if list_index > 0:
                    print()
                request_get_form_code(list_index)

        print()

        # Поиск ответственного менеджера в клиентской базе
        emile()

        # Формирование заголовка письма
        emile_title()


if __name__ == '__main__':
    init()
    win32api.SetConsoleTitle('SQL Integration v3.9')
    print(Back.YELLOW + Fore.BLACK + '   Fill out the integration email form   ' + Style.RESET_ALL)

    try:
        with pyodbc.connect('Driver={SQL Server};'
                            'Server=Server;'
                            'UID=Login;'
                            'PWD=Password') as connection:
            cursor = connection.cursor()
    except pyodbc.OperationalError:
        quit(print('*** Server DB not available! ***'))

    with sq.connect(fr'C:\Users\{os.getlogin()}\PycharmProjects\Mail_integration\UsersDB.db') as connectDB:
        cursor_sq = connectDB.cursor()

    try:
        main()
    except KeyboardInterrupt:
        print('*** Emergency exit from the program! ***')
