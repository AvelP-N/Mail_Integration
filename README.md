# Mail_Integration

Данный код содержит набор модулей для работы с различными типами файлов и баз данных. Программа использует библиотеки colorama, xml.etree.ElementTree, win32com.client, xlrd, glob, time, re, sqlite3, threading, win32api, pyodbc и os. Программа предназначена для автоматизации работы с файлами и базами данных. Она содержит функции для работы с XML-файлами, чтения данных из файлов формата XLS, поиска файлов по шаблону, работы с базами данных SQLite и MS SQL Server, а также для работы с потоками и консольным выводом.

- Программа для автоматического заполнения шаблона ответов по интеграции.
- В данной программе выполняются SQL запросы и обработка данных для формирования письма.
- Если база не возвращает информацию по заявке, то производится поиск XML файла в указанной папке и парсинг найденного
файла с данными клиента.
- Так же спользуется поиск менеджера отдела продаж в XLS файле Клиентской базы.
- Когда все данные найдены производится поиск в отправленных письмах на совпадение значений.
- Если письмо не отправлялось то формируется новое письмо.
