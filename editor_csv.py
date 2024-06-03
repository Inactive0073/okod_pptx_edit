import xlrd
import csv
from re import sub
from datasets import Employee, Date



count = 0 # счетчик дней рождений, которые надо будет разместить
path = 'Дни рождения.csv'
def convert_csv_from_xls() -> None:
    '''Конвертирует файл из формата xls в csv.'''
    work_book = xlrd.open_workbook_xls('дни рождения.xls') # создаем экземпляр с параметрами нашей книги
    print(f'Список листов: {work_book.sheet_names()}\nКоличество листов в документе: {work_book.nsheets}')
    sheet = work_book.sheet_by_name(work_book.sheet_names()[0])
    try:
        with open(path,'w', encoding='utf-8') as csvfile:
            wr = csv.writer(csvfile,quoting=csv.QUOTE_ALL)
            [wr.writerow(sheet.row_values(row_num)) for row_num in range(sheet.nrows)]
    except FileNotFoundError as err:
        print('Файл для конвертации не найден:', err)
    except Exception as err:
        print('Произошла непредвиденная ошибка:', err)
              
           
def get_data_from_csv() -> list[str]:
    '''Возвращает список людей с их ПД у которых были или есть дни рождения в текущую дату
    
    Returns:
        list[str]: Список сотрудников, состоящий из 4 столбцов.
    '''
    try:
        with open(path,encoding='utf-8') as csv_file:
            reader = csv.reader(csv_file)
            return [row for row in reader if row]
    except FileNotFoundError as err:
        print(f'Не найдено подходящего файла, возникла ошибка {err}')
    except:
        print('Произошла непредвиденная ошибка')


    
def filter_employees(employee_info:list[str]) -> list[str]:
    '''Возвращает отсортированный список людей в формате ФИО, должность, дата дд.мм, которые будут добавлены в слайд на презентации.\n
    Принимает список сотрудников состоящий из 4 столбцов
    
    Args:
        employee_info (list[str]): Список сотрудников, состоящий из 4 столбцов.
    
    Returns:
        list[tuple[str, str, str]]: Список кортежей с данными сотрудников.
        '''
    date_compare = input('\nВведите дату отсечки в формате: "dd.mm".\nЕсли нужен текущий день, нажмите "Enter".\n')
    date_finished = input('\nВведите дату итогового формирования в формате: "dd.mm".\nЕсли нужен текущий день, нажмите "Enter".\n')
    res_lst: list[str] = []
    for row in employee_info:
        _, fio, position, date = row
        emp = Employee(fio, position, date)
        if is_birthday_today(emp.date, date_compare, date_finished):
            res_lst.append((emp.fio,emp.position,'.'.join(emp.date.split('.')[:-1]))) # В результирующий список передаем данные 
    return res_lst


def edit_date(day,month) -> str:
    return f'{str(day).zfill(2)}.{str(month).zfill(2)}'


def is_birthday_today(date_birthday:str, date_compare=None, date_finish=None):
    '''Проверка на наличие дня рождения в эту дату или с момента даты отсечки, 
    указанной вручную в функции filter_employees
    
    Args:
        date_birthday (str): Дата рождения в формате 'dd.mm'.
        date_compare (str): Дата отсечки в формате 'dd.mm'.
    
    Returns:
        bool: True, если день рождения в эту дату или с момента даты отсечки, False иначе.
    
    '''
    global count
    d1 = Date(date_compare)
    if date_finish:
        cur_date = date_finish
    else:
        cur_date = edit_date(d1.cur_day,d1.cur_month)
    date_birthday = sub(r'(\d{2}\.\d+)\.\d{4}', r'\1', date_birthday) # дата рождения в формате dd.mm
    
    if not date_compare or date_birthday in (date_compare, cur_date):
        if date_birthday == cur_date:
            count += 1
            print(f'Найдено {count} дня(ей) рождений, которые будут размещены сегодня...')
            return True
    else:
        while date_compare != cur_date:
            if date_birthday == date_compare:
                count += 1
                print(f'Найдено {count} дня(ей) рождений, которые будут размещены сегодня...')
                return True
            else:
                d1.plus_day()
                date_compare = edit_date(d1.day,d1.month)


def sorted_list(lst):
    d = {}
    for row in lst:
        key = row[2]
        d[key] = d.get(key, []) + [row[:2]]
    return d


convert_csv_from_xls() # Конвертация
list_employees: list[str] = get_data_from_csv() # Сохранение в список из csv файла
pre_filtered_list_employees = filter_employees(list_employees)
filtered_list = sorted_list(pre_filtered_list_employees)

