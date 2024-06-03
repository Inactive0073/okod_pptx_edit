import datetime as dt


class Employee:
   def __init__(self, fio, position, date) -> None:
       self.fio:str = fio
       self.position:str = position
       self.date:str = date 


class Date:
    def __init__(self, compare_date=None) -> None:
        today = dt.datetime.today()
        self.cur_day:int = today.day
        self.cur_month:int = today.month
        if compare_date:
            day, month = map(int,compare_date.split('.'))
            self.full_date = dt.date(2024,month, day)
        # print(f'Проверка на дату прошла удачно\nПроверяется дата: {self.cur_day}.{self.cur_month}')
    
    
    @property
    def day(self) -> int:
        return self.full_date.day

    @property
    def month(self) -> int:
        return self.full_date.month
    
    
    def plus_day(self) -> None:
        self.full_date = dt.date(2024, self.month, self.day) + dt.timedelta(days=1)