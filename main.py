import editor_pptx
from os import system


try:
    if __name__ == 'main':
        editor_pptx.start()
        print('Функция завершила работу успешно.')
except Exception as err:
    print(err)
finally:
    system('pause')
    