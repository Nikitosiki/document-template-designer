
# Класс описывающий сущьность строки
# с сохранением статуса использованости.
# Используется в значениях, в словаре
class UseString:
    def __init__(self, value):
        self._value = value
        self._is_use = False

    @property
    def value(self):
        return self._value
    
    @property
    def is_use(self):
        return self._is_use
    
    @is_use.setter
    def is_use(self, is_use):
        self._is_use = is_use