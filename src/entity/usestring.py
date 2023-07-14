
# Класс описывающий сущьность строки
# с сохранением статуса использованости.
# Используется в значениях, в словаре
class UseString:
    def __init__(self, value, is_image, image_width = 0, image_height = 0):
        self._value = value
        self._is_image = is_image
        self._image_width = image_width
        self._image_height = image_height
        self._is_use = False

    @property
    def value(self):
        return self._value
    
    @property
    def is_use(self):
        return self._is_use
    
    @property
    def is_image(self):
        return self._is_image
    
    @property
    def image_width(self):
        return self._image_width
    
    @property
    def image_height(self):
        return self._image_height
    
    @is_use.setter
    def is_use(self, is_use):
        self._is_use = is_use