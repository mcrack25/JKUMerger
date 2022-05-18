from packages.config import Config
from app.controllers.StructController import StructController
from app.controllers.GetDataController import GetDataController
from app.controllers.SetDataController import SetDataController

if(__name__ == "__main__"):
    print('Программа запущена!!!')
    print()

    # Создаём структуру папок Storage
    StructController().makeStorage()

    data = GetDataController().getContent()
    SetDataController().setContent(data)

    print()
    print('Программа завершена!!!')
