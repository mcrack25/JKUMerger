from app.models.InputNew import Input

class GetDataController():

    def getContent(self):
        data = Input().getAll()
        return data