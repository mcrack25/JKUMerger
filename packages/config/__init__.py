import os
import os.path
import json

class Config:

    def __init__(self, fileName='app', encode='utf-8'):
        self.encode = encode
        self.fileName = fileName + '.json'

    def get(self, param=None):
        rootDir = os.getcwd()
        configFile = os.path.join(rootDir, 'config', self.fileName)

        with open(configFile, encoding=self.encode) as file:
            configs = json.load(file)
            if param == None:
                return configs
            elif param in configs:
                value = configs[param]
                return value
        return ''