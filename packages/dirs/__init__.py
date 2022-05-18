import os
import json

class Dirs:

    fileName = 'dirs.json'

    def get(self, path):
        rootDir = os.getcwd()
        dirConfigFile = os.path.join(rootDir, 'config', self.fileName)

        with open(dirConfigFile, encoding='windows-1251') as file:
            configs = json.load(file)

        if path in configs:
            oldPath = configs[path]
            pathArray = oldPath.split('.')
            halfPath = os.path.join(*pathArray)
            fullPath = os.path.join(rootDir, halfPath)
            return fullPath

        return ''