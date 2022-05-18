import os
from packages.dirs import Dirs

class StructController:
    dirs = []

    def __init__(self):
        self.dirs.append(Dirs().get('input'))
        self.dirs.append(Dirs().get('output'))
        self.dirs.append(Dirs().get('template'))
        self.dirs.append(Dirs().get('logs'))

    def makeStorage(self):
        for dir in self.dirs:
            if not (os.path.exists(dir)):
                os.makedirs(dir)