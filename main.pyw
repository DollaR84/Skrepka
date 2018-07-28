"""
Main module for running XLEditor.

Created on 13.06.2018

@author: Ruslan Dolovanyuk

"""

from api import ApiXL

from drawer import Drawer

from logger import Logger


if __name__ == '__main__':
    log = Logger('Skrepka')
    api_xl = ApiXL()
    drawer = Drawer(api_xl)
    drawer.mainloop()
    api_xl.close()
    log.finish()
