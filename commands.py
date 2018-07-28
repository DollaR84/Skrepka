"""
Commands for graphical interface.

Created on 13.06.2018

@author: Ruslan Dolovanyuk

"""

import logging
import os

from dialogs.dialogs import About


class Commands:
    """Helper class, contains command for bind events, menu and buttons."""

    def __init__(self, drawer, api_xl):
        """Initilizing commands class."""
        self.log = logging.getLogger()
        self.drawer = drawer
        self.api_xl = api_xl
        self.db = self.api_xl.db

        if not self.db.if_exists('window'):
            self.setup()

    def open(self, event):
        """Open excel file."""
        title = 'Выбор excel файла'
        wildcard = 'Excel 2007 files (*.xlsx)|*.xlsx|' \
                   'Excel 2003 files (*.xls)|*.xls|' \
                   'All files (*.*)|*.*'
        path = self.drawer.open(title, wildcard)
        if '' != path:
            self.api_xl.open(path)
            self.drawer.sheets(self.api_xl.sheet_names)
            self.drawer.log.AppendText('Файл %s открыт\n' % path)
            self.drawer.log.AppendText('Редактирование %s\n' % self.api_xl.sheet_names[0])

    def open_dir(self, event):
        """Open dir with excel files."""
        title = 'Выбор папки с excel файлами'
        path = self.drawer.open_dir(title)
        if '' != path:
            self.api_xl.open_dir(path)
            self.drawer.log.AppendText('Папка %s открыта для работы\n' % path)

    def save(self, event):
        """Save excel document after correct."""
        self.api_xl.save()
        self.drawer.log.AppendText('Файл %s сохранен\n' % self.api_xl.xl_name)

    def change_sheet(self, event):
        """Change sheet on excel workbook."""
        index = 0
        for item in self.drawer.items:
            if item.IsChecked():
                index = self.drawer.items.index(item)
        self.api_xl.change_sheet(index)
        self.drawer.log.AppendText('Редактирование %s\n' % self.api_xl.sheet_names[index])
        self.log.info('Редактирование %s' % self.api_xl.sheet_names[index])

    def clear(self, event):
        """Close current excel document."""
        self.drawer.log.AppendText('Файл %s закрыт\n' % self.api_xl.xl_name)
        self.log.info('Файл %s закрыт' % self.api_xl.xl_name)
        self.api_xl.clear()
        self.drawer.clear()

    def clear_dir(self, event):
        """Close current dir with excel documents."""
        self.drawer.log.AppendText('Папка %s закрыта\n' % self.api_xl.dir)
        self.log.info('Папка %s закрыта' % self.api_xl.dir)
        self.api_xl.clear()

    def verify(self, event):
        """Clear some columns."""
        if '' == self.api_xl.dir:
            self.verify_file()
        else:
            files = os.listdir(self.api_xl.dir)
            files = [x for x in files if ('.xlsx' == os.path.splitext(x)[1] or '.xls' == os.path.splitext(x)[1])]
            for file_name in files:
                self.api_xl.open(file_name)
                self.drawer.log.AppendText('Файл %s открыт\n' % self.api_xl.xl_name)
                self.drawer.log.AppendText('Редактирование %s\n' % self.api_xl.sheet_names[0])
                self.verify_file()
                self.save(None)
                self.drawer.log.AppendText('Файл %s закрыт\n' % self.api_xl.xl_name)
                self.log.info('Файл %s закрыт' % self.api_xl.xl_name)
                self.api_xl.clear()

    def verify_file(self):
        """Clear 1, 2, 8, 9, 10, 11, 12, 13, 14 columns in one file."""
        self.api_xl.verify()
        title = self.api_xl.sheet.title if self.api_xl.xlsx else self.api_xl.rsheet.name
        self.drawer.log.AppendText('Очищены некоторые столбцы %s\n' % title)

    def settings(self, event):
        """Open dialog for settings program."""
        self.api_xl.config.open_settings(self.drawer)

    def set_window(self):
        """Set size and position window from saving data."""
        script = 'SELECT * FROM window'
        data = self.db.get(script)[0]

        self.drawer.SetPosition((data[1], data[2]))
        self.drawer.SetSize((data[3], data[4]))
        self.drawer.Layout()

    def setup(self):
        """Create table window settings."""
        scripts = []
        script = '''CREATE TABLE window (
                    id INTEGER PRIMARY KEY NOT NULL,
                    px INTEGER NOT NULL,
                    py INTEGER NOT NULL,
                    sx INTEGER NOT NULL,
                    sy INTEGER NOT NULL) WITHOUT ROWID
                 '''
        scripts.append(script)
        script = '''INSERT INTO window (id, px, py, sx, sy)
                    VALUES (1, 0, 0, 800, 600)'''
        scripts.append(script)
        self.db.put(scripts)

    def about(self, event):
        """Run about dialog."""
        About(self.drawer,
              'О программе...',
              'Помошник бухгалтера',
              '1.0',
              'Руслан Долованюк').ShowModal()

    def close(self, event):
        """Close event for button close."""
        self.drawer.Close(True)

    def close_window(self, event):
        """Close window event."""
        pos = self.drawer.GetScreenPosition()
        size = self.drawer.GetSize()

        scripts = []
        script = 'UPDATE window SET px=%d, py=%d WHERE id=1' % tuple(pos)
        scripts.append(script)
        script = 'UPDATE window SET sx=%d, sy=%d WHERE id=1' % tuple(size)
        scripts.append(script)
        self.db.put(scripts)

        self.drawer.Destroy()
