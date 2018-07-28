"""
Config for XLEditor.

Created on 13.06.2018

@author: Ruslan Dolovanyuk

"""

import logging

from dialogs.dialogs import RetCode

from dialogs.settings import Settings


class Config:
    """Class settings for XLEditor."""

    def __init__(self, db):
        """Initialize config class."""
        self.log = logging.getLogger()
        self.db = db

        if not self.db.if_exists('settings'):
            self.setup()

        script = 'SELECT * FROM settings'
        data = self.db.get(script)
        self.ids = {}
        for line in data:
            setattr(self, line[1], line[2])
            self.ids[line[1]] = line[0]

    def open_settings(self, parent):
        """Open settings dialog."""
        dlg = Settings(parent, self)
        if RetCode.OK == dlg.ShowModal():
            scripts = []
            script = '''UPDATE settings SET value="%s" WHERE id=%d
                 ''' % (dlg.out_ctrl.GetPath(), self.ids['out_path'])
            scripts.append(script)
            self.db.put(scripts)
            self.log.info('saving settings...')
        dlg.Destroy()

    def setup(self):
        """Create table settings in database."""
        scripts = []
        script = '''CREATE TABLE settings (
                    id INTEGER PRIMARY KEY NOT NULL,
                    name TEXT NOT NULL,
                    value TEXT NOT NULL) WITHOUT ROWID
                 '''
        scripts.append(script)
        script = '''INSERT INTO settings (id, name, value)
                    VALUES (1, "out_path", "output")'''
        scripts.append(script)
        self.db.put(scripts)
