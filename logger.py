"""
Module logger for projects.

Created on 05.07.2018

@author: Ruslan Dolovanyuk

"""

import logging
import logging.handlers
import os


class Logger:
    """Class for initialization of the logging module."""

    def __init__(self, name):
        """Initialize logger class."""
        file = open(name + '.log', 'w')
        file.close()
        self.name = name
        self.log = logging.getLogger()
        self.log.setLevel(logging.INFO)
        formatter = logging.Formatter('[%(asctime)s] %(levelname)s: '
                                      '%(message)s', '%d.%m.%Y %H:%M %Ss')
        handler = logging.StreamHandler()
        handler.setFormatter(formatter)
        handler.setLevel(logging.DEBUG)
        self.log.addHandler(handler)
        handler = logging.FileHandler(name + '.log', 'a')
        handler.setFormatter(formatter)
        handler.setLevel(logging.INFO)
        self.log.addHandler(handler)
        self.log.info(">>>>> %s start <<<<<" % self.name.upper())

    def finish(self):
        """Shutdown the logger."""
        self.log.info(">>>>> %s end <<<<<" % self.name.upper())
        logging.shutdown()
