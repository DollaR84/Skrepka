"""
Graphic form change settings.

Created on 19.06.2018

@author: Ruslan Dolovanyuk

"""

import wx


class Settings(wx.Dialog):
    """Create interface settings dialog."""

    def __init__(self, parent, config):
        """Initialize interface."""
        super().__init__(parent, wx.ID_ANY, 'Настройки')
        self.config = config
        title_out = 'Выбор каталога для вывода результата:'

        box_out_browse = wx.StaticBox(self, wx.ID_ANY, 'Каталог вывода результата')
        self.out_ctrl = wx.DirPickerCtrl(box_out_browse,
                                         wx.ID_ANY,
                                         self.config.out_path,
                                         title_out,
                                         style=wx.DIRP_USE_TEXTCTRL |
                                         wx.DIRP_DIR_MUST_EXIST)
        self.out_ctrl.GetPickerCtrl().SetLabel('Обзор...')
        but_save = wx.Button(self, wx.ID_OK, 'Сохранить')
        but_cancel = wx.Button(self, wx.ID_CANCEL, 'Отмена')

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer_out = wx.StaticBoxSizer(box_out_browse, wx.HORIZONTAL)
        sizer_out.Add(self.out_ctrl, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(sizer_out, 0, wx.EXPAND | wx.ALL)
        sizer_but = wx.GridSizer(rows=1, cols=2, hgap=5, vgap=5)
        sizer_but.Add(but_save, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        sizer_but.Add(but_cancel, 0, wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL)
        sizer.Add(sizer_but, 0, wx.EXPAND | wx.ALL)
        self.SetSizer(sizer)
