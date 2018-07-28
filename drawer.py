"""
Graphical form for XLEditor.

Created on 15.06.2018

@author: Ruslan Dolovanyuk

"""

from commands import Commands

import wx


class Drawer:
    """Main class graphical form for XLEditor."""

    def __init__(self, api_xl):
        """Initilizing drawer form."""
        self.app = wx.App()
        self.wnd = XLEditorFrame(api_xl)
        self.wnd.Show(True)
        self.app.SetTopWindow(self.wnd)

    def mainloop(self):
        """Graphical main loop running."""
        self.app.MainLoop()


class XLEditorFrame(wx.Frame):
    """Create user interface."""

    def __init__(self, api_xl):
        """Initialize interface."""
        super().__init__(None, wx.ID_ANY, 'Помошник бухгалтера')
        self.command = Commands(self, api_xl)

        panel = wx.Panel(self, wx.ID_ANY)
        sizer_panel = wx.BoxSizer(wx.HORIZONTAL)
        sizer_panel.Add(panel, 1, wx.EXPAND | wx.ALL)
        self.SetSizer(sizer_panel)

        self.CreateStatusBar()
        self.__create_menu()

        self.log = wx.TextCtrl(panel, wx.ID_ANY,
                               style=wx.TE_MULTILINE | wx.TE_PROCESS_ENTER | wx.TE_READONLY)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.log, 1, wx.EXPAND | wx.ALL, 5)
        panel.SetSizer(sizer)

        self.Bind(wx.EVT_CLOSE, getattr(self.command, 'close_window'))
        
        self.command.set_window()

    def __create_menu(self):
        """Create menu interface."""
        menu_file = wx.Menu()
        xl_open = menu_file.Append(-1, 'Открыть...', 'Нажмите для открытия excel файла')
        xl_save = menu_file.Append(-1, 'Сохранить', 'Нажмите для сохранения excel файла')
        xl_close = menu_file.Append(-1, 'Закрыть', 'Нажмите для закрытия excel файла')
        menu_file.AppendSeparator()
        xl_dir = menu_file.Append(-1, 'Открыть папку...', 'Нажмите для указания папки для обработки')
        xl_close_dir = menu_file.Append(-1, 'Закрыть папку', 'Нажмите для закрытия папки с excel файлами')
        menu_file.AppendSeparator()
        exit = menu_file.Append(-1, 'Выход', 'Нажмите для выхода из программы')

        menu_sheets = wx.Menu()

        menu_operations = wx.Menu()
        verify = menu_operations.Append(-1, 'На сверку', 'Нажмите для очистки лишних столбцов')

        menu_options = wx.Menu()
        settings = menu_options.Append(-1, 'Настройки...', 'Нажмите для изменения настроек помошника')

        menu_help = wx.Menu()
        about = menu_help.Append(-1, 'О программе...', 'Выводит информацию о программе')

        self.Bind(wx.EVT_MENU, getattr(self.command, 'open'), xl_open)
        self.Bind(wx.EVT_MENU, getattr(self.command, 'save'), xl_save)
        self.Bind(wx.EVT_MENU, getattr(self.command, 'clear'), xl_close)
        self.Bind(wx.EVT_MENU, getattr(self.command, 'open_dir'), xl_dir)
        self.Bind(wx.EVT_MENU, getattr(self.command, 'clear_dir'), xl_close_dir)
        self.Bind(wx.EVT_MENU, getattr(self.command, 'close'), exit)
        self.Bind(wx.EVT_MENU, getattr(self.command, 'verify'), verify)
        self.Bind(wx.EVT_MENU, getattr(self.command, 'settings'), settings)
        self.Bind(wx.EVT_MENU, getattr(self.command, 'about'), about)

        menuBar = wx.MenuBar()
        menuBar.Append(menu_file, 'Файл')
        menuBar.Append(menu_sheets, 'Листы')
        menuBar.Append(menu_operations, 'Операции')
        menuBar.Append(menu_options, 'Опции')
        menuBar.Append(menu_help, 'Справка')
        self.SetMenuBar(menuBar)

        menuBar.EnableTop(1, False)

    def open(self, title, wildcard):
        """Open file dialog."""
        path = ''
        file_dlg = wx.FileDialog(self, title, '', '', wildcard, style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if wx.ID_OK == file_dlg.ShowModal():
            path = file_dlg.GetPath()
        return path

    def sheets(self, sheet_names):
        """Add sheets menu items in menu sheets."""
        menuBar = self.GetMenuBar()
        menu_sheets = menuBar.GetMenu(1)
        self.items = []
        for name in sheet_names:
            item = menu_sheets.AppendRadioItem(-1, name, 'Редактировать %s' % name)
            self.items.append(item)
        self.Bind(wx.EVT_MENU_RANGE, getattr(self.command, 'change_sheet'), self.items[0], self.items[-1])
        menuBar.EnableTop(1, True)

    def clear(self):
        """Clear menu sheets."""
        self.items.clear()
        menuBar = self.GetMenuBar()
        menuBar.EnableTop(1, False)

    def open_dir(self, title):
        """Open dir dialog."""
        path = ''
        dir_dlg = wx.DirDialog(self, title, '', style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
        if wx.ID_OK == dir_dlg.ShowModal():
            path = dir_dlg.GetPath()
        return path
