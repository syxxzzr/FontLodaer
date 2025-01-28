# coding: utf-8
import wx
import ctypes
import tempfile
import uuid
import os
import os.path
import wx.adv
import configparser
from fontTools.ttLib.ttFont import TTFont
from fontTools.ttLib import TTCollection


# HWND_BROADCAST = 0xFFFF
# SMTO_ABORTIFHUNG = 0x0002
# WM_FONTCHANGE = 0x001D

CONFIRM = 'confirm'
HIDE = 'hide'
DESTROY = 'destroy'


class MainFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title='Font Loader')

        data_sizer = wx.BoxSizer(wx.VERTICAL)

        self.fontsList = wx.ListCtrl(self, style=wx.LC_REPORT | wx.BORDER_SUNKEN)
        self.fontsList.AppendColumn('Font Name')
        self.fontsList.AppendColumn('File Path', width=wx.LIST_AUTOSIZE_USEHEADER)
        data_sizer.Add(self.fontsList, 1, wx.ALL | wx.EXPAND, 5)

        control_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.loadFont = wx.Button(self, label='Load Font')
        control_sizer.Add(self.loadFont, 0, wx.ALL | wx.EXPAND, 5)

        self.releaseFont = wx.Button(self, label='Release Font')
        control_sizer.Add(self.releaseFont, 0, wx.ALL | wx.EXPAND, 5)

        self.releaseAll = wx.Button(self, label='Release All')
        control_sizer.Add(self.releaseAll, 0, wx.ALL | wx.EXPAND, 5)

        data_sizer.Add(control_sizer, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.SetSizer(data_sizer)
        self.Layout()

        self.Centre(wx.BOTH)

        file_drop_target = FileDropTarget(self.load_font)
        self.fontsList.SetDropTarget(file_drop_target)

        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.loadFont.Bind(wx.EVT_BUTTON, self.on_load_font)
        self.releaseFont.Bind(wx.EVT_BUTTON, self.on_release_font)
        self.releaseAll.Bind(wx.EVT_BUTTON, self.on_release_all)

        self.taskbar_icon = TaskBarIcon(self)
        self.temp_dir = tempfile.TemporaryDirectory()

        try:
            config = configparser.ConfigParser()
            config.read(r'./fontLoader.conf', encoding='utf-8')
            confirm = config.get('Settings', 'close')
            assert confirm == HIDE or confirm == DESTROY
            self.confirm = confirm
        except:
            self.confirm = CONFIRM

    def Destroy(self):
        self.release_font(reversed(range(self.fontsList.GetItemCount())))
        self.temp_dir.cleanup()
        self.taskbar_icon.Destroy()
        super().Destroy()

    def on_close(self, event):
        if self.confirm == CONFIRM:
            ConfirmFrame(self).Show()

        elif self.confirm == HIDE:
            self.Hide()

        elif self.confirm == DESTROY:
            self.Destroy()

    def on_load_font(self, event):
        file_filter = 'Font File (*.ttf,*.ttc,*.otf)|*.ttf;*.ttc;*.otf'
        file_dialog = wx.FileDialog(self, message='Choose a font file', wildcard=file_filter, style=wx.FD_MULTIPLE)
        status = file_dialog.ShowModal()
        if status != wx.ID_OK:
            return

        file_list = file_dialog.GetPaths()
        self.load_font(file_list)

    def on_release_font(self, event):
        item = self.fontsList.GetFirstSelected()
        if item == -1:
            return

        self.release_font([item])

    def on_release_all(self, event):
        self.release_font(reversed(range(self.fontsList.GetItemCount())))

    def load_font(self, font_list):
        for font_path in font_list:
            try:
                if font_path.endswith('.ttc'):
                    ttc_font_list = []
                    for font in TTCollection(font_path).fonts:
                        temp_font_path = os.path.join(self.temp_dir.name, str(uuid.uuid4()) + '.ttf')
                        font.save(temp_font_path)
                        ttc_font_list.append(temp_font_path)
                    self.load_font(ttc_font_list)
                    continue

                font = TTFont(font_path)
                font_name = font.get('name').getDebugName(6)

                if ctypes.windll.gdi32.AddFontResourceW(font_path) == 0:
                    raise RuntimeError('Failed to load {}'.format(font_path))

            except Exception:
                warning_box = wx.MessageDialog(None, 'Failed to load {}'.format(font_path), 'Error',
                                               wx.YES_DEFAULT | wx.ICON_QUESTION)
                warning_box.ShowModal()
                warning_box.Destroy()
                continue

            # ctypes.windll.user32.SendMessageW(HWND_BROADCAST, WM_FONTCHANGE, 0, 0, SMTO_ABORTIFHUNG, 1000, None)
            self.fontsList.Append([font_name, font_path])
            self.fontsList.SetItemFont(self.fontsList.GetItemCount() - 1, wx.Font(wx.FontInfo(10).FaceName(font_name)))

    def release_font(self, items):
        for item in items:
            font_name = self.fontsList.GetItemText(item, 0)
            font_path = self.fontsList.GetItemText(item, 1)

            if ctypes.windll.gdi32.RemoveFontResourceW(font_path) == 0:
                warning_box = wx.MessageDialog(None, 'Failed to release {}'.format(font_name), 'Error',
                                               wx.YES_DEFAULT | wx.ICON_QUESTION)
                warning_box.ShowModal()
                warning_box.Destroy()
                continue

            if font_path.startswith(self.temp_dir.name):
                os.remove(font_path)

            # ctypes.windll.user32.SendMessageW(HWND_BROADCAST, WM_FONTCHANGE, 0, 0, SMTO_ABORTIFHUNG, 1000, None)
            self.fontsList.DeleteItem(item)


class FileDropTarget(wx.FileDropTarget):
    def __init__(self, load_font):
        super().__init__()
        self.load_font = load_font

    def OnDropFiles(self, x, y, file_list):
        self.load_font(file_list)
        return True


class TaskBarIcon(wx.adv.TaskBarIcon):
    def __init__(self, main_frame):
        super(TaskBarIcon, self).__init__()
        self.main_frame = main_frame

        icon = wx.Icon(r'./icon.ico', wx.BITMAP_TYPE_ICO)
        self.SetIcon(icon, 'Font Loader')

        self.Bind(wx.adv.EVT_TASKBAR_LEFT_DOWN, self.on_left_down)
        self.Bind(wx.adv.EVT_TASKBAR_RIGHT_DOWN, self.on_right_down)

    def CreatePopupMenu(self):
        menu = wx.Menu()

        restore_item = menu.Append(wx.ID_ANY, 'Restore')
        add_font_item = menu.Append(wx.ID_ANY, 'Load Font')
        release_all_item = menu.Append(wx.ID_ANY, 'Release All')
        exit_item = menu.Append(wx.ID_ANY, 'Exit')

        self.Bind(wx.EVT_MENU, self.on_restore, restore_item)
        self.Bind(wx.EVT_MENU, self.main_frame.on_load_font, add_font_item)
        self.Bind(wx.EVT_MENU, self.main_frame.on_release_all, release_all_item)
        self.Bind(wx.EVT_MENU, self.on_exit, exit_item)

        return menu

    def on_left_down(self, event):
        self.main_frame.Show()
        self.main_frame.Restore()

    def on_right_down(self, event):
        self.PopupMenu(self.CreatePopupMenu())

    def on_restore(self, event):
        self.main_frame.Show()
        self.main_frame.Restore()

    def on_exit(self, event):
        self.main_frame.Destroy()
        self.Destroy()


class ConfirmFrame(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, title='Hint')
        self.parent = parent

        main_sizer = wx.BoxSizer(wx.VERTICAL)

        self.hint_label = wx.StaticText(self, label='Exit or hide it in the system tray?')
        main_sizer.Add(self.hint_label, 1, wx.ALL | wx.EXPAND, 5)

        self.remember = wx.CheckBox(self, label="Don't ask again")
        main_sizer.Add(self.remember, 1, wx.ALL | wx.EXPAND, 5)

        button_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.cancel = wx.Button(self, label='Cancel')
        button_sizer.Add(self.cancel, 1, wx.ALL | wx.EXPAND, 5)

        self.hide = wx.Button(self, label='Hide')
        button_sizer.Add(self.hide, 1, wx.ALL | wx.EXPAND, 5)

        self.exit = wx.Button(self, label='Exit')
        button_sizer.Add(self.exit, 1, wx.ALL | wx.EXPAND, 5)

        main_sizer.Add(button_sizer, 1, wx.ALL | wx.EXPAND, 5)

        self.SetSizer(main_sizer)
        self.Layout()

        self.Centre(wx.BOTH)

        self.cancel.Bind(wx.EVT_BUTTON, self.on_cancel)
        self.exit.Bind(wx.EVT_BUTTON, self.on_exit)
        self.hide.Bind(wx.EVT_BUTTON, self.on_hide)

    def on_cancel(self, event):
        self.Destroy()

    def on_exit(self, event):
        self.__remember_config(DESTROY)
        self.Destroy()
        self.parent.Destroy()

    def on_hide(self, event):
        self.__remember_config(HIDE)
        self.Destroy()
        self.parent.Hide()

    def __remember_config(self, key):
        if not self.remember.GetValue():
            return

        config = configparser.ConfigParser()
        config.add_section('Settings')

        config.set('Settings', 'close', key)
        self.parent.confirm = key

        with open(r'./fontLoader.conf', 'w', encoding='utf-8') as conf_file:
            config.write(conf_file)


if __name__ == '__main__':
    app = wx.App()
    MainFrame().Show()
    app.MainLoop()

# TODO CIL支持
# TODO 显示字体修复
# TODO ?
