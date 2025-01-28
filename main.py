# coding: utf-8
from wx import (App, Frame, Dialog, MessageDialog, FileDialog, Menu, FileDropTarget,
                BoxSizer, ListCtrl, Button, CheckBox, StaticText,
                Colour)

from wx import (VERTICAL, HORIZONTAL, LC_REPORT, BORDER_SUNKEN, FD_MULTIPLE, YES_DEFAULT,
                ALL, EXPAND, ALIGN_CENTER_HORIZONTAL, LIST_AUTOSIZE_USEHEADER,
                MINIMIZE_BOX, DEFAULT_FRAME_STYLE, MAXIMIZE_BOX,
                BOTH,
                EVT_BUTTON, EVT_CLOSE, EVT_MENU,
                ICON_QUESTION,
                ID_OK, ID_ANY)

from wx.lib.embeddedimage import PyEmbeddedImage
from ctypes import WinDLL
from tempfile import TemporaryDirectory
from uuid import uuid4
from os import remove, startfile
from os.path import join
from wx.adv import TaskBarIcon, EVT_TASKBAR_LEFT_DOWN, EVT_TASKBAR_RIGHT_DOWN
from configparser import ConfigParser
from fontTools.ttLib.ttFont import TTFont
from fontTools.ttLib import TTCollection
from argparse import ArgumentParser
from icons.icon_data import ICON_BASE64

# HWND_BROADCAST = 0xFFFF
# SMTO_ABORTIFHUNG = 0x0002
# WM_FONTCHANGE = 0x001D

CONFIRM = 'confirm'
HIDE = 'hide'
DESTROY = 'destroy'

gdi32 = WinDLL("gdi32.dll")
# user32 = WinDLL("user32.dll")


class MainFrame(Frame):
    def __init__(self):
        super().__init__(None, title='Font Loader', size=(330, 500))
        self.SetBackgroundColour(Colour("WHITE"))
        self.SetIcon(PyEmbeddedImage(ICON_BASE64).getIcon())

        data_sizer = BoxSizer(VERTICAL)

        self.fontsList = ListCtrl(self, style=LC_REPORT | BORDER_SUNKEN)
        self.fontsList.AppendColumn('Font Name')
        self.fontsList.AppendColumn('File Path', width=LIST_AUTOSIZE_USEHEADER)
        data_sizer.Add(self.fontsList, 1, ALL | EXPAND, 5)

        control_sizer = BoxSizer(HORIZONTAL)

        self.loadFont = Button(self, label='Load Font')
        control_sizer.Add(self.loadFont, 0, ALL | EXPAND, 5)

        self.releaseFont = Button(self, label='Release Font')
        control_sizer.Add(self.releaseFont, 0, ALL | EXPAND, 5)

        self.releaseAll = Button(self, label='Release All')
        control_sizer.Add(self.releaseAll, 0, ALL | EXPAND, 5)

        data_sizer.Add(control_sizer, 0, ALL | ALIGN_CENTER_HORIZONTAL, 5)

        self.SetSizer(data_sizer)
        self.Layout()

        self.Centre(BOTH)

        file_drop_target = FontFileDropTarget(self.load_font)
        self.fontsList.SetDropTarget(file_drop_target)

        self.Bind(EVT_CLOSE, self.on_close)
        self.loadFont.Bind(EVT_BUTTON, self.on_load_font)
        self.releaseFont.Bind(EVT_BUTTON, self.on_release_font)
        self.releaseAll.Bind(EVT_BUTTON, self.on_release_all)

        self.taskbar_icon = FontTaskBarIcon(self)
        self.temp_dir = TemporaryDirectory()

        try:
            config = ConfigParser()
            config.read(r'./fontLoader.conf', encoding='utf-8')
            confirm = config.get('Settings', 'close')
            assert confirm == HIDE or confirm == DESTROY
            self.confirm = confirm
        except:
            self.confirm = CONFIRM

    def Destroy(self):
        try:
            self.release_font(reversed(range(self.fontsList.GetItemCount())))
            self.temp_dir.cleanup()
        except:
            warning_box = MessageDialog(None, "Failed to release all the fonts\n"
                                              "You'd better clean up temp directory \n{}\nby yourself later."
                                        .format(self.temp_dir.name),
                                        'Error', YES_DEFAULT | ICON_QUESTION)
            warning_box.ShowModal()
            warning_box.Destroy()
            startfile(self.temp_dir.name)
        finally:
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
        file_dialog = FileDialog(self, message='Choose a font file', wildcard=file_filter, style=FD_MULTIPLE)
        status = file_dialog.ShowModal()
        if status != ID_OK:
            return

        file_list = file_dialog.GetPaths()
        self.load_font(file_list)

    def on_release_font(self, event):
        selected_item = []
        item = self.fontsList.GetFirstSelected()
        while item != -1:
            selected_item.append(item)
            item = self.fontsList.GetNextSelected(item)

        self.release_font(reversed(selected_item))

    def on_release_all(self, event):
        self.release_font(reversed(range(self.fontsList.GetItemCount())))

    def load_font(self, font_list):
        for font_path in font_list:
            font_name = 'Unknown'
            try:
                if font_path.endswith('.ttc'):
                    ttc_font_list = []
                    for font in TTCollection(font_path).fonts:
                        temp_font_path = join(self.temp_dir.name, str(uuid4()) + '.ttf')
                        font.save(temp_font_path)
                        ttc_font_list.append(temp_font_path)
                    self.load_font(ttc_font_list)
                    continue

                font = TTFont(font_path)
                font_name = font.get('name').getDebugName(6)

                if gdi32.AddFontResourceW(font_path) == 0:
                    raise RuntimeError('Failed to load {}'.format(font_path))

            except Exception:
                warning_box = MessageDialog(None, 'Failed to load {}\nFont Path:\n{}'.format(font_name, font_path),
                                            'Error', YES_DEFAULT | ICON_QUESTION)
                warning_box.ShowModal()
                warning_box.Destroy()
                continue

            # user32.SendMessageW(HWND_BROADCAST, WM_FONTCHANGE, 0, 0, SMTO_ABORTIFHUNG, 1000, None)
            self.fontsList.Append([font_name, font_path])

    def release_font(self, items):
        for item in items:
            font_name = self.fontsList.GetItemText(item, 0)
            font_path = self.fontsList.GetItemText(item, 1)

            try:
                if gdi32.RemoveFontResourceW(font_path) == 0:
                    raise RuntimeError('Failed to load {}'.format(font_path))

                if font_path.startswith(self.temp_dir.name):
                    remove(font_path)

            except Exception:
                warning_box = MessageDialog(None, 'Failed to release {}\nFont Path:\n{}'.format(font_name, font_path),
                                            'Error', YES_DEFAULT | ICON_QUESTION)
                warning_box.ShowModal()
                warning_box.Destroy()
                continue

            # user32.SendMessageW(HWND_BROADCAST, WM_FONTCHANGE, 0, 0, SMTO_ABORTIFHUNG, 1000, None)
            self.fontsList.DeleteItem(item)


class FontFileDropTarget(FileDropTarget):
    def __init__(self, load_font):
        super().__init__()
        self.load_font = load_font

    def OnDropFiles(self, x, y, file_list):
        self.load_font(file_list)
        return True


class FontTaskBarIcon(TaskBarIcon):
    def __init__(self, main_frame):
        super(FontTaskBarIcon, self).__init__()
        self.main_frame = main_frame

        self.SetIcon(PyEmbeddedImage(ICON_BASE64).getIcon(), 'Font Loader')

        self.Bind(EVT_TASKBAR_LEFT_DOWN, self.on_left_down)
        self.Bind(EVT_TASKBAR_RIGHT_DOWN, self.on_right_down)

    def CreatePopupMenu(self):
        menu = Menu()

        restore_item = menu.Append(ID_ANY, 'Restore')
        add_font_item = menu.Append(ID_ANY, 'Load Font')
        release_all_item = menu.Append(ID_ANY, 'Release All')
        exit_item = menu.Append(ID_ANY, 'Exit')

        self.Bind(EVT_MENU, self.on_restore, restore_item)
        self.Bind(EVT_MENU, self.main_frame.on_load_font, add_font_item)
        self.Bind(EVT_MENU, self.main_frame.on_release_all, release_all_item)
        self.Bind(EVT_MENU, self.on_exit, exit_item)

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


class ConfirmFrame(Dialog):
    def __init__(self, parent):
        super().__init__(parent, title='Hint', size=(350, 150),
                         style=DEFAULT_FRAME_STYLE ^ MAXIMIZE_BOX ^ MINIMIZE_BOX)
        self.SetMaxSize((350, 150))
        self.SetMinSize((350, 150))
        self.SetBackgroundColour(Colour("WHITE"))
        self.SetIcon(PyEmbeddedImage(ICON_BASE64).getIcon())

        self.parent = parent
        self.parent.Enable(False)

        main_sizer = BoxSizer(VERTICAL)

        self.hint_label = StaticText(self, label='Exit or hide it in the task bar?')

        hint_font = self.hint_label.GetFont()
        hint_font.SetPointSize(15)
        self.hint_label.SetFont(hint_font)

        main_sizer.Add(self.hint_label, 1, ALL | EXPAND, 5)

        self.remember = CheckBox(self, label="Don't ask again")
        main_sizer.Add(self.remember, 1, ALL | EXPAND, 5)

        button_sizer = BoxSizer(HORIZONTAL)

        self.cancel = Button(self, label='Cancel')
        button_sizer.Add(self.cancel, 1, ALL, 5)

        self.hide = Button(self, label='Hide')
        button_sizer.Add(self.hide, 1, ALL, 5)

        self.exit = Button(self, label='Exit')
        button_sizer.Add(self.exit, 1, ALL, 5)

        main_sizer.Add(button_sizer, 1, ALL | ALIGN_CENTER_HORIZONTAL, 5)

        self.SetSizer(main_sizer)
        self.Layout()

        self.Centre(BOTH)

        self.Bind(EVT_CLOSE, self.on_cancel)
        self.cancel.Bind(EVT_BUTTON, self.on_cancel)
        self.exit.Bind(EVT_BUTTON, self.on_exit)
        self.hide.Bind(EVT_BUTTON, self.on_hide)

    def on_cancel(self, event):
        self.parent.Enable(True)
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

        config = ConfigParser()
        config.add_section('Settings')

        config.set('Settings', 'close', key)
        self.parent.confirm = key

        with open(r'./fontLoader.conf', 'w', encoding='utf-8') as conf_file:
            config.write(conf_file)


if __name__ == '__main__':
    parser = ArgumentParser(description='Load font temply to system.')

    parser.add_argument('font_list', nargs='*', help='Font list which will be loaded')
    parser.add_argument('-D', '--display', help="Display main window", action='store_true')
    args = parser.parse_args()

    app = App()
    main_frame = MainFrame()

    if len(args.font_list) != 0:
        main_frame.load_font(args.font_list)
        if args.display:
            main_frame.Show()

    else:
        main_frame.Show()

    app.MainLoop()
