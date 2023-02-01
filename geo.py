"""
Документация v1.0.3

Программа предназначена для геокодирования объектов недвижимости с помощью сервисов Yandex API.
Результатом являются координаты объекта и его нормализованный адрес.
Импользует метод загрузки csv файла с адресами (Test.csv) и выгрузки такого же csv файла (result.csv).
Предоставляет доступ к пользовательскому интерфейсу для отображения процесса загрузки.

- Функция distanceBetween вычисляет расстояние между двумя точками на карте
- Функция write_csv пишет данные полученые из геокодера в csv файл
- Функция read_csv получает данные из csv файла для дальнейщей обработки
- Функция geo_code основная функция выполнающая геокодирование адреса объекта
- Класс changeData используется для изменения стандартных значений
--- Функции onOk и onCancel обрабатывают события кнопок
- Класс Mywin используется для создания экранной формы
--- Функция InitUI содержит описания для каркаса формы
--- Функции startGEO и stopGEO запускают соответсвенно выполнение и остановку функции OnStart
--- Функция OnStart обрабатывает результат нажатия и запускает выполнение скрипта
--- Функция create_thread создает поток для работы OnStart
--- Функции GetNameFile и GetDataAPI обрабатывают изменения стандартных значений NAME_FILE_TO_GEO и API соответсвенно
"""

import wx
import time
import csv
import math
import threading
from geopy.geocoders import Yandex

import wx.lib.newevent
import sys
import traceback
import textwrap
import winsound

RANGE = 10
API = "0bef970e-5f6c-4b90-9ee6-665830d56159"
NAME_FILE_TO_GEO = "Test"

def show_error():
    message = ''.join(traceback.format_exception(*sys.exc_info()))
    dialog = wx.MessageDialog(None, message, 'Error!', wx.OK|wx.ICON_ERROR)
    dialog.ShowModal()

def distanceBetween(lat1, lon1, lat2, lon2):
    print("Начинаю вычисление расстояния")
    Rlat1 = (lat1 * 3.14) / 180
    Rlon1 = (lon1 * 3.14) / 180
    Rlat2 = (lat2 * 3.14) / 180
    Rlon2 = (lon2 * 3.14) / 180

    clat1 = math.cos(Rlat1)
    slat1 = math.sin(Rlat1)
    clat2 = math.cos(Rlat2)
    slat2 = math.sin(Rlat2)
    delta = Rlon2 - Rlon1
    cdelta = math.cos(delta)
    sdelta = math.sin(delta)

    pr1 = clat2 * sdelta
    pr2 = clat1 * slat2 - slat1 * clat2 * cdelta

    yy = math.sqrt(pow(pr1, 2) + pow(pr2, 2))
    xx = slat1 * slat2 + clat1 * clat2 * cdelta

    ad = math.atan2(yy, xx)
    result = ad * 6371
    result = round(result, 3)
    result = result * 1000
    print("Расстояние между точками равно =", result, "метров")
    return result

def read_csv():
    with open(NAME_FILE_TO_GEO + '.csv', 'r') as fp:
        reader = csv.reader(fp, delimiter=',', quotechar='"')
        A_D = [row for row in reader]
    return A_D

def write_csv(data):
    with open('result.csv', mode='a', encoding='cp1251', errors='replace') as f: #encoding='utf8'
        writer = csv.writer(f, delimiter=' ', lineterminator='\r')
        adr = data.pop('address')
        lt = data.pop('Lat')
        ln = data.pop('Lon')
        writer.writerow(("Адрес:", adr, 'Координаты:', lt, ln))

        listik = ['Нормализованный путь:']
        for itemKey, itemName in data.items():
            listik.append(itemKey)
            listik.append(':')
            listik.append(itemName)
            listik.append(',')
        writer.writerow((listik))
        writer.writerow('')

def geo_code(address):
    geolocator = Yandex(api_key=API, format_string="%s, Волгоградская область")
    location = geolocator.geocode(address, timeout=15)
    if location != None:
        # Координаты
        coordsLat = location.latitude
        coordsLon = location.longitude

        # Полный разбор
        RAW = location.raw
        getMetaData = RAW['metaDataProperty']
        getGeocoderData = getMetaData['GeocoderMetaData']
        full_address = getGeocoderData['text']
        getListAddress = getGeocoderData['Address']
        key = 'postal_code' in getListAddress
        postal_code = "Not found"
        if key != False:
            postal_code = getListAddress['postal_code']

        data = {'address': full_address, 'index': postal_code, 'Lat': coordsLat, 'Lon': coordsLon}

        getListComponents = getListAddress['Components']
        for getComponents in getListComponents:
            obj_type = getComponents['kind']
            obj_name = getComponents['name']
            data.setdefault(obj_type, obj_name)

    # Обработка исключений
    else:
        data = {'address': address, 'Lat': 'геокодирован', 'Lon': 'с ошибкой'}

    return data


class changeData(wx.Dialog):
    def __init__(self, parent, id=-1, title="Enter new value"):
        wx.Dialog.__init__(self, parent, id, title, size=(300, 150))

        self.mainSizer = wx.BoxSizer(wx.VERTICAL)
        self.buttonSizer = wx.BoxSizer(wx.HORIZONTAL)

        self.label = wx.StaticText(self, label="Enter value:")
        self.field = wx.TextCtrl(self, value="", size=(300, 20))
        self.okbutton = wx.Button(self, label="OK", id=wx.ID_OK)

        self.mainSizer.Add(self.label, 0, wx.ALL, 8 )
        self.mainSizer.Add(self.field, 0, wx.ALL, 8 )

        self.buttonSizer.Add(self.okbutton, 0, wx.ALL, 8 )

        self.mainSizer.Add(self.buttonSizer, 0, wx.ALL, 0)

        self.Bind(wx.EVT_BUTTON, self.onOK, id=wx.ID_OK)
        self.Bind(wx.EVT_TEXT_ENTER, self.onOK)

        self.SetSizer(self.mainSizer)
        self.result = None

    def onOK(self, event):
        self.result = self.field.GetValue()
        self.Destroy()

    def onCancel(self, event):
        self.result = None
        self.Destroy()



#///////////////////////////////////


class ExceptionDialog(wx.Dialog):
    """This class displays an error dialog with details information about the
    input exception, including a traceback."""

    def __init__(self, exception_type, exception, tb):
        wx.Dialog.__init__(self, None, -1, title="Unhandled error",
                           style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)

        self.SetSize((640, 480))
        self.SetMinSize((420, 200))

        self.exception = (exception_type, exception, tb)
        self.initialize_ui()

        winsound.MessageBeep(winsound.MB_ICONHAND)

    def initialize_ui(self):
        extype, exception, tb = self.exception

        panel = wx.Panel(self, -1)

        # Create the top row, containing the error icon and text message.
        top_row_sizer = wx.BoxSizer(wx.HORIZONTAL)

        error_bitmap = wx.ArtProvider.GetBitmap(
            wx.ART_ERROR, wx.ART_MESSAGE_BOX
        )
        error_bitmap_ctrl = wx.StaticBitmap(panel, -1)
        error_bitmap_ctrl.SetBitmap(error_bitmap)

        message_text = textwrap.dedent("""\
            I'm afraid there has been an unhandled error. Please send the
            contents of the text control below to the application developer.\
        """)
        message_label = wx.StaticText(panel, -1, message_text)

        top_row_sizer.Add(error_bitmap_ctrl, flag=wx.ALL, border=10)
        top_row_sizer.Add(message_label, flag=wx.ALIGN_CENTER_VERTICAL)

        # Create the text control with the error information.
        exception_info_text = textwrap.dedent("""\
            Exception type: {}

            Exception: {}

            Traceback:
            {}\
        """)
        exception_info_text = exception_info_text.format(
            extype, exception, ''.join(traceback.format_tb(tb))
        )

        text_ctrl = wx.TextCtrl(panel, -1,
                                style=wx.TE_MULTILINE | wx.TE_DONTWRAP)
        text_ctrl.SetValue(exception_info_text)

        # Create the OK button in the bottom row.
        ok_button = wx.Button(panel, -1, 'OK')
        self.Bind(wx.EVT_BUTTON, self.on_ok, source=ok_button)
        ok_button.SetFocus()
        ok_button.SetDefault()

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(top_row_sizer)
        # sizer.Add(message_label, flag=wx.ALL | wx.EXPAND, border=10)
        sizer.Add(text_ctrl, proportion=1, flag=wx.EXPAND)
        sizer.Add(ok_button, flag=wx.ALIGN_CENTER | wx.ALL, border=5)

        panel.SetSizer(sizer)

    def on_ok(self, event):
        self.Destroy()

#////////////////////////////////////
def custom_excepthook(exception_type, value, tb):
    dialog = ExceptionDialog(exception_type, value, tb)
    dialog.ShowModal()

class Mywin(wx.Frame):

    def __init__(self, parent, title):
        super(Mywin, self).__init__(parent, title=title, size=(300, 200))
        self.InitUI()

        ico = wx.Icon('py.ico', wx.BITMAP_TYPE_ICO)
        self.SetIcon(ico)

    def InitUI(self):
        self.count = 0
        pnl = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        hbox3 = wx.BoxSizer(wx.HORIZONTAL)
        hbox4 = wx.BoxSizer(wx.HORIZONTAL)
        hbox5 = wx.BoxSizer(wx.HORIZONTAL)

        LISST = read_csv()
        RANGE = len(LISST)

        self.gauge = wx.Gauge(pnl, range=RANGE, size=(350, 25), style=wx.GA_HORIZONTAL)
        self.btn1 = wx.Button(pnl, label="Start")
        self.Bind(wx.EVT_BUTTON, self.startGEO, self.btn1)
        self.btn4 = wx.Button(pnl, label="Stop")
        self.Bind(wx.EVT_BUTTON, self.stopGEO, self.btn4)
        self.text = wx.StaticText(pnl, label='press "Start" to geocode')

        self.text2 = wx.StaticText(pnl, label='Default = '+ NAME_FILE_TO_GEO)
        self.btn2 = wx.Button(pnl, label="Change")
        self.Bind(wx.EVT_BUTTON, self.GetNameFile, self.btn2)

        self.text3 = wx.StaticText(pnl, label='Default API key')
        self.btn3 = wx.Button(pnl, label="Change")
        self.Bind(wx.EVT_BUTTON, self.GetDataAPI, self.btn3)

        hbox1.Add(self.gauge, proportion=1, flag=wx.ALIGN_CENTRE)
        hbox2.Add(self.btn1, proportion=1, flag=wx.RIGHT, border=10)
        hbox2.Add(self.btn4, proportion=1, flag=wx.RIGHT, border=10)
        hbox3.Add(self.text, proportion=1, flag=wx.RIGHT, border=10)
        hbox4.Add(self.text2, proportion=1, flag=wx.RIGHT, border=10)
        hbox4.Add(self.btn2, proportion=1, flag=wx.RIGHT, border=10)
        hbox5.Add(self.text3, proportion=1, flag=wx.RIGHT, border=10)
        hbox5.Add(self.btn3, proportion=1, flag=wx.RIGHT, border=10)

        vbox.Add((0, 30))
        vbox.Add(hbox1, flag=wx.ALIGN_CENTRE)
        vbox.Add((0, 20))
        vbox.Add(hbox2, proportion=1, flag=wx.ALIGN_CENTRE)
        vbox.Add((0, 10))
        vbox.Add(hbox3, proportion=1, flag=wx.ALIGN_CENTRE)
        vbox.Add((0, 20))
        vbox.Add(hbox4, proportion=1, flag=wx.ALIGN_CENTRE)
        vbox.Add((0, 20))
        vbox.Add(hbox5, proportion=1, flag=wx.ALIGN_CENTRE)
        pnl.SetSizer(vbox)

        self.SetSize((400, 300))
        self.Centre()
        self.stop = False
        self.Show(True)

    def custom_excepthook(exception_type, value, tb):
        dialog = ExceptionDialog(exception_type, value, tb)
        dialog.ShowModal()

    def create_thread(self, target):
        thread = threading.Thread(target=target)
        thread.daemon = True
        thread.start()

    def GetNameFile(self, e):
        dlg = changeData(self)
        dlg.ShowModal()
        global NAME_FILE_TO_GEO
        NAME_FILE_TO_GEO = dlg.result
        address_List = read_csv()
        global RANGE
        RANGE = len(address_List)
        self.gauge.SetRange(RANGE)
        self.text2.SetLabel(dlg.result)

    def GetDataAPI(self, e):
        dlg = changeData(self)
        dlg.ShowModal()
        global API
        API = dlg.result
        self.text3.SetLabel(dlg.result)

    def OnStart(self):
        address_List = read_csv()
        RANGE = len(address_List)
        self.text.SetLabel('Geocoding in progress')

        for address in address_List:
            if self.stop:
                break
            data = geo_code(address)
            write_csv(data)
            time.sleep(1)
            self.count = self.count + 1
            self.gauge.SetValue(self.count)
            if self.count >= RANGE:
                self.text.SetLabel('  Geocoding complete')
                return
        self.stop = False


    def startGEO(self, e):
        self.create_thread(self.OnStart)

    def stopGEO(self, e):
        self.stop = True
        self.text.SetLabel('Geocoding is stopped')

    def OnCloseWindow(self, e):
        self.Destroy()

def main():
    ex = wx.App()
    try:
        Mywin(None, 'GEO')
        ex.MainLoop()
    except:
        show_error()



if __name__ == '__main__':
    main()