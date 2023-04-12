import os
import re
import subprocess
import pythoncom
import json
import qrcode
from win32com.client import Dispatch, gencache
from tkinter import Tk
from tkinter.filedialog import askopenfilenames

# Просмотр всех ячеек
def parse_stamp(doc7, number_sheet):
    stamp = doc7.LayoutSheets.Item(number_sheet).Stamp
    for i in range(10000):
        if stamp.Text(i).Str:
            print('Номер ячейки = %-5d Значение = %s' % (i, stamp.Text(i).Str))

# Подключение к API7 программы Компас 3D
def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    return module, api, const


# Функция проверки, запущена-ли программа КОМПАС 3D
def is_running():
    proc_list = \
    subprocess.Popen('tasklist /NH /FI "IMAGENAME eq KOMPAS*"', shell=False, stdout=subprocess.PIPE).communicate()[0]
    return True if proc_list else False


# Посчитаем количество листов каждого из формата
def amount_sheet(doc7):
    sheets = {"A0": 0, "A1": 0, "A2": 0, "A3": 0, "A4": 0, "A5": 0}
    for sheet in range(doc7.LayoutSheets.Count):
        format = doc7.LayoutSheets.Item(sheet).Format  # sheet - номер листа, отсчёт начинается от 0
        sheets["A" + str(format.Format)] += 1 * format.FormatMultiplicity
    return sheets


# Прочитаем основную надпись чертежа
def stamp(doc7):
    for sheet in range(doc7.LayoutSheets.Count):
        style_filename = os.path.basename(doc7.LayoutSheets.Item(sheet).LayoutLibraryFileName)
        style_number = int(doc7.LayoutSheets.Item(sheet).LayoutStyleNumber)

        if style_filename in ['graphic.lyt', 'Graphic.lyt'] and style_number == 1:
            stamp = doc7.LayoutSheets.Item(sheet).Stamp
            return {"Scale": re.findall(r"\d+:\d+", stamp.Text(6).Str)[0],
                    "Designer": stamp.Text(110).Str}

    return 'Неопределенный стиль оформления'


# Подсчет технических требований, в том случае, если включена автоматическая нумерация
def count_demand(doc7, module7):
    IDrawingDocument = doc7._oleobj_.QueryInterface(module7.NamesToIIDMap['IDrawingDocument'], pythoncom.IID_IDispatch)
    drawing_doc = module7.IDrawingDocument(IDrawingDocument)
    text_demand = drawing_doc.TechnicalDemand.Text

    count = 0  # Количество пунктов технических требований
    for i in range(text_demand.Count):  # Прохоим по каждой строчке технических требований
        if text_demand.TextLines[i].Numbering == 1:  # и проверяем, есть ли у строки нумерация
            count += 1

    # Если нет нумерации, но есть текст
    if not count and text_demand.TextLines[0]:
        count += 1

    return count

def kompass_str_refresh(text):
    text = text.replace('@54~', 'Δ')
    text = text.replace('^(Symbol type A)+61508~', 'Δ')
    text = text.replace('^(Symbol type A)+61549~', 'μ')
    text = text.replace('@2~', 'Ø')
    text = text.replace('$m;', '')
    text = text.replace('$', '')
    text = text.replace('\'', '')
    return text

def parse_data(doc7, module7):
    IKompasDocument2D = doc7._oleobj_.QueryInterface(module7.NamesToIIDMap['IKompasDocument2D'],
                                                     pythoncom.IID_IDispatch)
    doc2D = module7.IKompasDocument2D(IKompasDocument2D)
    views = doc2D.ViewsAndLayersManager.Views
    print('Количество Видов ',views.Count)
    count_dim = 0
    for i in range(views.Count):
        try:

            ISymbols2DContainer = views.View(i)._oleobj_.QueryInterface(module7.NamesToIIDMap['ISymbols2DContainer'],
                                                                        pythoncom.IID_IDispatch)
            dimensions = module7.ISymbols2DContainer(ISymbols2DContainer)
            # print(dir(dimensions.BreakRadialDimensions))
            # print(dir(dimensions.BreakRadialDimensions.BreakRadialDimension(1)))
            try:
                print(dir(dimensions.BreakRadialDimensions.BreakRadialDimension(0)))
                radial_count = dimensions.BreakRadialDimensions.Count
                radius_list = []
                for i in range(0, radial_count):
                    radius = dimensions.BreakRadialDimensions.BreakRadialDimension(i).Radius
                    radius_list.append(radius)

                    print('Radius ',i+1,' ',radius )

                radial_count = dimensions.RadialDimensions.Count
                for i in range(0, radial_count):
                    print(dir(dimensions.RadialDimensions.RadialDimension(0)))
                    radius = dimensions.RadialDimensions.RadialDimension(i).Radius
                    radius_list.append(radius)
                    print(radius_list)
                    print('Radius ',i+1,' ',radius )
            except:
                pass

            print('-----------------------------------------------------------------------------------------------------------')
            try:
                drawing_table = dimensions.DrawingTables.DrawingTable(0)
                # print(dimensions.DrawingTables.Count)
                table = module7.ITable(drawing_table)
                pair_list = []
                for i in range(1,1000, 2):
                    try:
                        if i%1 == 0:
                            print(i//2 + 1,end=' ')
                        table_cell1 = table.CellById(i)
                        text1 = module7.IText(table_cell1.Text)
                        text1 = kompass_str_refresh(text1.Str)

                        table_cell2 = table.CellById(i + 1)
                        text2 = module7.IText(table_cell2.Text)
                        text2 = kompass_str_refresh(text2.Str)

                        pair = (text1,text2)
                        pair_list.append(pair)
                        # print(pair)
                    except:
                        break
            except:
                pass
            result = {
                'table': pair_list,
                'radius':radius_list
            }
        except:
            print('Break on ', i)
            pass
    return result



# Подсчёт размеров на чертеже, для каждого вида по отдельности
def count_dimension(doc7, module7):
    IKompasDocument2D = doc7._oleobj_.QueryInterface(module7.NamesToIIDMap['IKompasDocument2D'],
                                                     pythoncom.IID_IDispatch)
    doc2D = module7.IKompasDocument2D(IKompasDocument2D)
    views = doc2D.ViewsAndLayersManager.Views

    count_dim = 0
    for i in range(views.Count):
        ISymbols2DContainer = views.View(i)._oleobj_.QueryInterface(module7.NamesToIIDMap['ISymbols2DContainer'],
                                                                    pythoncom.IID_IDispatch)
        dimensions = module7.ISymbols2DContainer(ISymbols2DContainer)

        print('Просто Радиусы ',dimensions.RadialDimensions.Count)
        count_dim += dimensions.AngleDimensions.Count + \
                     dimensions.ArcDimensions.Count + \
                     dimensions.Bases.Count + \
                     dimensions.BreakLineDimensions.Count + \
                     dimensions.BreakRadialDimensions.Count + \
                     dimensions.DiametralDimensions.Count + \
                     dimensions.Leaders.Count + \
                     dimensions.LineDimensions.Count + \
                     dimensions.RadialDimensions.Count + \
                     dimensions.RemoteElements.Count + \
                     dimensions.Roughs.Count + \
                     dimensions.Tolerances.Count

    return count_dim


def parse_design_documents(paths):
    is_run = is_running()  # True, если программа Компас уже запущена

    module7, api7, const7 = get_kompas_api7()  # Подключаемся к программе
    app7 = api7.Application  # Получаем основной интерфейс программы
    app7.Visible = True  # Показываем окно пользователю (если скрыто)
    app7.HideMessage = const7.ksHideMessageNo  # Отвечаем НЕТ на любые вопросы программы

    table = []  # Создаём таблицу парметров
    for path in paths:

        doc7 = app7.Documents.Open(PathName=path,
                                   Visible=True,
                                   ReadOnly=True)  # Откроем файл в видимом режиме без права его изменять

        # parse_stamp(doc7,0)
        # row = amount_sheet(doc7)  # Посчитаем кол-во листов каждого формат
        row = stamp(doc7)  # Читаем основную надпись
        # print(dir(doc7))
        row.update({
            "Filename": doc7.Name,  # Имя файла
            "CountTD": count_demand(doc7, module7),  # Количество пунктов технических требований
            "CountDim": count_dimension(doc7, module7), # Количество пунктов технических требований
            'Data': parse_data(doc7, module7)
        })
        table.append(row)  # Добавляем строку параметров в таблицу

        # doc7.Close(const7.kdDoNotSaveChanges)  # Закроем файл без изменения

    if not is_run: app7.Quit()  # Закрываем программу при необходимости
    return table

def QR_stamp(data_str):
    imgname = 'QR.png'
    # generate qr code
    img = qrcode.make(data_str)
    # save img to a file
    img.save(imgname)

    kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    #  Подключим описание интерфейсов API5
    kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
    kompas_object = kompas6_api5_module.KompasObject(
        Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    iDocument2D = kompas_object.ActiveDocument2D()
    iRasterParam = kompas6_api5_module.ksRasterParam(kompas_object.GetParamStruct(kompas6_constants.ko_RasterParam))
    iRasterParam.Init()
    iRasterParam.embeded = True
    imgName = os.getcwd() + '\\' + imgname
    # print(imgName)
    iRasterParam.fileName = imgName
    iPlacementParam = kompas6_api5_module.ksPlacementParam(
        kompas_object.GetParamStruct(kompas6_constants.ko_PlacementParam))
    iPlacementParam.Init()
    iPlacementParam.angle = 0
    iPlacementParam.scale_ = 0.5
    iPlacementParam.xBase = 0
    iPlacementParam.yBase = 0
    iRasterParam.SetPlace(iPlacementParam)
    iDocument2D.ksInsertRaster(iRasterParam)

if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Скрываем основное окно и сразу окно выбора файлов

    filenames = askopenfilenames(title="Выберети чертежи деталей", filetypes=[('Компас 3D', '*.cdw'), ])
    # filenames = ('C:/Users/KonBas/PycharmProjects/KOMPAS/Чертеж Линзы КП.0123.3301.112.cdw',)
    data = parse_design_documents(filenames)
    data = data[0]['Data']
    # print(data)

    data_str = ''
    dist = 12
    line = '-'*(dist*2+1) + '\n'
    data_str += line
    data_str += 'Radius' + '\n'
    data_str += line

    for i in range(0, len(data['radius'])):
        data_str += f'R{i} ' + str(round(data['radius'][i], 2)) + '\n'
    data_str += line
    data_str += 'TABLE' + '\n'
    data_str += line
    for i in range(0, len(data['table'])):
        # data_str += (f'{i}.'+data['table'][i][0]).ljust(dist) + '|' + data['table'][i][1].ljust(dist) + '\n'
        data_str += (f'{i}.' + data['table'][i][0])+ '  ' + data['table'][i][1] + '\n'
    data_str += line
    print(data_str)
    QR_stamp(data_str)



    # root.destroy()  # Уничтожаем основное окно
    # root.mainloop()