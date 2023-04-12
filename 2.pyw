# -*- coding: utf-8 -*-

#  PyKompasMacro https://slaviationsoft.blogspot.com
#  КОМПАС-3D (18, 1, 55, 0)
#  PyKompasMacro.exe (1.6.45.92)
#  root 14-12-2022 16:22:36

import pythoncom
from win32com.client import Dispatch, gencache, VARIANT

#  Получи константы
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

#  Получи API интерфейсов версии 5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))

#  Получи API интерфейсов версии 7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
kompas_api_object = kompas_api7_module.IKompasAPIObject(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch))
application = kompas_api_object.Application

#  Получи интерфейс активного документа
kompas_document = application.ActiveDocument

#  Создай графический объект "Таблица"
kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
views_and_layers_manager = kompas_document_2d.ViewsAndLayersManager
views = views_and_layers_manager.Views
view = views.ActiveView
symbols_2d_container = kompas_api7_module.ISymbols2DContainer(view)
drawing_tables = symbols_2d_container.DrawingTables
drawing_table = drawing_tables.Add(4, 3, 10.0, 50.0, 1)
drawing_table.X = 106.7616738
drawing_table.Y = 234.4172838
drawing_table.Angle = 0.0
drawing_table.FixedCellsSize = False
drawing_table.FixedRowCount = False
drawing_table.FixedColumnCount = False

table = kompas_api7_module.ITable(drawing_table)

table_cell = table.CellById(1)
cell_format = kompas_api7_module.ICellFormat(table_cell)
cell_format.TextStyle = 9
cell_format.ReadOnly = False
cell_format.OneLine = False
cell_format.LeftEdge = 0.5
cell_format.RightEdge = -0.5
cell_format.SpaceBefore = 0.0
cell_format.SpaceAfter = 0.0
cell_format.Width = 50.0
cell_format.Height = 10.0
cell_format.HFormat = kompas6_constants.ksHFormatStrNarrowing
cell_format.VFormat = True

cell_boundaries = kompas_api7_module.ICellBoundaries(table_cell)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBLeftBorder, kompas6_constants.ksCSNormal)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBRightBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBTopBorder, kompas6_constants.ksCSNormal)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBBottomBorder, kompas6_constants.ksCSNormal)

text = kompas_api7_module.IText(table_cell.Text)
text.Str = "Первый столбец"


table_cell = table.CellById(2)
cell_format = kompas_api7_module.ICellFormat(table_cell)
cell_format.TextStyle = 9
cell_format.ReadOnly = False
cell_format.OneLine = False
cell_format.LeftEdge = 0.5
cell_format.RightEdge = -0.5
cell_format.SpaceBefore = 0.0
cell_format.SpaceAfter = 0.0
cell_format.Width = 50.0
cell_format.Height = 10.0
cell_format.HFormat = kompas6_constants.ksHFormatStrNarrowing
cell_format.VFormat = True

cell_boundaries = kompas_api7_module.ICellBoundaries(table_cell)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBLeftBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBRightBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBTopBorder, kompas6_constants.ksCSNormal)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBBottomBorder, kompas6_constants.ksCSNormal)

text = kompas_api7_module.IText(table_cell.Text)
text.Str = "Второй столбец"


table_cell = table.CellById(3)
cell_format = kompas_api7_module.ICellFormat(table_cell)
cell_format.TextStyle = 9
cell_format.ReadOnly = False
cell_format.OneLine = False
cell_format.LeftEdge = 0.5
cell_format.RightEdge = -0.5
cell_format.SpaceBefore = 0.0
cell_format.SpaceAfter = 0.0
cell_format.Width = 50.0
cell_format.Height = 10.0
cell_format.HFormat = kompas6_constants.ksHFormatStrNarrowing
cell_format.VFormat = True

cell_boundaries = kompas_api7_module.ICellBoundaries(table_cell)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBLeftBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBRightBorder, kompas6_constants.ksCSNormal)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBTopBorder, kompas6_constants.ksCSNormal)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBBottomBorder, kompas6_constants.ksCSNormal)

text = kompas_api7_module.IText(table_cell.Text)
text.Str = "Третий столбец"


table_cell = table.CellById(4)
cell_format = kompas_api7_module.ICellFormat(table_cell)
cell_format.TextStyle = 10
cell_format.ReadOnly = False
cell_format.OneLine = False
cell_format.LeftEdge = 0.5
cell_format.RightEdge = -0.5
cell_format.SpaceBefore = 0.0
cell_format.SpaceAfter = 0.0
cell_format.Width = 50.0
cell_format.Height = 10.0
cell_format.HFormat = kompas6_constants.ksHFormatStrNarrowing
cell_format.VFormat = True

cell_boundaries = kompas_api7_module.ICellBoundaries(table_cell)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBLeftBorder, kompas6_constants.ksCSNormal)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBRightBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBTopBorder, kompas6_constants.ksCSNormal)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBBottomBorder, kompas6_constants.ksCSThin)

text = kompas_api7_module.IText(table_cell.Text)
text.Str = "АВС"


table_cell = table.CellById(5)
cell_format = kompas_api7_module.ICellFormat(table_cell)
cell_format.TextStyle = 10
cell_format.ReadOnly = False
cell_format.OneLine = False
cell_format.LeftEdge = 0.5
cell_format.RightEdge = -0.5
cell_format.SpaceBefore = 0.0
cell_format.SpaceAfter = 0.0
cell_format.Width = 50.0
cell_format.Height = 10.0
cell_format.HFormat = kompas6_constants.ksHFormatStrNarrowing
cell_format.VFormat = True

cell_boundaries = kompas_api7_module.ICellBoundaries(table_cell)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBLeftBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBRightBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBTopBorder, kompas6_constants.ksCSNormal)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBBottomBorder, kompas6_constants.ksCSThin)

text = kompas_api7_module.IText(table_cell.Text)
text.Str = "5"


table_cell = table.CellById(6)
cell_format = kompas_api7_module.ICellFormat(table_cell)
cell_format.TextStyle = 10
cell_format.ReadOnly = False
cell_format.OneLine = False
cell_format.LeftEdge = 0.5
cell_format.RightEdge = -0.5
cell_format.SpaceBefore = 0.0
cell_format.SpaceAfter = 0.0
cell_format.Width = 50.0
cell_format.Height = 10.0
cell_format.HFormat = kompas6_constants.ksHFormatStrNarrowing
cell_format.VFormat = True

cell_boundaries = kompas_api7_module.ICellBoundaries(table_cell)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBLeftBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBRightBorder, kompas6_constants.ksCSNormal)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBTopBorder, kompas6_constants.ksCSNormal)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBBottomBorder, kompas6_constants.ksCSThin)

text = kompas_api7_module.IText(table_cell.Text)
text.Str = "Рис.1"


table_cell = table.CellById(7)
cell_format = kompas_api7_module.ICellFormat(table_cell)
cell_format.TextStyle = 10
cell_format.ReadOnly = False
cell_format.OneLine = False
cell_format.LeftEdge = 0.5
cell_format.RightEdge = -0.5
cell_format.SpaceBefore = 0.0
cell_format.SpaceAfter = 0.0
cell_format.Width = 50.0
cell_format.Height = 10.0
cell_format.HFormat = kompas6_constants.ksHFormatStrNarrowing
cell_format.VFormat = True

cell_boundaries = kompas_api7_module.ICellBoundaries(table_cell)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBLeftBorder, kompas6_constants.ksCSNormal)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBRightBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBTopBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBBottomBorder, kompas6_constants.ksCSThin)

text = kompas_api7_module.IText(table_cell.Text)
text.Str = "АВС1"


table_cell = table.CellById(8)
cell_format = kompas_api7_module.ICellFormat(table_cell)
cell_format.TextStyle = 10
cell_format.ReadOnly = False
cell_format.OneLine = False
cell_format.LeftEdge = 0.5
cell_format.RightEdge = -0.5
cell_format.SpaceBefore = 0.0
cell_format.SpaceAfter = 0.0
cell_format.Width = 50.0
cell_format.Height = 10.0
cell_format.HFormat = kompas6_constants.ksHFormatStrNarrowing
cell_format.VFormat = True

cell_boundaries = kompas_api7_module.ICellBoundaries(table_cell)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBLeftBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBRightBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBTopBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBBottomBorder, kompas6_constants.ksCSThin)

text = kompas_api7_module.IText(table_cell.Text)
text.Str = "6"


table_cell = table.CellById(9)
cell_format = kompas_api7_module.ICellFormat(table_cell)
cell_format.TextStyle = 10
cell_format.ReadOnly = False
cell_format.OneLine = False
cell_format.LeftEdge = 0.5
cell_format.RightEdge = -0.5
cell_format.SpaceBefore = 0.0
cell_format.SpaceAfter = 0.0
cell_format.Width = 50.0
cell_format.Height = 10.0
cell_format.HFormat = kompas6_constants.ksHFormatStrNarrowing
cell_format.VFormat = True

cell_boundaries = kompas_api7_module.ICellBoundaries(table_cell)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBLeftBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBRightBorder, kompas6_constants.ksCSNormal)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBTopBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBBottomBorder, kompas6_constants.ksCSThin)

text = kompas_api7_module.IText(table_cell.Text)
text.Str = "Рис.2"


table_cell = table.CellById(10)
cell_format = kompas_api7_module.ICellFormat(table_cell)
cell_format.TextStyle = 10
cell_format.ReadOnly = False
cell_format.OneLine = False
cell_format.LeftEdge = 0.5
cell_format.RightEdge = -0.5
cell_format.SpaceBefore = 0.0
cell_format.SpaceAfter = 0.0
cell_format.Width = 50.0
cell_format.Height = 10.0
cell_format.HFormat = kompas6_constants.ksHFormatStrNarrowing
cell_format.VFormat = True

cell_boundaries = kompas_api7_module.ICellBoundaries(table_cell)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBLeftBorder, kompas6_constants.ksCSNormal)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBRightBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBTopBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBBottomBorder, kompas6_constants.ksCSNormal)

text = kompas_api7_module.IText(table_cell.Text)
text.Str = "АВС2"


table_cell = table.CellById(11)
cell_format = kompas_api7_module.ICellFormat(table_cell)
cell_format.TextStyle = 10
cell_format.ReadOnly = False
cell_format.OneLine = False
cell_format.LeftEdge = 0.5
cell_format.RightEdge = -0.5
cell_format.SpaceBefore = 0.0
cell_format.SpaceAfter = 0.0
cell_format.Width = 50.0
cell_format.Height = 10.0
cell_format.HFormat = kompas6_constants.ksHFormatStrNarrowing
cell_format.VFormat = True

cell_boundaries = kompas_api7_module.ICellBoundaries(table_cell)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBLeftBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBRightBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBTopBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBBottomBorder, kompas6_constants.ksCSNormal)

text = kompas_api7_module.IText(table_cell.Text)
text.Str = "7"


table_cell = table.CellById(12)
cell_format = kompas_api7_module.ICellFormat(table_cell)
cell_format.TextStyle = 10
cell_format.ReadOnly = False
cell_format.OneLine = False
cell_format.LeftEdge = 0.5
cell_format.RightEdge = -0.5
cell_format.SpaceBefore = 0.0
cell_format.SpaceAfter = 0.0
cell_format.Width = 50.0
cell_format.Height = 10.0
cell_format.HFormat = kompas6_constants.ksHFormatStrNarrowing
cell_format.VFormat = True

cell_boundaries = kompas_api7_module.ICellBoundaries(table_cell)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBLeftBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBRightBorder, kompas6_constants.ksCSNormal)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBTopBorder, kompas6_constants.ksCSThin)
cell_boundaries.SetLineStyle(kompas6_constants.ksCBBottomBorder, kompas6_constants.ksCSNormal)

text = kompas_api7_module.IText(table_cell.Text)
text.Str = "Рис.3"

drawing_table.Update()

