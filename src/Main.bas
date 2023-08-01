Attribute VB_Name = "Main"
Option Explicit
Option Base 1
'123456789012345678901234567890123456h8nor@уа56789012345678901234567890123456789

Public Const WORKBOOKS_FILTER$ = "*", CRC_HOST% = 256, REV% = &H13A
Public Const IS_DEBUG = False, PERSON_LIST = "Ф/Л,Ю/Л" ' Не менять! "*/Л,Ф/Л,Ю/Л"

' Коллекции: ключи коллекции xID, рабочие листы и колонки
Public xID As New Collection
' Коллекция: реквизиты и контакты контрагентов, цены
Public xSUPP As New Collection
' Константы модуля
Private Const QT As String = ": Количество "
' Динамический массив для сбора данных
Private Table() As Variant

Static Sub Main_Sub(ByVal dateBegin As Date, ByVal dateEnd As Variant)
  Attribute Main_Sub.VB_Description = "r316 ¦ Сбор данных"
  Dim tBegin As Double, eZ As Byte, nZ As Integer
  
  With Application
    .ReferenceStyle = xlA1 ' Абсолютные ссылки
    .ScreenUpdating = False: .EnableEvents = False
    .Calculation = xlCalculationManual
  End With: tBegin = Timer
  
  ' Создаём динамический двумерный массив для выгрузки данных
  'ReDim Table(1, 1) As Variant
  
  GetBanks xID, xSUPP
  If xID("key").Count = 0 Then HookMsg Format(Timer - tBegin, "0.000 с:") & _
    "ОШИБКА! Ни один Банк не найден!", vbRetryCancel: _
    Application.Cursor = xlDefault: End ' HotFix!
  
  HookMsg GetRecord(0, "sheet", 1), vbRetryCancel
  HookMsg DateSerial(IIf(Month(DateAdd("m", -6, dateEnd)) < 12, _
    Year(DateAdd("m", -6, dateEnd)), Year(DateAdd("m", -6, dateEnd)) + 1), _
    IIf(Month(dateEnd) - 5 < 1, 12, 0) + Month(dateEnd) - 5, 0), vbRetryCancel
  Stop
  
  With Application
    .Calculation = xlCalculationAutomatic: .Cursor = xlDefault
    .EnableEvents = True: .ScreenUpdating = True
  End With
End Sub
