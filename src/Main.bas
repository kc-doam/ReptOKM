Attribute VB_Name = "Main"
Option Explicit
Option Base 1
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890

Public Const CRC_HOST As Integer = 127, REV As Integer = &H139
Public Const PERSON_LIST = "Ф/Л,Ю/Л" ' Не менять! "*/Л,Ф/Л,Ю/Л"

' Коллекции: ключи коллекции xID, рабочие листы и колонки
Public xID As New Collection
' Коллекция: реквизиты и контакты контрагентов, цены
Public xSUPP As New Collection
' Константы модуля
Private Const QT As String = ": Количество "
' Динамический массив для сбора данных
Private Table() As Variant

Static Sub Main_Sub(ByVal BeginDate As Date, ByVal EndDate As Variant)
  Attribute Main_Sub.VB_Description = "r310 ¦ Сбор данных"
  Dim tBegin As Double, eZ As Byte, nZ As Integer
  
  With Application
    .ReferenceStyle = xlA1 ' Абсолютные ссылки
    .ScreenUpdating = False: .EnableEvents = False
    .Calculation = xlCalculationManual
  End With: tBegin = Timer
  
  ' Создаём динамический двумерный массив для выгрузки данных
  'ReDim Table(1, 1) As Variant
  
  GetBanks xID, xSUPP
  If xID("key").Count = 0 Then Debug.Print Format(Timer - tBegin, "0.000 с:"), _
    "Ни один Банк не найден!": Application.Cursor = xlDefault: End ' HotFix!
  
  Debug.Print GetRecord(0, "sheet", 1)
  Debug.Print DateSerial(IIf(Month(DateAdd("m", -6, EndDate)) < 12, _
    Year(DateAdd("m", -6, EndDate)), Year(DateAdd("m", -6, EndDate)) + 1), _
    IIf(Month(EndDate) - 5 < 1, 12, 0) + Month(EndDate) - 5, 0) ' Для заголовка
  Stop
  
  With Application
    .Calculation = xlCalculationAutomatic: .Cursor = xlDefault
    .EnableEvents = True: .ScreenUpdating = True
  End With
End Sub
