Attribute VB_Name = "Main"
Option Explicit
Option Base 1
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890
Public Const CRC_HOST As Integer = 127, REV As Integer = &H137
Public Const PERSON_LIST = "Ф/Л,Ю/Л" ' Не менять! "*/Л,Ф/Л,Ю/Л"

' Коллекции: ключи коллекции xID, рабочие листы и колонки
Private xID As New Collection
' Коллекция: реквизиты и контакты контрагентов, цены
Private xSUPP As New Collection
' Константы модуля
Private Const QT As String = ": Количество "
' Динамический массив для сбора данных
Private Table() As Variant

Static Sub Main_Sub(ByVal DateBegin As Date, ByVal DateEnd As Variant)
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
  Debug.Print GetRecord(0, "sheet", 1)
  Debug.Print DateSerial(IIf(Month(DateAdd("m", -6, DateEnd)) < 12, _
    Year(DateAdd("m", -6, DateEnd)), Year(DateAdd("m", -6, DateEnd)) + 1), _
    IIf(Month(DateEnd) - 5 < 1, 12, 0) + Month(DateEnd) - 5, 0) ' Для заголовка
  Stop
  
  With Application
    .Calculation = xlAutomatic
    .EnableEvents = True: .ScreenUpdating = True
  End With
End Sub
