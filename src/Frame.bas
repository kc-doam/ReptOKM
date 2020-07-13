Attribute VB_Name = "Frame"
Option Explicit
Option Base 1
Option Private Module
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890

' Коллекция с именами папок статистик
Public DirName As New Collection
' Коллекция файлов статистик менеджеров (индекс совпадает с DirName)
Public FileName As New Collection
' Коллекция с именами менеджеров (индекс совпадает с DirName)
Public Manager As New Collection

Private Enum DialogType
  dtDateRange = 0
  dtDateMonth = 1
  dtDateQuarter = 2
  dtDateHalfYear = 3
End Enum

Private objDialogBox As DialogSheet

Function GetUserName(Optional ByVal SetUserDomain = False) As String
  Attribute GetUserName.VB_Description = "r311 ¦ Получить имя текущей учётной записи"
  GetUserName = IIf(SetUserDomain, Environ("UserDomain") & "\", "") _
    & Environ("UserName")
End Function

Private Sub Auto_Open() ' book.onLoad - Подсчёт CRC_HOST = SUM( 2 ^ (item - 1) )
  Attribute Auto_Open.VB_Description = "r313 ¦ Автозапуск"
  Dim max As Integer, modulo As Integer, item As Variant, Paths() As Variant
  Const HOST As String = "#Finansist\YCHET\"
  
  With ActiveWindow
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
    .DisplayWorkbookTabs = True
  End With
  ' Очищаем коллекции с именами папок и с именами файлов
  Set DirName = Nothing: Set FileName = Nothing: Set Manager = Nothing
  ' Массив каталогов ВСЕХ существующих Банков (+ косая черта в конце строки)
  Paths = Array(HOST, HOST & "Авторы-бренды\", HOST & "Рецензирование ИБ " _
    & "Финансист\", HOST & "ИБ Юридическая пресса\", HOST & "Азбука права\", _
    HOST & "Интернет-статьи\", HOST & "Вопросы под заказ\Базы\")
  
  ' Подсчёт CRC_HOST = SUM( 2 ^ (item - 1) )
  If CRC_HOST > 2 ^ UBound(Paths) Then _
    MsgBox "CRC_HOST Error", vbCritical: Exit Sub
  max = CRC_HOST Mod 2 ^ UBound(Paths) ' КОНТРОЛЬНОЕ_ЧИСЛО % 2^[N+1]
  
  For item = UBound(Paths) To 1 Step -1
    modulo = CRC_HOST Mod 2 ^ (item - 1) ' КОНТРОЛЬНОЕ_ЧИСЛО % 2^[N]
    ' Если ТЕКУЩИЙ остаток <= ПРЕДЫДУЩИЙ остаток, То исключить каталог
    If max <= modulo Then Paths(item) = Empty
    max = modulo
  Next item
  For Each item In Paths ' Назначаем коллекцию листов для каждого Банка
    If Len(item) > 0 Then max = max + 1: GetWorkbooks item: _
      If Not Right(item, 1) = Chr(&H5C) Then max = max - 1
  Next item
  
  Paths = Array(String(2, vbCr), Replace(Join(Paths), "\ ", "\" & vbLf), _
    "Сформировать отчёт по файлам в каталогах? ", _
    " (каталоги: " & max & ", файлы: " & FileName.count & ")") ' ИНФО сообщение
  For modulo = 1 To FileName.count
    For Each item In Workbooks ' Проверка: закрыть все книги
      If item.Name = FileName(modulo) Then MsgBox "Необходимо закрыть файл """ _
        & FileName(modulo) & """", vbCritical: item = vbNo: Exit Sub
    Next item: Paths(1) = Paths(1) & FileName(modulo) & vbCr
  Next modulo: If max = 0 Then Paths = Array(True, Date, Empty, Paths(4))
  If FileName.Count > 0 Then GetForm_DialogElements dtDateRange, Paths
  '-> NEXT
  
  If Not Paths(LBound(Paths)) Then ActiveWorkbook.Saved = True Else _
    If Not IsEmpty(Paths(2)) Then Main_Sub Paths(3), Paths(2)
End Sub

Static Sub DialogButtons_Click()
  Attribute DialogButtons_Click.VB_Description = "r310 ¦ События кнопок диалогового окна"
  Dim item As Variant, str As String: str = Empty
  
  If objDialogBox Is Nothing Then Exit Sub ' HotFix!
  With objDialogBox
    Select Case .Buttons(Application.Caller).Index
      Case 1
        For Each item In FileName: str = str & vbCr & item: Next item: MsgBox _
          "Список файлов для формирования отчёта: " & vbCr & str, vbInformation
      Case 3: .Visible = xlSheetVisible ' IsChanged = -1
    End Select
  End With
End Sub

Private Sub GetForm_DialogElements(ByVal DType As DialogType, _
  ByRef Lbls As Variant)
  Attribute GetForm_DialogElements.VB_Description = "r313 ¦ Создание диалогового окна"
  Const PIXEL As Single = 5.25 ' Lbls: 1= Files, 2= Dirs, 3= Text, 4= Title
  
  Application.DisplayAlerts = False
  
    While DialogSheets.count > 0 ' Удаление всех временных форм
      DialogSheets(1).Delete
    Wend: Set objDialogBox = DialogSheets.Add
    With objDialogBox
      With .DialogFrame.ShapeRange ' Диалоговое окно
        .Width = PIXEL * 50: .Height = PIXEL * 35
        .Parent.Caption = Application.Name & " - r" & REV & Lbls(UBound(Lbls))
      End With ': .Buttons(1).Delete
      With .Buttons(1)
        If FileName.Count = 0 Then .Enabled = False
        .Left = PIXEL * 50: .Top = PIXEL * 20: .Text = "Список"
        .Width = PIXEL * 10: .Height = PIXEL * 3
        .OnAction = "DialogButtons_Click": .DismissButton = False ' Отклонить
      End With: With .Buttons(2)
        .Left = PIXEL * 50: .Top = PIXEL * 12: .Text = "Нет"
        .Width = PIXEL * 10: .Height = PIXEL * 3
      End With

      ' Граница объектов: Left[P=>13], Top[P=>6], Width[P=<50], Heigth[P=<35]
      With .Labels.Add(PIXEL * 15, PIXEL * 8, PIXEL * 20, PIXEL * 3)
        .text = "Введите начало периода: "
      End With
      With .EditBoxes.Add(PIXEL * 38, PIXEL * 8, PIXEL * 10, PIXEL * 3)
        .Name = "DateBegin"
        If Month(Date) > 3 And Month(Date) < 10 Then
          .text = Replace(DateSerial(Year(Date) - 1, 10, 1), "/", ".") ' Октябрь
        Else
          .text = Replace(DateSerial(IIf(Month(Date) < 4, Year(Date) - 1, _
            Year(Date)), 4, 1), "/", ".") ' Апрель
        End If
      End With
      With .Labels.Add(PIXEL * 15, PIXEL * 12, PIXEL * 20, PIXEL * 3)
        .text = "Введите конец периода: "
      End With
      With .EditBoxes.Add(PIXEL * 38, PIXEL * 12, PIXEL * 10, PIXEL * 3)
        .Name = "DateEnd"
        If Month(Date) > 3 And Month(Date) < 10 Then
          .text = Replace(DateSerial(Year(Date), 3 + 1, 0), "/", ".") ' Март
        Else
          .text = Replace(DateSerial(IIf(Month(Date) < 4, Year(Date) - 1, _
            Year(Date)), 9 + 1, 0), "/", ".") ' Сентябрь
        End If
      End With
      With .Labels.Add(PIXEL * 15, PIXEL * 16, PIXEL * 35, PIXEL * 21)
        Lbls(2) = ClearSpacesInText(Lbls(2))
        .text = Lbls(3) & String(2, vbLf) & Replace(Lbls(2), "#Finansist\", "> ")
      End With
      With .Buttons.Add(PIXEL * 50, PIXEL * 8, PIXEL * 10, PIXEL * 3) ' Btn "Да"
        If FileName.Count = 0 Then .Enabled = False
        .DismissButton = True ' Отклонить = .Hide
        .text = "Да": .OnAction = "DialogButtons_Click"
      End With
      .Name = "DialogBox": .Visible = xlSheetHidden ' isChanged = 0
    End With

    With objDialogBox
      .Show: Lbls(LBound(Lbls)) = CBool(.Visible)
      If CBool(.Visible) Then ' Если кнопка "Да"
        Lbls(2) = Empty: Lbls(3) = Empty
        On Error Resume Next
          Lbls(2) = CDate(.EditBoxes("DateBegin").text)
          Lbls(3) = CDate(.EditBoxes("DateEnd").text)
        On Error GoTo 0
        .Delete: Set objDialogBox = Nothing
      End If
    End With
  
  Application.DisplayAlerts = True
End Sub

Private Sub GetWorkbooks(ByVal PathName As String) ' Все статистики
  Attribute GetWorkbooks.VB_Description = "r311 ¦ Запись найденных баз/статистик в коллекцию"
  Dim strName As String: strName = GetMainPath & PathName
  
  ' Возвращаем в strName первый найденный файл по маске *.xl*
  On Error GoTo ErrDir
    If (GetAttr(strName) And vbDirectory) = vbDirectory Then
      With ThisWorkbook
        If DirName.Count = 0 And Right(.Name, 2) = "sm" Then _
          WriteLog Left(strName, InStr(23, strName, "\")) & "Архив\", _
          IIf(.ReadOnly, "Чтение", "Запись")
      End With
      ' Возвращаем в strName первый найденный файл по маске *.xl*
      strName = Dir(strName & "*.xl*", vbNormal)
      Do While strName <> vbNullString ' Выполнять ПОКА
        ' Применяем дополнительную маску для выборки файлов
        If Not strName Like "*.lnk" And Not LCase(strName) Like "*копия*" _
        And Not LCase(strName) Like "*отдел*" And (IIf(PathName Like "*\Баз*", _
          strName Like "База_*", strName Like "[Сс]татистика_*")) Then
          DirName.Add GetMainPath & PathName: FileName.Add strName
        End If: strName = Dir
      Loop
    End If
  Exit Sub
  
  ErrDir:
    If Not strName Like "*\*.xl*" And Err.Number = 53 Then Err.Number = 75
    Select Case Err.Number
      Case Is = 53: strName = "Файл не найден: "
      Case Is = 75: strName = "Нет доступа к файлу: "
      Case Is = 457: Exit Sub
      Case Else: strName = "Проверьте сетевой путь. Нет доступа к каталогу: "
    End Select: MsgBox strName & vbCr & GetMainPath & PathName _
      & IIf(Err.Number = 53, "#FILE", ""), vbCritical
    If Err.Number = 5 Or Err.Number = 53 Or Err.Number >= 75 Then End
End Sub

Function Taxpayer_Number_CRC(ByVal ITN12orTIN10 As Double) As Boolean
  Attribute INNCRC.VB_Description = "r313 ¦ Проверка контрольной суммы ИНН"
  Attribute INNCRC.VB_ProcData.VB_Invoke_Func = " \n9"
  Dim CodeLen(11) As Byte, eZ As Byte, mZ As Integer, nZ As Integer
  
  CodeLen(1) = 3
  CodeLen(2) = 7
  CodeLen(3) = 2
  CodeLen(4) = 4
  CodeLen(5) = 10
  CodeLen(6) = 3
  CodeLen(7) = 5
  CodeLen(8) = 9
  CodeLen(9) = 4
  CodeLen(10) = 6
  CodeLen(11) = 8
  
  eZ = Len(ITN12orTIN10) ' По длине определяем: Физ или Юр лицо
  If eZ - 1 > 12 Then Exit Function ' False, если больше 12 цифр
  
  For mZ = eZ - 1 To 1 Step -1
    ' Добавлен CByte()
    nZ = nZ + CByte(Mid(ITN12orTIN10, mZ, 1)) * CodeLen(12 - eZ + mZ)
  Next mZ: mZ = (nZ \ 11) * 11
  If Right(ITN12orTIN10, 1) = Right(nZ - mZ, 1) Then INNCRC = True
  If eZ = 12 Then If Not INNCRC(Left(ITN12orTIN10), eZ - 1 ) Then INNCRC = False
End Function

Function ChoiceCategory(ByVal CurrentRow As Integer) As Byte
  Attribute ChoiceCategory.VB_Description = "r313 ¦ Матрица"
  Dim Category(16) As String, eZ As Byte
  
  Category(1) = "МИНФИН" ' 1
  Category(2) = "ФНС" ' 2
  Category(3) = "СЧ[ЁЕ]ТНАЯ ПАЛАТА*"
  Category(4) = "МИНИСТЕРСТВО ТРУДА*"
  Category(5) = "РОСТРУД"
  Category(6) = "*ИНСПЕКЦИЯ ТРУДА*"
  Category(7) = "*ФТС*"
  Category(8) = "*ТАМОЖНЯ*"
  Category(9) = "ВЕД*" ' 3
  Category(10) = "НЕК*" ' 4
  Category(11) = "КОМ*" ' 5
  Category(12) = "КАЦБУН"
  Category(15) = "РИЦ" ' 6 не просматривается ' Category = 13, 14, 15
  Category(16) = "*КЦ" ' 7
  
  With Worksheets(xSUPP("sheet"))
    ' Если «Тип организации» = Ведомство (с подписью) и Ф/Л
    If UCase(.Cells(CurrentRow, xSUPP("Org_type"))) Like Category(9) _
    And Not UCase(.Cells(CurrentRow, xSUPP("Org_type"))) Like "*БЕЗ ПОДП*" Then
      ' Если «Организация в системе» МИНФИН, ФНС, ...
      For eZ = LBound(Category) To LBound(Category) + 7
        If UCase(.Cells(CurrentRow, xSUPP("NameL"))) Like Category(eZ) Then _
          Exit For
      Next eZ: ChoiceCategory = eZ
    ElseIf .Cells(CurrentRow, xSUPP("Org_base")) > 0 And .Cells(CurrentRow, _
      xSUPP("Org_base")) < 999 Then ' Если «Источник» = [Номер РИЦ]
      If UCase(.Cells(CurrentRow, xSUPP("Org_town"))) Like "М*ВА" Then
        ChoiceCategory = 13 ' k = 6
      ElseIf UCase(.Cells(CurrentRow, xSUPP("Org_town"))) Like "С*РГ" Then
        ChoiceCategory = 14
      Else
        ChoiceCategory = 15
      End If
    ' Если «Организация в системе» КАЦБУН
    ElseIf UCase(.Cells(CurrentRow, xSUPP("NameL"))) Like Category(12) Then
      ChoiceCategory = 12
    Else ' Если совпадения не найдутся, То посчитать в «Коммерч.»
      ChoiceCategory = 11 ' k = 5
      For eZ = LBound(Category) To UBound(Category)
        If UCase(.Cells(CurrentRow, xSUPP("Org_type"))) Like Category(eZ) And _
        Not UCase(.Cells(CurrentRow, xSUPP("Org_type"))) Like "*БЕЗ ПОДП*" Then
          ChoiceCategory = eZ: Exit For ' k = 4, 5, 7
        End If
      Next eZ
    End If
  End With
End Function

Private Function GetMainPath(Optional ByVal DiskOnly As Boolean) As String
  Attribute GetMainPath.VB_Description = "r311 ¦ Определить директорию/диск для поиска статистики"
  Const DecCharCode_from32 As String = "5727613487858083644661"
  Dim strSym(Len(DecCharCode_from32) \ 2) As String, eZ As Byte
  
  For eZ = 1 To UBound(strSym)
    strSym(eZ) = Chr(CByte(Mid(DecCharCode_from32, eZ * 2 - 1, 2)) + &H1F)
    If DiskOnly And strSym(eZ) = Chr(&H5C) Then Exit For
  Next eZ: GetMainPath = Join(strSym, "")
End Function

Private Sub WriteLog(ByVal LogDir As String, ByVal AccessMode As String)
  Attribute WriteLog.VB_Description = "r310 ¦ Запись в журнал об открытии отчёта"
  Dim strName As String
  
  With CreateObject("Scripting.FileSystemObject")
    If .FolderExists(LogDir) Then
      LogDir = LogDir & "Журнал_доступа.csv"
      If .FileExists(LogDir) Then .GetFile(LogDir).Attributes = 0 _
      Else Open LogDir For Append As #1: Print #1, _
        "Дата;Время;Логин;Версия;Файл;Путь;Доступ": Close #1
      Open LogDir For Append As #1
      With ThisWorkbook: strName = Left(.Name, InStrRev(.Name, ".") - 1)
        Print #1, Date & ";" & Time & ";" & GetUserName & ";" & Chr(&H72) _
          & REV & ";" & strName & ";" & .Path & ";" & AccessMode: Close #1
      End With: .GetFile(LogDir).Attributes = 1
    End If
  End With
End Sub

Function FileUnlocked(ByRef FileName As String) As Boolean
  Attribute FileUnlocked.VB_Description = "r270 ¦ Проверить занятость файла"
  On Error Resume Next
    Open FileName For Binary Access Read Write Lock Read Write As #1
    Close #1
    If Err.Number <> 0 Then FileUnlocked = True: Err.Clear
End Function
