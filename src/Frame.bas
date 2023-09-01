Attribute VB_Name = "Frame"
Option Explicit
Option Base 1
Option Private Module
'123456789012345678901234567890123456h8nor@уа56789012345678901234567890123456789

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
  dtDateSemester = 3 ' r314
End Enum

Private objDialogBox As DialogSheet

Public Function GetUserName(Optional ByVal setUserDomain = False) As String
  Attribute GetUserName.VB_Description = "r314 ¦ Получить имя текущей учётной записи"
  GetUserName = IIf(setUserDomain, Environ("UserDomain") & "\", "") _
    & Environ("UserName")
End Function

Private Sub Auto_Open() ' book.onLoad - Подсчёт CRC_HOST = SUM( 2 ^ (item - 1) )
  Attribute Auto_Open.VB_Description = "r317 ¦ Автозапуск"
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
  Paths = Array(0, Replace(HOST, "YCHET", "#KF_KBO") & "POSTE\", _
    HOST & "Вопросы под заказ\", HOST & "Вопросы под заказ\Базы\", _
    HOST & "Рецензирование ИБ Финансист\", HOST & "ИБ Юридическая пресса\", _
    HOST & "Азбука \", HOST & UCase("Перезакупка\"), HOST) ' r315
  
  ' Подсчёт CRC_HOST = SUM( 2 ^ (item - 1) )
  If CRC_HOST > 2 ^ UBound(Paths) Then _
    HookMsg "CRC_HOST Error", vbCritical: Exit Sub
  max = CRC_HOST Mod 2 ^ UBound(Paths) ' КОНТРОЛЬНОЕ_ЧИСЛО % 2^[N+1]
  
  For item = UBound(Paths) To 2 Step -1
    modulo = CRC_HOST Mod 2 ^ (item - 1) ' КОНТРОЛЬНОЕ_ЧИСЛО % 2^[N]
    ' Если ТЕКУЩИЙ остаток <= ПРЕДЫДУЩИЙ остаток, То исключить каталог
    If max <= modulo Then Paths(item) = Empty
    max = modulo
    If Len(Paths(item)) > 0 Then
      Call GetWorkbooks(Paths(item))
      If Right(Paths(item), 1) = Chr(&H5C) Then Paths(1) = Paths(1) + 1
    End If
  Next item: max = Paths(1): Paths(1) = Empty ' HotFix!
  If CRC_HOST > 0 Then Paths = Array(String(2, vbCr), Replace(Join(Paths), _
    "\ ", "\" & vbLf), "Сформировать отчёт по файлам в каталогах? ", _
    " (каталоги: " & max & ", файлы: " & FileName.Count & ")") ' ИНФО сообщение
  For modulo = 1 To FileName.Count
    For Each item In Workbooks ' Проверка: закрыть все книги
      If item.Name = FileName(modulo) Then HookMsg "Необходимо закрыть файл """ _
        & FileName(modulo) & """", vbCritical: item = vbNo: Exit Sub
    Next item: Paths(1) = Paths(1) & FileName(modulo) & vbCr
  Next modulo: If max = 0 Then Paths = Array(True, Date, Empty, Paths(4))
  If FileName.Count > 0 Then GetForm_DialogElements dtDateRange, Paths _
    Else HookMsg "ОШИБКА! Нет книг по критерием отбора", vbRetryCancel  
  '-> NEXT
  
  If Not Paths(LBound(Paths)) Then ActiveWorkbook.Saved = True Else _
    If Not IsEmpty(Paths(2)) Then If IS_DEBUG Then Application _
      .VBE.MainWindow.Visible = False Else Main_Sub Paths(3), Paths(2) ' r317
End Sub

Static Sub DialogButtons_Click()
  Attribute DialogButtons_Click.VB_Description = "r310 ¦ События кнопок диалогового окна"
  Dim item As Variant, str As String: str = Empty
  
  If objDialogBox Is Nothing Then Exit Sub ' HotFix!
  With objDialogBox
    Select Case .Buttons(Application.Caller).Index
      Case Is = 1
        For Each item In FileName: str = str & vbCr & item: Next item: MsgBox _
          "Список файлов для формирования отчёта: " & vbCr & str, vbInformation
      Case Is = 3: .Visible = xlSheetVisible ' IsChanged = -1
    End Select
  End With
End Sub

Private Sub GetForm_DialogElements(ByVal formType As DialogType, _
  ByRef Ref_Lbls As Variant)
  Attribute GetForm_DialogElements.VB_Description = "r317 ¦ Создание диалогового окна"
  Const PIXEL As Single = 5.25 ' Lbls: 1= Files, 2= Dirs, 3= Text, 4= Title
  
  Application.DisplayAlerts = False
  
  While DialogSheets.Count > 0 ' Удаление всех временных форм
    DialogSheets(1).Delete
  Wend: Set objDialogBox = DialogSheets.Add
  With objDialogBox
    With .DialogFrame.ShapeRange ' Диалоговое окно
      .Width = PIXEL * 50: .Height = PIXEL * 35
      .Parent.Caption = Application.Name & " - r" & REV & Ref_Lbls(UBound(Ref_Lbls))
    End With ': .Buttons(1).Delete
    With .Buttons(1)
      If FileName.Count = 0 Then .Enabled = False
      .Left = PIXEL * 50: .Top = PIXEL * 20: .text = "Список"
      .Width = PIXEL * 10: .Height = PIXEL * 3
      .OnAction = "DialogButtons_Click": .DismissButton = False ' Отклонить
    End With: With .Buttons(2)
      .Left = PIXEL * 50: .Top = PIXEL * 12: .text = "Нет"
      .Width = PIXEL * 10: .Height = PIXEL * 3
    End With

    ' Граница объектов: Left[P=>13], Top[P=>6], Width[P=<50], Heigth[P=<35]
    Select Case formType

      Case Is = dtDateRange ' ДИАПАЗОН ' r314
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
        
      Case Is = dtDateMonth ' МЕСЯЦ ' r317
        With .Labels.Add(PIXEL * 15, PIXEL * 8, PIXEL * 20, PIXEL * 3)
          .text = "Выберите месяц: "
        End With
        With .DropDowns.Add(PIXEL * 36, PIXEL * 8, PIXEL * 12, PIXEL * 3)
          .Name = "MonthNum"
          .DropDownLines = 12
          .AddItem "Январь"
          .AddItem "Февраль"
          .AddItem "Март"
          .AddItem "Апрель"
          .AddItem "Май"
          .AddItem "Июнь"
          .AddItem "Июль"
          .AddItem "Август"
          .AddItem "Сентябрь"
          .AddItem "Октябрь"
          .AddItem "Ноябрь"
          .AddItem "Декабрь"
          .Value = IIf(Month(Date) = 1, 12, Month(Date) - 1)
        End With
        With .Labels.Add(PIXEL * 15, PIXEL * 12, PIXEL * 20, PIXEL * 3)
          .text = "Введите год: "
        End With
        With .EditBoxes.Add(PIXEL * 36, PIXEL * 12, PIXEL * 12, PIXEL * 3)
          .Name = "MonthYear"
          .text = IIf(Month(Date) > 1, Year(Date), Year(Date) - 1)
        End With
        
      Case Is = dtDateQuarter ' КВАРТАЛ ' r315
        With .Labels.Add(PIXEL * 15, PIXEL * 8, PIXEL * 20, PIXEL * 3)
          .text = "Выберите квартал: "
        End With
        With .DropDowns.Add(PIXEL * 36, PIXEL * 8, PIXEL * 12, PIXEL * 3)
          .Name = "QuarterNum"
          .DropDownLines = 4
          .AddItem "1 кварт."
          .AddItem "2 кварт."
          .AddItem "3 кварт."
          .AddItem "4 кварт."
          .Value = IIf((Month(Date) - 1) \ 3 < 1, 4, (Month(Date) - 1) \ 3)
        End With
        With .Labels.Add(PIXEL * 15, PIXEL * 12, PIXEL * 20, PIXEL * 3)
          .text = "Введите год: "
        End With
        With .EditBoxes.Add(PIXEL * 36, PIXEL * 12, PIXEL * 12, PIXEL * 3)
          .Name = "QuarterYear"
          .text = IIf(Month(Date) >= 3, Year(Date), Year(Date) - 1)
        End With
        
      Case Is = dtDateSemester ' ПОЛУГОДИЕ ' r315
        With .Labels.Add(PIXEL * 15, PIXEL * 8, PIXEL * 21, PIXEL * 3)
          .text = "Выберите полугодие: "
        End With
        With .DropDowns.Add(PIXEL * 36, PIXEL * 8, PIXEL * 12, PIXEL * 3)
          .Name = "SemesterNum"
          .DropDownLines = 2
          .AddItem "Окт - Мар"
          .AddItem "Апр - Сен"
          .Value = IIf((Month(Date) - 1) < 4 Or (Month(Date) - 1) > 9, 2, 1)
        End With
        With .Labels.Add(PIXEL * 15, PIXEL * 12, PIXEL * 21, PIXEL * 3)
          .text = "Введите год: "
        End With
        With .EditBoxes.Add(PIXEL * 36, PIXEL * 12, PIXEL * 12, PIXEL * 3)
          .Name = "SemesterYear"
          .text = IIf(Month(Date) >= 3, Year(Date), Year(Date) - 1)
        End With
      
    End Select
    
    With .Labels.Add(PIXEL * 15, PIXEL * 16, PIXEL * 35, PIXEL * 21)
      Ref_Lbls(2) = ClearSpacesInText(Ref_Lbls(2))
      .text = Ref_Lbls(3) & String(2, vbLf) & Replace(Ref_Lbls(2), "#Finansist\", "> ")
    End With
    With .Buttons.Add(PIXEL * 50, PIXEL * 8, PIXEL * 10, PIXEL * 3) ' Btn "Да"
      If FileName.Count = 0 Then .Enabled = False
      .DismissButton = True ' Отклонить = .Hide
      .text = "Да": .OnAction = "DialogButtons_Click"
    End With
    .Name = "DialogBox": .Visible = xlSheetHidden ' isChanged = 0
  End With

  With objDialogBox
    .Show: Ref_Lbls(LBound(Ref_Lbls)) = CBool(.Visible)
    If CBool(.Visible) Then ' Если кнопка "Да"
      Ref_Lbls(2) = Empty: Ref_Lbls(3) = Empty
      On Error Resume Next
      '< <<<
      Select Case formType ' r315
        
        Case Is = dtDateRange ' ДИАПАЗОН
          Ref_Lbls(2) = CDate(.EditBoxes("DateEnd").text)
          Ref_Lbls(3) = CDate(.EditBoxes("DateBegin").text)

        Case Is = dtDateMonth ' МЕСЯЦ ' r317
          Ref_Lbls(2) = DateSerial(.EditBoxes("MonthYear").text, _
            .DropDowns("MonthNum").Value + 1, 0)
          Ref_Lbls(3) = DateSerial(.EditBoxes("MonthYear").text, _
            .DropDowns("MonthNum").Value, 1)

        Case Is = dtDateQuarter ' КВАРТАЛ
          Ref_Lbls(2) = DateSerial(.EditBoxes("QuarterYear").text, _
            .DropDowns("QuarterNum").Value * 3 + 1, 0)
          Ref_Lbls(3) = DateSerial(IIf(Month(Ref_Lbls(2)) >= 3, Year( _
            Ref_Lbls(2)), Year(Ref_Lbls(2)) - 1), Month(Ref_Lbls(2)) - 2, 1)

        Case Is = dtDateSemester ' ПОЛУГОДИЕ
          Ref_Lbls(2) = DateSerial(.EditBoxes("SemesterYear").text, _
            IIf(.DropDowns("SemesterNum").Value > 1, 10, 4), 0)
          Ref_Lbls(3) = DateAdd("m", -6, DateAdd("d", 1, Ref_Lbls(2)))
        
      End Select
      '> >>>
      On Error GoTo 0
      .Delete: Set objDialogBox = Nothing
    End If
  End With
  
  Application.DisplayAlerts = True
End Sub

Private Sub GetWorkbooks(ByVal pathName As String) ' Все статистики
  Attribute GetWorkbooks.VB_Description = "r316 ¦ Запись найденных баз/статистик в коллекцию"
  Dim strName As String, Item As Variant: pathName = GetMainPath & pathName
  Const HOST As String = "*Finansist\YCHET\"
  
  ' Возвращаем в strName первый найденный файл по маске *.xl*
  On Error GoTo ErrDir
  '< <<<
  If (GetAttr(pathName) And vbDirectory) = vbDirectory Then
    With ThisWorkbook
      If DirName.Count = 0 And Right(.Name, 2) = "sm" Then _
        WriteLog Left(pathName, InStr(23, pathName, "\")) & "Архив\", _
          IIf(.ReadOnly, "Чтение", "Запись")
    End With
    ' Возвращаем в strName первый найденный файл по маске *.xl*
    strName = Dir(pathName & "*.xl*", vbNormal)
    Do While strName <> vbNullString ' Выполнять ПОКА
      ' Применяем дополнительную маску для выборки файлов
      For Each Item In Split(WORKBOOKS_FILTER, "$")
        If strName Like Item Then
          Select Case True
            
            Case Is = pathName Like Replace(HOST, "YCHET", "[#]KF_KBO") & "POSTE\"
              HookMsg " SKIP #KF_KBO: " & strName, vbOKCancel: Item = Empty
              
            Case Is = pathName Like HOST & "Вопросы под заказ\"
              If strName Like "Вопросы_*.xls*" Then Item = strName
              
            Case Is = pathName Like HOST & "Вопросы под заказ\Базы\"
              If strName Like "База_*.xls*" Then Item = strName
              
            Case Is = pathName Like HOST & "Рецензирование ИБ Финансист\", _
            pathName Like HOST & "ИБ Юридическая пресса\", _
            pathName Like HOST & "Азбука*", pathName Like HOST
              If strName Like "*[Сс]татистика_*.xls*" Then Item = strName
              
            Case Is = pathName Like HOST & UCase("Перезакупка\")
              If strName Like "*перезакупк[аи]_*.xls*" Then Item = strName
              
            Case Else
              HookMsg "#ELSE: " & strName, vbRetryCancel: Item = Empty
              If strName Like "*[Сс]татистика_*.xls*" Then Item = strName ' TEST
              Stop
              
          End Select
          If Not Item Like "*.lnk" And Not LCase(Item) Like "*копия*" _
          And Not LCase(Item) Like "*отдел*" And Len(Item) > 5 Then
            DirName.Add pathName: FileName.Add Item
          
            HookMsg "+DONE " & Item, vbOKCancel
          Else
            HookMsg " SKIP " & strName, vbOKCancel
          End If
          
        Else
          HookMsg "-SKIP " & strName & " LIKE " & Item, vbOKCancel
          
        End If
      Next Item: strName = Dir
    Loop
  End If
  '> >>>
  Exit Sub
  
  ErrDir:
    'If Err.Number = 76 Then pathName = ActiveWorkbook.Path & "\": Resume Next ' TEST ' r314
    If Not strName Like "*\*.xl*" And Err.Number = 53 Then Err.Number = 75
    Select Case Err.Number
      Case Is = 53: strName = "Файл не найден: "
      Case Is = 75: strName = "Нет доступа к файлу: "
      Case Is = 457: Exit Sub
      Case Else: strName = "Проверьте сетевой путь. Нет доступа к каталогу: "
    End Select: HookMsg strName & vbCr & pathName _
      & IIf(Err.Number = 53, "#FILE", ""), vbCritical
    If Err.Number = 5 Or Err.Number = 53 Or Err.Number >= 75 Then End
End Sub

Public Function Taxpayer_Number_CRC(ByVal ITN12orTIN10 As Double) As Boolean
  Attribute Taxpayer_Number_CRC.VB_Description = "r314 ¦ Проверка контрольной суммы ИНН"
  Attribute Taxpayer_Number_CRC.VB_ProcData.VB_Invoke_Func = " \n9"
  Dim CodeLen(11) As Byte, eZ As Byte, mZ As Integer, nZ As Integer
  
  CodeLen(01) = 3
  CodeLen(02) = 7
  CodeLen(03) = 2
  CodeLen(04) = 4
  CodeLen(05) = 10
  CodeLen(06) = 3
  CodeLen(07) = 5
  CodeLen(08) = 9
  CodeLen(09) = 4
  CodeLen(10) = 6
  CodeLen(11) = 8
  
  eZ = Len(ITN12orTIN10) ' По длине определяем: Физ или Юр лицо
  If eZ - 1 > 12 Then Exit Function ' False, если больше 12 цифр
  
  For mZ = eZ - 1 To 1 Step -1
    nZ = nZ + CByte(Mid(ITN12orTIN10, mZ, 1)) * CodeLen(12 - eZ + mZ)
  Next mZ: mZ = (nZ \ 11) * 11
  If Right(ITN12orTIN10, 1) = Right(nZ - mZ, 1) Then Taxpayer_Number_CRC = True
  If eZ = 12 Then If Not Taxpayer_Number_CRC(Left(ITN12orTIN10, eZ - 1)) Then _
    Taxpayer_Number_CRC = False
End Function

Public Function ChoiceCategory(ByVal currentRow As Integer) As Byte
  Attribute ChoiceCategory.VB_Description = "r314 ¦ Матрица"
  Dim Category(16) As String, eZ As Byte
  
  Category(01) = "МИНФИН" ' 1
  Category(02) = "ФНС" ' 2
  Category(03) = "СЧ[ЁЕ]ТНАЯ ПАЛАТА*"
  Category(04) = "МИНИСТЕРСТВО ТРУДА*"
  Category(05) = "РОСТРУД"
  Category(06) = "*ИНСПЕКЦИЯ ТРУДА*"
  Category(07) = "*ФТС*"
  Category(08) = "*ТАМОЖНЯ*"
  Category(09) = "ВЕД*" ' 3
  Category(10) = "НЕК*" ' 4
  Category(11) = "КОМ*" ' 5
  Category(12) = "КАЦБУН"
  Category(15) = "РИЦ" ' 6 не просматривается ' Category = 13, 14, 15
  Category(16) = "*КЦ" ' 7
  
  With Worksheets(xSUPP("sheet"))
    ' Если «Тип организации» = Ведомство (с подписью) и Ф/Л
    If UCase(.Cells(currentRow, xSUPP("Org_type"))) Like Category(9) _
    And Not UCase(.Cells(currentRow, xSUPP("Org_type"))) Like "*БЕЗ ПОДП*" Then
      ' Если «Организация в системе» МИНФИН, ФНС, ...
      For eZ = LBound(Category) To LBound(Category) + 7
        If UCase(.Cells(currentRow, xSUPP("NameL"))) Like Category(eZ) Then _
          Exit For
      Next eZ: ChoiceCategory = eZ
    ElseIf .Cells(currentRow, xSUPP("Org_base")) > 0 And .Cells(currentRow, _
      xSUPP("Org_base")) < 999 Then ' Если «Источник» = [Номер РИЦ]
      If UCase(.Cells(currentRow, xSUPP("Org_town"))) Like "М*ВА" Then
        ChoiceCategory = 13 ' k = 6
      ElseIf UCase(.Cells(currentRow, xSUPP("Org_town"))) Like "С*РГ" Then
        ChoiceCategory = 14
      Else
        ChoiceCategory = 15
      End If
    ' Если «Организация в системе» КАЦБУН
    ElseIf UCase(.Cells(currentRow, xSUPP("NameL"))) Like Category(12) Then
      ChoiceCategory = 12
    Else ' Если совпадения не найдутся, То посчитать в «Коммерч.»
      ChoiceCategory = 11 ' k = 5
      For eZ = LBound(Category) To UBound(Category)
        If UCase(.Cells(currentRow, xSUPP("Org_type"))) Like Category(eZ) And _
        Not UCase(.Cells(currentRow, xSUPP("Org_type"))) Like "*БЕЗ ПОДП*" Then
          ChoiceCategory = eZ: Exit For ' k = 4, 5, 7
        End If
      Next eZ
    End If
  End With
End Function

Private Function GetMainPath(Optional ByVal diskOnly As Boolean) As String
  Attribute GetMainPath.VB_Description = "r314 ¦ Определить директорию/диск для поиска статистики"
  Const DecCharCode_from32 As String = "5727613487858083644661"
  Dim strSym(Len(DecCharCode_from32) \ 2) As String, eZ As Byte
  
  For eZ = 1 To UBound(strSym)
    strSym(eZ) = Chr(CByte(Mid(DecCharCode_from32, eZ * 2 - 1, 2)) + &H1F)
    If diskOnly And strSym(eZ) = Chr(&H5C) Then Exit For
  Next eZ: GetMainPath = Join(strSym, "")
End Function

Private Sub WriteLog(ByVal logDir As String, ByVal accessMode As String)
  Attribute WriteLog.VB_Description = "r314 ¦ Запись в журнал об открытии отчёта"
  Dim strName As String
  
  With CreateObject("Scripting.FileSystemObject")
    If .FolderExists(logDir) Then
      logDir = logDir & "Журнал_доступа.csv"
      If .FileExists(logDir) Then .GetFile(logDir).Attributes = 0 _
      Else Open logDir For Append As #1: Print #1, _
        "Дата;Время;Логин;Версия;Файл;Путь;Доступ": Close #1
      Open logDir For Append As #1
      With ThisWorkbook: strName = Left(.Name, InStrRev(.Name, ".") - 1)
        Print #1, Date & ";" & Time & ";" & GetUserName & ";" & Chr(&H72) _
          & REV & ";" & strName & ";" & .Path & ";" & accessMode: Close #1
      End With: .GetFile(logDir).Attributes = 1
    End If
  End With
End Sub

Public Function FileUnlocked(ByRef Ref_FileName As String) As Boolean
  Attribute FileUnlocked.VB_Description = "r314 ¦ Проверить занятость файла"
  On Error Resume Next
    '< <<<
    Open Ref_FileName For Binary Access Read Write Lock Read Write As #1
    Close #1
    If Err.Number <> 0 Then FileUnlocked = True: Err.Clear
End Function
