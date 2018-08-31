Attribute VB_Name = "Frame"
Option Explicit
Option Base 1
Option Private Module
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890

Public Property Get GetUserName() As String
  Attribute SettingsBankID.VB_Description = "r300 Получить имя текущей учётной записи"
  GetUserName = Environ("UserName")
End Property

Public Sub SettingsBankID(ByRef BankID As Collection, Optional ByRef BankSUPP _
  As Collection)
  Attribute SettingsBankID.VB_Description = "r300 Определить структуру файла статистики"
  Dim iND As Object, bank As String, sub_bank As String
  ' ВАЖНАЯ ЧАСТЬ! Занесение в коллекцию выбранные именованные диапазоны,
  '               необходимые для создания отчёта
  Dim Bank_Key As New Collection
  Dim Bank_SheetID As New Collection
  Dim Bank_HeaderRow As New Collection
  Dim Номер_вопроса As New Collection ' ver.2 rev.300
  Dim Имя_Контрагента As New Collection
  Dim Дата_Поступления As New Collection
  Dim Дата_передачи_аутсорсерам As New Collection
  Dim Дата_Акта As New Collection
  Dim Номер_Акта As New Collection
  Dim Дата_Договора As New Collection
  Dim Номер_Договора As New Collection
  Dim Дата_Перечислений As New Collection
  Dim Сумма_Итого As New Collection
  
  BankID.Add Bank_Key, "key"
  ' Коллекция с индексами листов для каждого Банка
  ' Включаемые Банки: БО, КФ, П, ПВ, СВ, ЛК, ЮП
  BankID.Add Bank_SheetID, "sheet"
  ' Номер строки, в которой находятся заголовки таблицы банка «HEAD»
  BankID.Add Bank_HeaderRow, "head"
  ' Номер колонки «№ вопроса» ver.2 rev.300
  BankID.Add Номер_вопроса, "QNum"
  ' Номер колонки «Поставщик (кратко)»
  BankID.Add Имя_Контрагента, "NameS"
  ' Номер колонки, к которой находится «Дата поступления»
  BankID.Add Дата_Поступления, "Date_mail"
  ' Номер колонки «Дата передачи аутсорсерам»
  BankID.Add Дата_передачи_аутсорсерам, "Date_OSend"
  ' Номер колонки, к которой находится «Дата акта»
  BankID.Add Дата_Акта, "Date_akt"
  ' Номер колонки, к которой находится «Номер акта»
  BankID.Add Номер_Акта, "Num_akt"
  ' Номер колонки, к которой находится «Дата договора»
  BankID.Add Дата_Договора, "Date_dog"
  ' Номер колонки, к которой находится «Номер договора»
  BankID.Add Номер_Договора, "Num_dog"
  ' Номер колонки, к которой находится «Дата перечислений»
  BankID.Add Дата_Перечислений, "Date_APay"
  ' Номер колонки, к которой находится «Итого...» Сумма по документу
  BankID.Add Сумма_Итого, "Sum_All"
  
  For Each iND In ActiveWorkbook.Names
    With iND
      On Error Resume Next
      bank = Empty: bank = Left(.Name, InStr(.Name, "_") - 1)
      sub_bank = Right(.Name, Len(.Name) - Len(bank) - 1)
      ' Если появляется Банк ...
      If Len(bank) = 2 Then
        'If GetSheetID("STAT_") > 0 Then ' «Костыль»
        '  If bank = "OT" Then sub_bank = "OE"
        'End If
        ' ... смотрим, является ли Банк новым И Ссылка не битая
        If Bank_Key(Bank_Key.Count) <> "_" & bank _
        And Not .Value Like "*[#]*" Then
  '        Debug.Print ActiveWorkbook.Name & ": " & "STAT_" & bank
          Bank_HeaderRow.Add .RefersToRange.Row, "STAT_" & bank
          ' ЛУЧШЕ проверять по Worksheets.CodeName (read-only),
          ' если структура Книги создаётся с нуля
          Bank_SheetID.Add GetSheetID(.Value), "STAT_" & bank
          ' Вписываем новый Банк и имя листа, на котором он находится
          Bank_Key.Add "_" & bank, "STAT_" & bank
          ' Активируем лист, на котором расположен Банк
          If Worksheets(Bank_SheetID("STAT_" & bank)).Visible < 0 Then _
            Worksheets(Bank_SheetID("STAT_" & bank)).Select
        End If
      End If
      If GetSheetID("STAT_") > 0 Then ' «Костыль»
        Select Case sub_bank
          Case "Quant_inbox": sub_bank = "AMT_source"
          Case "Quant_new": sub_bank = "AimAMT"
          Case "Quant_In": sub_bank = "AimAMT_gb"
          Case "Goszak_In": sub_bank = "AimAMT_gz"
          Case "Quant_pay": sub_bank = "AcceptAMT"
          Case "Quant_Out": sub_bank = "AcceptAMT_gb"
          Case "Goszak_Out": sub_bank = "AcceptAMT_gz"
        End Select
      End If
      ' Если диапазон из одного значения
      If .RefersToRange.Count = 1 And Len(bank) = 2 Then
        BankID(sub_bank).Add .RefersToRange.column, "STAT_" & bank
      ElseIf bank = "PART" Then
        If BankID(sub_bank).Item("STAT_OT") And Err.Number = 5 Then ' «Костыль»
          BankID(sub_bank).Add .RefersToRange.column, "STAT_BO"
          BankID(sub_bank).Add .RefersToRange.column, "STAT_KF"
        Else ' after := "STAT_OT"
          BankID(sub_bank).Add .RefersToRange.column, "STAT_BO", "STAT_OT"
          BankID(sub_bank).Add .RefersToRange.column, "STAT_KF", "STAT_OT"
        End If
      ElseIf bank = "ARCH" Or bank = "SUPP" Then
  '      Debug.Print ActiveWorkbook.Name & ": " & sub_bank
        ' Заносятся реквизиты контрагентов...
        BankSUPP.Add .RefersToRange.column, sub_bank
        ' ... и Номер строки заголовка реквизитов «HEAD»
        If sub_bank = "NameS" Then
          BankSUPP.Add GetSheetID(.Value), "sheet"
          BankSUPP.Add .RefersToRange.Row, "head"
        End If
      End If
      ' Если имя диапазона содержит Date*, То применить формат «Костыль»
      If sub_bank Like "Date*" And .RefersToRange.Count = 1 Then
        With Worksheets(Bank_SheetID("STAT_" & bank)).Cells(1, _
        .RefersToRange.column).EntireColumn.Interior
          .NumberFormat = "m/d/yyyy"
          '.ColorIndex = 44 ' Отключить при открытии файла для редактирования
        End With
      End If
      
      If Err.Number = 1004 And Not IsEmpty(bank) Then
        On Error GoTo 0
        MsgBox "В книге """ & ActiveWorkbook.Name & """ поломан именованный " & _
          "диапазон """ & .Name & """, либо лист защищён от записи. " & vbCrLf _
          & "Для просмотра связей используйте комбинацию кнопок Ctrl+F3. ", _
          vbCritical: End
        End
      End If
    End With
  Next iND: Set iND = Nothing
End Sub


Function ChoiceCategory(ByVal row As Integer) As Byte
  Attribute SettingsBankID.VB_Description = "r270 Матрица"
  Dim Category(16) As String, a As Byte
  
  Category(1) = "МИНФИН" ' 1
  Category(2) = "ФНС" ' 2
  Category(3) = "СЧ[Е,Ё]ТНАЯ ПАЛАТА*"
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
  
  With Worksheets(BankSUPP("sheet"))
    ' Если «Тип организации» = Ведомство (с подписью) и Ф/Л
    If UCase(.Cells(row, BankSUPP("Org_type"))) Like Category(9) _
    And Not UCase(.Cells(row, BankSUPP("Org_type"))) Like "*БЕЗ ПОДП*" Then
      ' Если «Организация в системе» МИНФИН, ФНС, ...
      For a = LBound(Category) To LBound(Category) + 7
        If UCase(.Cells(row, BankSUPP("NameL"))) Like Category(a) Then Exit For
      Next a: ChoiceCategory = a
    ElseIf .Cells(row, BankSUPP("Org_base")) > 0 And .Cells(row, _
      BankSUPP("Org_base")) < 999 Then ' Если «Источник» = [Номер РИЦ]
      If UCase(.Cells(row, BankSUPP("Org_town"))) Like "М*ВА" Then
        ChoiceCategory = 13 ' k = 6
      ElseIf UCase(.Cells(row, BankSUPP("Org_town"))) Like "С*РГ" Then
        ChoiceCategory = 14
      Else
        ChoiceCategory = 15
      End If
    ' Если «Организация в системе» КАЦБУН
    ElseIf UCase(.Cells(row, BankSUPP("NameL"))) Like Category(12) Then
      ChoiceCategory = 12
    Else ' Если совпадения не найдутся, То посчитать в «Коммерч.»
      ChoiceCategory = 11 ' k = 5
      For a = LBound(Category) To UBound(Category)
        If UCase(.Cells(row, BankSUPP("Org_type"))) Like Category(a) _
        And Not UCase(.Cells(row, BankSUPP("Org_type"))) Like "*БЕЗ ПОДП*" Then
          ChoiceCategory = a: Exit For ' k = 4, 5, 7
        End If
      Next a
    End If
  End With
End Function

Property Get GetMainPath(Optional ByVal DiskOnly As Boolean) As String
  Attribute SettingsBankID.VB_Description = "r302 Определить директорию/диск для поиска статистики"
  Const DecCharCode_from32 As String = "5727613487858083644661"
  Dim s As Byte, strSym(Len(DecCharCode_from32) \ 2) As String
  
  For s = 1 To UBound(strSym)
    strSym(s) = Chr(CByte(Mid(DecCharCode_from32, s * 2 - 1, 2)) + &H1F)
    If DiskOnly And strSym(s) = Chr(&H5C) Then Exit For
  Next s: GetMainPath = Join(strSym, "")
End Property

Public Sub RecLog(ByVal LogDir As String, ByVal AccessMode As String)
  Attribute SettingsBankID.VB_Description = "r302 Запись в журнал об открытии отчёта"
  Dim Name As String
  
  With CreateObject("Scripting.FileSystemObject")
    If .FolderExists(LogDir) Then
      LogDir = LogDir & "Журнал_доступа.csv"
      If .FileExists(LogDir) Then .GetFile(LogDir).Attributes = 0 _
      Else Open LogDir For Append As #1: Print #1, _
        "Дата;Время;Логин;Версия;Файл;Путь;Доступ": Close #1
      Open LogDir For Append As #1
      With ThisWorkbook: Name = Left(.Name, InStrRev(.Name, ".") - 1)
        Print #1, Date & ";" & Time & ";" & GetUserName & ";" & Chr(&H72) _
          & REV & ";" & Name & ";" & .Path & ";" & AccessMode: Close #1
      End With: .GetFile(LogDir).Attributes = 1
    End If
  End With
End Sub

Function FileUnlocked(ByRef strFileName As String) As Boolean
  Attribute SettingsBankID.VB_Description = "r270 Проверить занятость файла"
  On Error Resume Next
  Open strFileName For Binary Access Read Write Lock Read Write As #1
  Close #1
  If Err.Number <> 0 Then FileUnlocked = True: Err.Clear
End Function

Public Sub DeleteModulesAndCode(ByRef WB As Object) ' Удалить модули
  Attribute SettingsBankID.VB_Description = "r270 Удалить модули из книги"
  ' Центр управления безопасностью -> Параметры макросов -> Доверять доступ к VBA
  Dim iVBComponents As Object, iVBComponent As Object
  
  Set iVBComponents = WB.VBProject.VBComponents
  For Each iVBComponent In iVBComponents
    Select Case iVBComponent.Type
      Case 1 To 3
        iVBComponents.Remove iVBComponent
      Case 100
        With iVBComponent.CodeModule
          .DeleteLines 1, .CountOfLines
        End With
    End Select
  Next iVBComponent
End Sub
