Attribute VB_Name = "Sheet"
Option Explicit
Option Base 1
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890

Public Sub GetBanks(ByRef BankID As Collection, Optional ByRef BankSUPP _
  As Collection)
  Attribute GetBanks.VB_Description = "r310 Определение структуры файла статистики"
  If Not oID.Count = 0 Then Exit Sub ' HoxFix!
  Dim iND As Object, bank As String, sub_bank As String
  ' ВАЖНАЯ ЧАСТЬ! Заголовки найденных именованных диапазонов в коллекции
  Dim Bank_Key As New Collection
  Dim Bank_SheetID As New Collection
  Dim Bank_HeaderRow As New Collection
  Dim Номер_вопроса As New Collection
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
  ' Включаемые Банки: СВ, ПВ, П+А, БО, КФ, ЮП, VIP, SEO, ЛК, РЦ
  BankID.Add Bank_SheetID, "sheet"
  ' Номер строки, в которой находятся заголовки таблицы банка «HEAD»
  BankID.Add Bank_HeaderRow, "head"
  ' Номер колонки «№ вопроса»
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
        'If GetSheetID("STAT_", False) > 0 Then ' HotFix!
        '  If bank = "OT" Then sub_bank = "OE"
        'End If
        ' ... смотрим, является ли Банк новым И Ссылка не битая
        If Bank_Key(Bank_Key.Count) <> "_" & bank _
        And Not .Value Like "*[#]*" Then
  '        Debug.Print ActiveWorkbook.Name & ": " & "STAT_" & bank
          Bank_HeaderRow.Add .RefersToRange.Row, "STAT_" & bank
          ' ЛУЧШЕ проверять по Worksheets.CodeName (read-only),
          ' если структура Книги создаётся с нуля
          Bank_SheetID.Add GetSheetID(.Value, False), "STAT_" & bank
          ' Вписываем новый Банк и имя листа, на котором он находится
          Bank_Key.Add "_" & bank, "STAT_" & bank
          ' Активируем лист, на котором расположен Банк
          If Worksheets(Bank_SheetID("STAT_" & bank)).Visible < 0 Then _
            Worksheets(Bank_SheetID("STAT_" & bank)).Select
        End If
      End If
      If GetSheetID("STAT_", False) > 0 Then ' HotFix!
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
        If BankID(sub_bank).Item("STAT_OT") And Err.Number = 5 Then ' HotFix!
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
        If Not BankSUPP Is Nothing And sub_bank = "NameS" Then
          BankSUPP.Add GetSheetID(.Value, False), "sheet"
          BankSUPP.Add .RefersToRange.Row, "head"
          With Worksheets(BankSUPP("sheet")) ' Данные контрагентов
            bank = .Cells.SpecialCells(xlLastCell).Row
            While IsEmpty(.Cells(bank, BankSUPP("NameS")))
              bank = bank - 1
            Wend: BankSUPP.Add .Range(.Cells(BankSUPP("head") + 1, 1), .Cells( _
              bank, .Cells.SpecialCells(xlLastCell).column)).Value2, "Data"
          End With
        End If
      End If
      ' Если имя диапазона содержит Date*, То применить формат HotFix!
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
          "диапазон """ & .Name & """, либо лист защищён от записи. " & vbCr _
          & "Для просмотра связей используйте комбинацию кнопок Ctrl+F3. ", _
          vbCritical: End
        End
      End If
    End With
  Next iND: Set iND = Nothing
End Sub

Function GetSupplerRec(ByVal NameSupp As String, ByVal CheckDate As Variant, _
  Optional ByVal isSetBounds As Boolean = True) As Integer
  Attribute GetSupplerRec.VB_Description = "r310 Поиск записи контрагента"
  Dim RelevantDate As Double, aU As Variant
  Dim eZ As Byte, nZ As Integer
  
  If IsNumeric(CheckDate) And Not IsEmpty(NameSupp) Then
    If CheckDate > 0 Then
      NameSupp = Trim(NameSupp)
      aU = xSUPP("Data"): eZ = xSUPP("NameS")
      For nZ = LBound(aU, 2) To UBound(aU, 2)
        If aU(eZ, nZ) = NameSupp Then
          'Debug.Print "Рецензент " & NameSupp & " найден в стороке " & nZ
          'Debug.Print CDate(CheckDate) & " >= " & CDate(aU(xSUPP("DateD"), nZ)) & " >= " & CDate(rZ)
          
          If Not isSetBounds And GetSupplerRec = 0 Then RelevantDate = -aU(xSUPP("DateD"), nZ)
          
          ' Если требуется найти Имя контрагента И Дату актуальности, То
          If aU(xSUPP("DateD"), nZ) >= RelevantDate Then
            If CheckDate >= aU(xSUPP("DateD"), nZ) Then
              RelevantDate = aU(xSUPP("DateD"), nZ): GetSupplerRec = nZ
            
            ' Если требуется найти хотя бы Имя контрагента, То
            ElseIf Not isSetBounds And -aU(xSUPP("DateD"), nZ) >= RelevantDate Then
              RelevantDate = -aU(xSUPP("DateD"), nZ): GetSupplerRec = nZ
            End If
          End If
        End If
      Next nZ
    Else
      Debug.Print CheckDate & " isn't a Date"
    End If
  End If
End Function

Function GetRecord(ByVal Row As Integer, ByVal BankID As String, _
  ByVal SheetID As Byte, Optional ByVal Prefix_ShName As String = "STAT") _
  As Variant ' Если "Row < 1" искать заголовок коллекции xID
  Attribute GetRecord.VB_Description = "r310 Значение ячейки из коллекции xID"
  Dim CodeName As String: CodeName = Prefix_ShName & xID("key").Item(SheetID)
  
  ' Если "Row > 0", то взять ЗНАЧЕНИЕ ячейки по коллекции xID
  GetRecord = xID(BankID).item(CodeName): If Row > 0 Then GetRecord = _
    Worksheets(xID("sheet").item(CodeName)).Cells(Row, GetRecord).Value2
End Function

Function GetSheetID(ByVal SheetCodeName As String, Optional ByRef ThisBook _
  As Boolean = True) As Byte ' ThisBook - ДА, эта книга
  Attribute GetSheetID.VB_Description = "r310 Найти индекс листа по CodeName"
  Dim GetBook As Workbook, GetSheet As Worksheet
  
  Set GetBook = IIf(ThisBook, ThisWorkbook, ActiveWorkbook)
  If InStr(SheetCodeName, "!") > 0 Then _
    SheetCodeName = Replace(Mid(SheetCodeName, 2, InStr( _
      SheetCodeName, "!") - 2), "'", "") ' Имя листа должно быть БЕЗ "!"
  For Each GetSheet In GetBook.Worksheets
    If InStr(1, GetSheet.CodeName, SheetCodeName, vbTextCompare) _
    Or InStr(1, GetSheet.Name, SheetCodeName, vbTextCompare) Then _
      GetSheetID = GetSheet.Index: Exit For
  Next GetSheet: Set GetSheet = Nothing: Set GetBook = Nothing
End Function

Public Sub DeleteModulesAndCode(ByRef WB As Object) ' Удалить модули
  Attribute DeleteModulesAndCode.VB_Description = "r270 Удалить модули из книги"
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
