Attribute VB_Name = "Sheet"
Option Explicit
Option Base 1
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890
Private Const PRE = "STAT_"

Public Sub GetBanks(ByRef objBankID As Collection, Optional ByRef objBankSUPP _
  As Collection)
  Attribute GetBanks.VB_Description = "r311 ¦ Определение структуры файла статистики"
  If Not objBankID.Count = 0 Then Exit Sub ' HoxFix!
  Dim objNamed As Object, bank As String, field As String
  ' ВАЖНАЯ ЧАСТЬ! Заголовки найденных именованных диапазонов в коллекции
  Dim Bank_Key As New Collection
  Dim Bank_SheetID As New Collection
  Dim Bank_HeaderRow As New Collection
  Dim Номер_вопроса As New Collection
  Dim Имя_Контрагента As New Collection
  Dim Дата_Поступления As New Collection
  Dim Дата_Передачи_аутсорсерам As New Collection
  Dim Дата_Акта As New Collection
  Dim Номер_Акта As New Collection
  Dim Дата_Договора As New Collection
  Dim Номер_Договора As New Collection
  Dim Дата_Перечислений As New Collection
  Dim Сумма_Итого As New Collection
  
  objBankID.Add Bank_Key, "key"
  ' Коллекция с индексами листов для каждого Банка
  ' Включаемые Банки: СВ, ПВ, П+А, БО, КФ, ЮП, VIP, SEO, ЛК, РЦ
  objBankID.Add Bank_SheetID, "sheet"
  ' Номер строки, в которой находятся заголовки таблицы банка «HEAD»
  objBankID.Add Bank_HeaderRow, "head"
  ' Номер колонки «№ вопроса»
  objBankID.Add Номер_вопроса, "QNum"
  ' Номер колонки «Поставщик (кратко)»
  objBankID.Add Имя_Контрагента, "NameS"
  ' Номер колонки, к которой находится «Дата поступления»
  objBankID.Add Дата_Поступления, "Date_mail"
  ' Номер колонки «Дата передачи аутсорсерам»
  objBankID.Add Дата_Передачи_аутсорсерам, "Date_OSend"
  ' Номер колонки, к которой находится «Дата акта»
  objBankID.Add Дата_Акта, "Date_akt"
  ' Номер колонки, к которой находится «Номер акта»
  objBankID.Add Номер_Акта, "Num_akt"
  ' Номер колонки, к которой находится «Дата договора»
  objBankID.Add Дата_Договора, "Date_dog"
  ' Номер колонки, к которой находится «Номер договора»
  objBankID.Add Номер_Договора, "Num_dog"
  ' Номер колонки, к которой находится «Дата перечислений»
  objBankID.Add Дата_Перечислений, "Date_APay"
  ' Номер колонки, к которой находится «Итого...» Сумма по документу
  objBankID.Add Сумма_Итого, "Sum_All"
  
  For Each objNamed In ActiveWorkbook.Names
    With objNamed
      On Error Resume Next
        bank = Empty: bank = Left(.Name, InStr(.Name, "_") - 1)
        field = Right(.Name, Len(.Name) - Len(bank) - 1)
        ' Если появляется Банк ...
        If Len(bank) = 2 Then
          ' ... смотрим, является ли Банк новым И Ссылка не битая
          If Bank_Key(Bank_Key.Count) <> "_" & bank _
          And Not .Value Like "*[#]*" Then
  '          Debug.Print ActiveWorkbook.Name & ": " & PRE & bank
            Bank_HeaderRow.Add .RefersToRange.Row, PRE & bank
            ' ЛУЧШЕ проверять по Worksheets.CodeName (read-only),
            ' если структура Книги создаётся с нуля
            Bank_SheetID.Add GetSheetID(.Value, False), PRE & bank
            ' Вписываем новый Банк и имя листа, на котором он находится
            Bank_Key.Add "_" & bank, PRE & bank
            ' Активируем лист, на котором расположен Банк
            If Worksheets(Bank_SheetID(PRE & bank)).Visible < 0 Then _
              Worksheets(Bank_SheetID(PRE & bank)).Select
          End If
        End If
        If GetSheetID(PRE, False) > 0 Then ' HotFix!
          Select Case field
            Case "Quant_inbox": field = "AMT_source"
            Case "Quant_new": field = "AimAMT"
            Case "Quant_In": field = "AimAMT_gb"
            Case "Goszak_In": field = "AimAMT_gz"
            Case "Quant_pay": field = "AcceptAMT"
            Case "Quant_Out": field = "AcceptAMT_gb"
            Case "Goszak_Out": field = "AcceptAMT_gz"
          End Select
        End If
        ' Если диапазон из одного значения
        If .RefersToRange.Count = 1 And Len(bank) = 2 Then
          objBankID(field).Add .RefersToRange.Column, PRE & bank
        ElseIf bank = "PART" Then
          If objBankID(field).Item(PRE & "OT") And Err.Number = 5 Then ' HotFix!
            objBankID(field).Add .RefersToRange.Column, PRE & "BO"
            objBankID(field).Add .RefersToRange.Column, PRE & "KF"
          Else ' after := PRE & "OT"
            objBankID(field).Add .RefersToRange.Column, PRE & "BO", PRE & "OT"
            objBankID(field).Add .RefersToRange.Column, PRE & "KF", PRE & "OT"
          End If
        ElseIf bank = "ARCH" Or bank = "SUPP" Then
  '        Debug.Print ActiveWorkbook.Name & ": " & field
          ' Заносятся реквизиты контрагентов...
          objBankSUPP.Add .RefersToRange.Column, field
          ' ... и Номер строки заголовка реквизитов «HEAD»
          If Not objBankSUPP Is Nothing And field = "NameS" Then
            objBankSUPP.Add GetSheetID(.Value, False), "sheet"
            objBankSUPP.Add .RefersToRange.Row, "head"
            With Worksheets(objBankSUPP("sheet")) ' Данные контрагентов
              bank = .Cells.SpecialCells(xlLastCell).Row
              While IsEmpty(.Cells(bank, objBankSUPP("NameS")))
                bank = bank - 1
              Wend: objBankSUPP.Add .Range(.Cells(objBankSUPP("head") + 1, 1), _
                .Cells(bank, .Cells.SpecialCells(xlLastCell).Column)).Value2, _
                "Data"
            End With
          End If
        End If
        ' Если имя диапазона содержит Date*, То применить формат HotFix!
        If field Like "Date*" And .RefersToRange.Count = 1 Then
          With Worksheets(Bank_SheetID(PRE & bank)).Cells(1, _
          .RefersToRange.Column).EntireColumn.Interior
            .NumberFormat = "m/d/yyyy"
            '.ColorIndex = 44 ' Отключить при открытии файла для редактирования
          End With
        End If
        
        If Err.Number = 1004 And Not IsEmpty(bank) Then
          On Error GoTo 0
          MsgBox "В книге """ & ActiveWorkbook.Name & """ поломан именованный " _
            & "диапазон """ & .Name & """, либо лист защищён от записи. " _
            & vbCr & "Для просмотра связей используйте комбинацию кнопок " _
            & "Ctrl+F3. ", vbCritical: End
        End If
    End With
  Next objNamed: Set objNamed = Nothing
End Sub

Function GetSupplerRec(ByVal SuppName As String, ByVal CheckDate As Variant, _
  Optional ByVal IsSelectSupp_Forced As Boolean = False) As Integer
  Attribute GetSupplerRec.VB_Description = "r311 ¦ Поиск записи контрагента"
  Dim tRelevant As Double, aU As Variant, eZ As Byte, nZ As Integer

  If IsNumeric(CheckDate) And Not IsEmpty(SuppName) Then
    If CheckDate > 0 Then
      SuppName = Trim(SuppName)
      aU = Sheets(xSUPP("sheet")).Cells(xSUPP("head") + 1, 1) _
        .CurrentRegion.Value2: eZ = xSUPP("NameS")
      aU = xSUPP("Data"): eZ = xSUPP("NameS")
      For nZ = LBound(aU, 1) To UBound(aU, 1)
        If aU(nZ, eZ) = SuppName Then
  '        Debug.Print "Рецензент " & SuppName & " найден в стороке " & nZ
  '        Debug.Print CDate(CheckDate) & " >= " & CDate(aU(nZ, xSUPP("DateD"))) & " >= " & CDate(tRelevant)
          
          If IsSelectSupp_Forced And GetSupplerRec = 0 Then _
            tRelevant = -aU(nZ, xSUPP("DateD"))
          
          ' Если требуется найти Имя контрагента И Дату актуальности, То
          If aU(nZ, xSUPP("DateD")) >= tRelevant Then
            If CheckDate >= aU(nZ, xSUPP("DateD")) Then
              tRelevant = aU(nZ, BankSUPP("DateD")): GetSupplerRec = nZ
            
            ' Если требуется найти хотя бы Имя контрагента, То
            ElseIf IsSelectSupp_Forced _
            And -aU(nZ, xSUPP("DateD")) >= tRelevant Then
              tRelevant = -aU(nZ, xSUPP("DateD")): GetSupplerRec = nZ
            End If
          End If
        End If
      Next nZ: GetSupplerRec = GetSupplerRec + BankSUPP("head") - 1
    Else
      Debug.Print CheckDate & " isn't a Date"
    End If
  End If
End Function

Function GetRecord(ByVal CurrentRow As Integer, ByVal BankID As String, _
  ByVal SheetID As Byte, Optional ByVal Prefix_ShName As String = "STAT") _
  As Variant ' Если "CurrentRow < 1" искать заголовок коллекции xID
  Attribute GetRecord.VB_Description = "r311 ¦ Значение ячейки из коллекции xID"
  Dim CodeName As String: CodeName = Prefix_ShName & xID("key").Item(SheetID)
  
  ' Если "CurrentRow > 0", то взять ЗНАЧЕНИЕ ячейки по коллекции xID
  GetRecord = xID(BankID).item(CodeName): If CurrentRow > 0 Then GetRecord = _
    Worksheets(xID("sheet").item(CodeName)).Cells(CurrentRow, GetRecord).Value2
End Function

Function GetSheetID(ByVal SheetCodeName As String, Optional ByRef ThisBook _
  As Boolean = True) As Byte ' ThisBook - ДА, эта книга
  Attribute GetSheetID.VB_Description = "r310 ¦ Найти индекс листа по CodeName"
  Dim objBook As Workbook, objSheet As Worksheet
  
  Set objBook = IIf(ThisBook, ThisWorkbook, ActiveWorkbook)
  If InStr(SheetCodeName, "!") > 0 Then _
    SheetCodeName = Replace(Mid(SheetCodeName, 2, InStr( _
      SheetCodeName, "!") - 2), "'", "") ' Имя листа должно быть БЕЗ "!"
  For Each objSheet In objBook.Worksheets
    If InStr(1, objSheet.CodeName, SheetCodeName, vbTextCompare) _
    Or InStr(1, objSheet.Name, SheetCodeName, vbTextCompare) Then _
      GetSheetID = objSheet.Index: Exit For
  Next objSheet: Set objSheet = Nothing: Set objBook = Nothing
End Function

Public Sub DeleteModulesAndCode(ByRef objBook As Object) ' Удалить модули
  Attribute DeleteModulesAndCode.VB_Description = "r270 ¦ Удалить модули из книги"
  ' Центр управления безопасностью -> Параметры макросов -> Доверять доступ к VBA
  Dim objProjectComponents As Object, objElement As Object
  
  Set objProjectComponents = objBook.VBProject.VBComponents
  For Each objElement In objProjectComponents
    Select Case objElement.Type
      Case 1 To 3
        objProjectComponents.Remove objElement
      Case 100
        With objElement.CodeModule
          .DeleteLines 1, .CountOfLines
        End With
    End Select
  Next objElement
End Sub
