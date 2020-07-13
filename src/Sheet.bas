Attribute VB_Name = "Sheet"
Option Explicit
Option Base 1
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890

Private Const PRE = "STAT_"

Public Sub GetBanks(ByRef objBankID As Collection, Optional ByRef objBankSUPP _
  As Collection)
  Attribute GetBanks.VB_Description = "r313 ¦ Определение структуры файла статистики"
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
  Dim Принято_На_проверку As New Collection
  Dim Принято_После_проверки As New Collection
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
  ' Номер колонки, к которой находится «Кол-во поступивших материалов»
  objBankID.Add Принято_На_проверку, "AimAMT"
  ' Номер колонки, к которой находится «Кол-во материалов после проверки»
  objBankID.Add Принято_После_проверки, "AcceptAMT"
  ' Номер колонки, к которой находится «Итого...» Сумма по документу
  objBankID.Add Сумма_Итого, "Sum_All"
  
  For Each objNamed In ActiveWorkbook.Names
    With objNamed
      If .Visible And TypeName(.Parent) = "Workbook" Then
        bank = Empty: field = Empty
        
  '      If .RefersTo Like "*[#]NAME[?]*" Or .RefersTo Like "*[#]REF!*" Then
  '        bank = .Parent.Worksheets(GetSheetID("STAT", False)).Index
        ' Внутренние диапазоны ' r313
        If .Name Like "_xl*" Then Debug.Print "Системный диапазон "; .Name
        
        If .Name Like "*_*" Xor .Name Like "*!_*" Then
  '        Debug.Print .Name; " Is In Worksheet = "; .ValidWorkbookParameter
          On Error Resume Next
            Select Case .RefersToRange.Count
              Case Is = 1
                bank = .RefersToRange.Worksheet.CodeName ' r313
                bank = IIf(Left(.Name, InStr(.Name, "_") - 1) = "SUPP", _
                  "SUPP", Left(.Name, InStr(.Name, "_")) _
                  & Mid(bank, InStr(bank, "_") + 1)) ' HotFix! ' r313
                field = Mid(.Name, InStr(.Name, "_") + 1) ' r313
            End Select
          On Error GoTo 0
        End If
        
        If field Like "Quant_*" Or field Like "Goszak_*" Then ' HotFix!
        If GetSheetID(PRE, False) > 0 Then
          Select Case field
            Case Is = "Quant_inbox": field = "AMT_source"
            Case Is = "Quant_new": field = "AimAMT"
            Case Is = "Quant_In": field = "AimAMT_gb"
            Case Is = "Goszak_In": field = "AimAMT_gz"
            Case Is = "Quant_pay": field = "AcceptAMT"
            Case Is = "Quant_Out": field = "AcceptAMT_gb"
            Case Is = "Goszak_Out": field = "AcceptAMT_gz"
          End Select
        End If
        End If
        
        ' Если появляется Банк ...
        If Len(bank) = 2 or Len(bank) = 5 Then ' r313
          On Error Resume Next
            ' ... смотрим, является ли Банк новым И Ссылка не битая
            If Bank_Key(Bank_Key.Count) <> IIf(Len(bank) = 2, "_", "") & bank Then ' r313
  '            Debug.Print ActiveWorkbook.Name & ": " & bank
              Bank_HeaderRow.Add .RefersToRange.Row, bank
              ' ЛУЧШЕ проверять по Worksheets.CodeName (ReadOnly),
              ' если структура Книги создаётся с нуля
              Bank_SheetID.Add GetSheetID(.Value, False), bank
              ' Вписываем новый Банк и имя Листа, на котором он находится
              Bank_Key.Add IIf(Len(bank) = 2, "_", "") & bank, bank
              ' Активируем лист, на котором расположен Банк
              If Worksheets(Bank_SheetID(bank)).Visible < 0 Then _
                Worksheets(Bank_SheetID(bank)).Activate
            End If
            
            objBankID(field).Add .RefersToRange.Column, bank
          On Error GoTo 0
        ElseIf bank = "PART" Then ' НЕ ТРЕБУЕТСЯ
          On Error Resume Next
          If objBankID(field).item(PRE & "OT") And Err.Number = 5 Then ' HotFix!
            objBankID(field).Add .RefersToRange.Column, PRE & "BO"
            objBankID(field).Add .RefersToRange.Column, PRE & "KF"
          Else ' after := PRE & "OT"
            objBankID(field).Add .RefersToRange.Column, PRE & "BO", PRE & "OT"
            objBankID(field).Add .RefersToRange.Column, PRE & "KF", PRE & "OT"
          End If
          On Error GoTo 0
        ElseIf bank = "ARCH" Or bank = "SUPP" Then
  '        Debug.Print ActiveWorkbook.Name & ": " & field
          ' Заносятся реквизиты контрагентов...
          objBankSUPP.Add .RefersToRange.Column, field
          ' ... и Номер строки заголовка реквизитов «HEAD»
          If Not objBankSUPP Is Nothing And field = "NameS" Then
            objBankSUPP.Add GetSheetID(.Value, False), "sheet"
            objBankSUPP.Add .RefersToRange.Row, "head"
            With Worksheets(objBankSUPP("sheet")) ' Данные контрагентов
              If .AutoFilterMode Then
                If .AutoFilter.FilterMode Then .ShowAllData
                ' https://msdn.microsoft.com/ru-ru/vba/excel-vba
                '   /articles/range-sort-method-excel
                '   /articles/xlsortdataoption-enumeration-excel
                
                ' Реверс. "Сорт. диапазона" перед внесением "Data" в коллекцию
                .Cells(objBankSUPP("head"), objBankSUPP("NameS")).Sort _
                  Key1:=.Cells(objBankSUPP("head"), objBankSUPP("NameS")), _
                  Order1:=xlAscending, Key2:=.Cells(objBankSUPP("head"), _
                  objBankSUPP("DateD")), Order2:=xlDescending, Header:=xlYes
              End If
              bank = .Cells.SpecialCells(xlLastCell).Row
              While IsEmpty(.Cells(bank, objBankSUPP("NameS")))
                bank = bank - 1
              Wend: objBankSUPP.Add .Range(.Cells(objBankSUPP("head") + 1, 1), _
                .Cells(bank, .Cells.SpecialCells(xlLastCell).Column)).Value2, _
                "Data"
            End With
          End If
        End If
        
        On Error Resume Next
        ' Если имя диапазона содержит Date*, То применить формат HotFix!
        If field Like "Date*" Then 'And Not .ProtectContents Then
          With Worksheets(Bank_SheetID(bank))
            With .Cells(1, objNamed.RefersToRange.Column).EntireColumn
              .NumberFormat = "m/d/yyyy"
              '.Interior.ColorIndex = 44 ' Отключить при открытии файла для редактирования
            End With
          End With
        End If
        
        If Err.Number = 1004 And Not Len(bank) > 0 Then
          MsgBox "В книге """ & ActiveWorkbook.Name & """ поломан именован" _
            & "ный диапазон """ & .Name & """, либо лист защищён от записи. " _
            & vbCr & "Для просмотра связей используйте комбинацию кнопок " _
            & "Ctrl+F3. ", vbCritical: End
        End If
        On Error GoTo 0
      End If
    End With
  Next objNamed: Set objNamed = Nothing
End Sub

Function GetSupplerRec(ByVal SuppName As String, ByVal CheckDate As Variant, _
  Optional ByVal IsSelectSupp_Forced As Boolean = False) As Integer
  Attribute GetSupplerRec.VB_Description = "r313 ¦ Поиск записи контрагента"
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
              tRelevant = aU(nZ, xSUPP("DateD")): GetSupplerRec = nZ
            
            ' Если требуется найти хотя бы Имя контрагента, То
            ElseIf IsSelectSupp_Forced _
            And -aU(nZ, xSUPP("DateD")) >= tRelevant Then
              tRelevant = -aU(nZ, xSUPP("DateD")): GetSupplerRec = nZ
            End If
          End If
        End If
      Next nZ: GetSupplerRec = GetSupplerRec + xSUPP("head") - 1
    Else
      Debug.Print CheckDate & " isn't a Date"
    End If
  End If
End Function


Public Function GetRecord(ByRef CurrentRow As Integer, ByVal KeyID As String, _
  Optional ByVal SheetIndex As Byte, Optional ByVal BankIndex As Byte) As Variant
  Attribute GetRecord.VB_Description = "r313 ¦ Значение ячейки из коллекции xID"
  Dim CodeName As Variant ' Если "CurrentRow < 1" искать заголовок коллекции xID
  
  Select Case BankIndex
    Case Is > 0: CodeName = xID("key").item(BankIndex)
    Case Else
      For Each CodeName In xID("key")
        If xID("sheet").item(CodeName) = SheetIndex Then Exit For
      Next CodeName
      If IsEmpty(CodeName) Then Exit Function ' HotFix!
  End Select
  
  ' Если "CurrentRow > 0", то взять ЗНАЧЕНИЕ ячейки по коллекции xID
  GetRecord = xID(KeyID).item(CodeName): If CurrentRow > 0 Then GetRecord = _
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
  Attribute DeleteModulesAndCode.VB_Description = "r313 ¦ Удалить модули из книги"
  ' Центр управления безопасностью -> Параметры макросов -> Доверять доступ к VBA
  Dim objProjectComponents As Object, objElement As Object
  
  On Error Resume Next
  Set objProjectComponents = objBook.VBProject.VBComponents
  For Each objElement In objProjectComponents
    Select Case objElement.Type
      Case 1 To 3
        objProjectComponents.Remove objElement
      Case Is = 100
        With objElement.CodeModule
          .DeleteLines 1, .CountOfLines
        End With
    End Select
  Next objElement
  On Error Goto 0
  
  With objBook
    .Parent.DisplayAlerts = False
    Set objElement = .Sheets(.Sheets.Count)
    If objElement.Name = "Temp0" Then objElement.Delete
    .Parent.DisplayAlerts = True
  End With
End Sub
