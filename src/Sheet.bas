Attribute VB_Name = "Sheet"
Option Explicit
Option Base 1
'123456789012345678901234567890123456h8nor@уа56789012345678901234567890123456789

Private Const PRE = "STAT_"

Public Sub GetBanks(ByRef Ref_ID As Collection, _
  Optional ByRef Ref_SUPP As Collection)
  Attribute GetBanks.VB_Description = "r316 ¦ Определение структуры файла статистики"
  If Not Ref_ID.Count = 0 Then Exit Sub ' HoxFix!
  Dim objNamed As Object, bank As String, field As String
  ' ВАЖНАЯ ЧАСТЬ! Заголовки найденных именованных диапазонов в коллекции
  Dim Банк_Ключ As New Collection ' r314
  Dim Банк_N_Листа As New Collection ' r314
  Dim Банк_N_Строки_заголовка As New Collection ' r314
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
  
  Ref_ID.Add Банк_Ключ, "key" ' r314
  ' Коллекция с индексами листов для каждого Банка
  ' Включаемые Банки: СВ, ПВ, П+А, БО, КФ, ЮП, VIP, SEO, ЛК, РЦ
  Ref_ID.Add Банк_N_Листа, "sheet" ' r314
  ' Номер строки, в которой находятся заголовки таблицы банка «HEAD»
  Ref_ID.Add Банк_N_Строки_заголовка, "head" ' r314
  ' Номер колонки «№ вопроса»
  Ref_ID.Add Номер_вопроса, "QNum"
  ' Номер колонки «Поставщик (кратко)»
  Ref_ID.Add Имя_Контрагента, "NameS"
  ' Номер колонки, к которой находится «Дата поступления»
  Ref_ID.Add Дата_Поступления, "Date_mail"
  ' Номер колонки «Дата передачи аутсорсерам»
  Ref_ID.Add Дата_Передачи_аутсорсерам, "Date_OSend"
  ' Номер колонки, к которой находится «Дата акта»
  Ref_ID.Add Дата_Акта, "Date_akt"
  ' Номер колонки, к которой находится «Номер акта»
  Ref_ID.Add Номер_Акта, "Num_akt"
  ' Номер колонки, к которой находится «Дата договора»
  Ref_ID.Add Дата_Договора, "Date_dog"
  ' Номер колонки, к которой находится «Номер договора»
  Ref_ID.Add Номер_Договора, "Num_dog"
  ' Номер колонки, к которой находится «Дата перечислений»
  Ref_ID.Add Дата_Перечислений, "Date_APay"
  ' Номер колонки, к которой находится «Кол-во поступивших материалов»
  Ref_ID.Add Принято_На_проверку, "AimAMT"
  ' Номер колонки, к которой находится «Кол-во материалов после проверки»
  Ref_ID.Add Принято_После_проверки, "AcceptAMT"
  ' Номер колонки, к которой находится «Итого...» Сумма по документу
  Ref_ID.Add Сумма_Итого, "Sum_All"
  
  For Each objNamed In ActiveWorkbook.Names
    With objNamed
      If .Visible And TypeName(.Parent) = "Workbook" Then
        bank = Empty: field = Empty
        
  '      If .RefersTo Like "*[#]NAME[?]*" Or .RefersTo Like "*[#]REF!*" Then
  '        bank = .Parent.Worksheets(GetSheetID("STAT", False)).Index
        ' Внутренние диапазоны ' r313
        If .Name Like "_xl*" Then HookMsg "Системный диапазон " & .Name, vbOKCancel
        
        If .Name Like "*_*" Xor .Name Like "*!_*" Then
  '        Debug.Print .Name; " Is In Worksheet = "; .ValidWorkbookParameter
          On Error Resume Next
          '< <<<
          Select Case .RefersToRange.Count
            Case Is = 1
              bank = .RefersToRange.Worksheet.CodeName ' r313
              bank = IIf(Left(.Name, InStr(.Name, "_") - 1) = "SUPP", _
                "SUPP", Left(.Name, InStr(.Name, "_")) _
                & Mid(bank, InStr(bank, "_") + 1)) ' HotFix! ' r313
              field = Mid(.Name, InStr(.Name, "_") + 1) ' r313
          End Select
          '> >>>
          On Error GoTo 0
        End If
        
        If field Like "Quant_*" Or field Like "Goszak_*" Then ' HotFix!
          If GetSheetID(PRE, False) > 0 Then
            Select Case field
              Case Is = "Quant_inbox": field = "AMT_seed" ' r314
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
          '< <<<
          ' ... смотрим, является ли Банк новым И Ссылка не битая
          If Банк_Ключ(Банк_Ключ.Count) <> IIf(Len(bank) = 2, "_", "") & bank Then ' r313
'            Debug.Print ActiveWorkbook.Name & ": " & bank
            Банк_N_Строки_заголовка.Add .RefersToRange.Row, bank
            ' ЛУЧШЕ проверять по Worksheets.CodeName (ReadOnly),
            ' если структура Книги создаётся с нуля
            Банк_N_Листа.Add GetSheetID(.Value, False), bank
            ' Вписываем новый Банк и имя Листа, на котором он находится
            Банк_Ключ.Add IIf(Len(bank) = 2, "_", "") & bank, bank
            ' Активируем лист, на котором расположен Банк
            If Worksheets(Банк_N_Листа(bank)).Visible < 0 Then _
              Worksheets(Банк_N_Листа(bank)).Activate
          End If
          
          Ref_ID(field).Add .RefersToRange.Column, bank
          '> >>>
          On Error GoTo 0
        ElseIf bank = "PART" Then ' НЕ ТРЕБУЕТСЯ
          On Error Resume Next
          '< <<<
          If Ref_ID(field).item(PRE & "OT") And Err.Number = 5 Then ' HotFix!
            Ref_ID(field).Add .RefersToRange.Column, PRE & "BO"
            Ref_ID(field).Add .RefersToRange.Column, PRE & "KF"
          Else ' after := PRE & "OT"
            Ref_ID(field).Add .RefersToRange.Column, PRE & "BO", PRE & "OT"
            Ref_ID(field).Add .RefersToRange.Column, PRE & "KF", PRE & "OT"
          End If
          '> >>>
          On Error GoTo 0
        ElseIf bank = "ARCH" Or bank = "SUPP" Then
  '        Debug.Print ActiveWorkbook.Name & ": " & field
          ' Заносятся реквизиты контрагентов...
          Ref_SUPP.Add .RefersToRange.Column, field
          ' ... и Номер строки заголовка реквизитов «HEAD»
          If Not Ref_SUPP Is Nothing And field = "NameS" Then
            Ref_SUPP.Add GetSheetID(.Value, False), "sheet"
            Ref_SUPP.Add .RefersToRange.Row, "head"
            With Worksheets(Ref_SUPP("sheet")) ' Данные контрагентов
              If .AutoFilterMode Then
                If .AutoFilter.FilterMode Then .ShowAllData
                ' https://msdn.microsoft.com/ru-ru/vba/excel-vba
                '   /articles/range-sort-method-excel
                '   /articles/xlsortdataoption-enumeration-excel
                
                ' Реверс. "Сорт. диапазона" перед внесением "Data" в коллекцию
                .Cells(Ref_SUPP("head"), Ref_SUPP("NameS")).Sort _
                  Key1:=.Cells(Ref_SUPP("head"), Ref_SUPP("NameS")), _
                  Order1:=xlAscending, Key2:=.Cells(Ref_SUPP("head"), _
                  Ref_SUPP("DateD")), Order2:=xlDescending, Header:=xlYes
              End If
              bank = .Cells.SpecialCells(xlLastCell).Row
              While IsEmpty(.Cells(bank, Ref_SUPP("NameS")))
                bank = bank - 1
              Wend: Ref_SUPP.Add .Range(.Cells(Ref_SUPP("head") + 1, 1), _
                .Cells(bank, .Cells.SpecialCells(xlLastCell).Column)).Value2, _
                "Data"
            End With
          End If
        End If
        
        On Error Resume Next
        '< <<<
        ' Если имя диапазона содержит Date*, То применить формат HotFix!
        If field Like "Date*" Then 'And Not .ProtectContents Then
          With Worksheets(Банк_N_Листа(bank))
            With .Cells(1, objNamed.RefersToRange.Column).EntireColumn
              .NumberFormat = "m/d/yyyy"
              '.Interior.ColorIndex = 44 ' Отключить при открытии файла для редактирования
            End With
          End With
        End If
        
        If Err.Number = 1004 And Not Len(bank) > 0 Then
          HookMsg "В книге """ & ActiveWorkbook.Name & """ поломан именованный" _
            & " диапазон """ & .Name & """, либо лист защищён от записи. " _
            & vbCr & "Для просмотра связей используйте комбинацию кнопок " _
            & "Ctrl+F3. ", vbCritical: End ' r314
        End If
        '> >>>
        On Error GoTo 0
      End If
    End With
  Next objNamed: Set objNamed = Nothing
End Sub

Function GetSupplerRec(ByVal suppName As String, ByVal checkDate As Variant, _
  Optional ByVal isForce_SearchSupp As Boolean = False) As Integer
  Attribute GetSupplerRec.VB_Description = "r316 ¦ Поиск записи контрагента"
  Dim tRelevant As Double, aU As Variant, eZ As Byte, nZ As Integer
  
  If IsNumeric(checkDate) And Not IsEmpty(suppName) Then
    If checkDate > 0 Then
      suppName = Trim(suppName)
      aU = Sheets(xSUPP("sheet")).Cells(xSUPP("head") + 1, 1) _
        .CurrentRegion.Value2: eZ = xSUPP("NameS")
      aU = xSUPP("Data"): eZ = xSUPP("NameS")
      For nZ = LBound(aU, 1) To UBound(aU, 1)
        If aU(nZ, eZ) = suppName Then
  '        Debug.Print "Рецензент " & suppName & " найден в стороке " & nZ
  '        Debug.Print CDate(checkDate) & " >= " & CDate(aU(nZ, xSUPP("DateD"))) & " >= " & CDate(tRelevant)
          
          If isForce_SearchSupp And GetSupplerRec = 0 Then _
            tRelevant = -aU(nZ, xSUPP("DateD"))
          
          ' Если требуется найти Имя контрагента И Дату актуальности, То
          If aU(nZ, xSUPP("DateD")) >= tRelevant Then
            If checkDate >= aU(nZ, xSUPP("DateD")) Then
              tRelevant = aU(nZ, xSUPP("DateD")): GetSupplerRec = nZ
            
            ' Если требуется найти хотя бы Имя контрагента, То
            ElseIf isForce_SearchSupp _
            And -aU(nZ, xSUPP("DateD")) >= tRelevant Then
              tRelevant = -aU(nZ, xSUPP("DateD")): GetSupplerRec = nZ
            End If
          End If
        End If
      Next nZ: GetSupplerRec = GetSupplerRec + xSUPP("head") - 1
    Else
      HookMsg checkDate & " isn't a Date", vbOKCancel
    End If
  End If
End Function


Public Function GetRecord(ByRef Ref_Row As Integer, ByVal keyID As String, _
  Optional ByVal sheetIndex As Byte, Optional ByVal bankKeyIndex As Byte) As Variant
  Attribute GetRecord.VB_Description = "r314 ¦ Значение ячейки из коллекции xID"
  Dim CodeName As Variant ' Если "Ref_Row < 1" искать заголовок коллекции xID
  
  Select Case bankKeyIndex
    Case Is > 0: CodeName = xID("key").item(bankKeyIndex)
    Case Else
      For Each CodeName In xID("key")
        If xID("sheet").item(CodeName) = sheetIndex Then Exit For
      Next CodeName
      If IsEmpty(CodeName) Then Exit Function ' HotFix!
  End Select
  
  ' Если "Ref_Row > 0", то взять ЗНАЧЕНИЕ ячейки по коллекции xID
  GetRecord = xID(keyID).item(CodeName): If Ref_Row > 0 Then GetRecord = _
    Worksheets(xID("sheet").item(CodeName)).Cells(Ref_Row, GetRecord).Value2
End Function

Function GetSheetID(ByVal sheetCodeName As String, _
  Optional ByRef Ref_isThisBook As Boolean = True) As Byte ' ThisBook - ДА, эта книга
  Attribute GetSheetID.VB_Description = "r314 ¦ Найти индекс листа по CodeName"
  Dim objBook As Workbook, objSheet As Worksheet
  
  Set objBook = IIf(Ref_isThisBook, ThisWorkbook, ActiveWorkbook)
  If InStr(sheetCodeName, "!") > 0 Then _
    sheetCodeName = Replace(Mid(sheetCodeName, 2, InStr( _
      sheetCodeName, "!") - 2), "'", "") ' Имя листа должно быть БЕЗ "!"
  For Each objSheet In objBook.Worksheets
    If InStr(1, objSheet.CodeName, sheetCodeName, vbTextCompare) _
    Or InStr(1, objSheet.Name, sheetCodeName, vbTextCompare) Then _
      GetSheetID = objSheet.Index: Exit For
  Next objSheet: Set objSheet = Nothing: Set objBook = Nothing
End Function

Public Sub HookMsg(ByVal promt As Variant, ByVal style As VbMsgBoxStyle, _
  Optional title As String)
  Attribute HookMsg.VB_Description = "r316 ¦ Вывод сообщений в окно 'Immediate' в режиме IS_DEBUG"
  If style > vbRetryCancel And IS_DEBUG Then style = vbRetryCancel
  Select Case style
    Case Is = vbOKCancel
      Debug.Print promt
    Case Is = vbRetryCancel
      Debug.Print IIf(IS_DEBUG, "[DEBUG MODE] ", ""); promt
    Case Else
      MsgBox promt, style, IIf(Trip(Len(title)) > 0, title, Application.Name)
  End Select
End Sub

Public Sub DeleteModulesAndCode(ByRef Ref_Book As Object) ' Удалить модули
  Attribute DeleteModulesAndCode.VB_Description = "r314 ¦ Удалить модули из книги"
  ' Центр управления безопасностью -> Параметры макросов -> Доверять доступ к VBA
  Dim objProjectComponents As Object, objElement As Object
  
  On Error Resume Next
  '< <<<
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
  '> >>>
  On Error Goto 0
  
  With Ref_Book
    .Parent.DisplayAlerts = False
    Set objElement = .Sheets(.Sheets.Count)
    If objElement.Name = "Temp0" Then objElement.Delete
    .Parent.DisplayAlerts = True
  End With
End Sub
