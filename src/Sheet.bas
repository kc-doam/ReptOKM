Attribute VB_Name = "Sheet"
Option Explicit
Option Base 1
'123456789012345678901234567890123456h8nor@уа56789012345678901234567890123456789

Private LastRow as Long ' r317

Public Sub GetBanks(ByRef Ref_ID As Collection, _
  Optional ByRef Ref_SUPP As Collection)
  Attribute GetBanks.VB_Description = "r318 ¦ Определение структуры файла статистики"
  If Not Ref_ID.Count = 0 Then Exit Sub ' HoxFix!
  Dim objNamed As Object, bank As String, field As String
  ' ВАЖНАЯ ЧАСТЬ! Заголовки найденных именованных диапазонов в коллекции
  Dim Банк_Ключ As New Collection ' r314
  Dim Банк_N_Листа As New Collection ' r314
  Dim Банк_N1_Записи As New Collection ' r317
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
  Ref_ID.Add Банк_N1_Записи, "line" ' r317
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
          HookMsg .Name & " Is In Worksheet = " & .ValidWorkbookParameter, vbRetryCancel
          ' Вероятнее всего банк имеет поле 'Comment'  ' r317
          If field = "Comment" Then HookMsg "Банк" & Replace(bank, _
            "_PART", "_SF") & " имеет Комментарий", vbRetryCancel ' r317
          On Error Resume Next
          '< <<<
          Select Case .RefersToRange.Count
            Case Is = 1
              bank = .RefersToRange.Worksheet.CodeName ' r313
              bank = Replace(bank, "_PART", "_SF") ' r317
              bank = IIf(Left(.Name, InStr(.Name, "_") - 1) = "SUPP", _
                "SUPP", Left(.Name, InStr(.Name, "_")) _
                & Mid(bank, InStr(bank, "_") + 1)) ' HotFix! ' r313
              field = Mid(.Name, InStr(.Name, "_") + 1) ' r313
          End Select
          '> >>>
          On Error GoTo 0
        End If
        
        If field Like "Quant_*" Or field Like "Goszak_*" Then ' HotFix!
        ''  If GetSheetID(PRE, False) > 0 Then ' r317
            Select Case field
              Case Is = "Quant_inbox": field = "AMT_seed" ' r314
              Case Is = "Quant_new": field = "AimAMT"
              Case Is = "Quant_In": field = "AimAMT_gb"
              Case Is = "Goszak_In": field = "AimAMT_gz"
              Case Is = "Quant_pay": field = "AcceptAMT"
              Case Is = "Quant_Rec": field = "AcceptAMT" ' QT ' r318
              Case Is = "Quant_Out": field = "AcceptAMT_gb"
              Case Is = "Goszak_Out": field = "AcceptAMT_gz"
            End Select
        ''  End If ' r317
        End If
        
        ' Если появляется Банк ...
        If Len(bank) = 5 Then ' r317
          On Error Resume Next
          '< <<<
          ' ... смотрим, является ли Банк новым И Ссылка не битая
          If Банк_Ключ(Банк_Ключ.Count) <> IIf(Len(bank) = 2, "_", "") & bank Then ' r313
            HookMsg ActiveWorkbook.Name & ": " & bank, vbRetryCancel

            Банк_N1_Записи.Add .RefersToRange.Row + 1, bank ' r317
            ' ЛУЧШЕ проверять по Worksheets.CodeName (ReadOnly),
            ' если структура Книги создаётся с нуля
            Банк_N_Листа.Add GetSheetID(.Value, False), bank
            ' Вписываем новый Банк и имя Листа, на котором он находится
            Банк_Ключ.Add IIf(Len(bank) = 2, "_", "") & bank, bank
            ' Активируем лист, на котором расположен Банк
            'If Worksheets(Банк_N_Листа(bank)).Visible < 0 Then _
              Worksheets(Банк_N_Листа(bank)).Activate
          End If
          
          Ref_ID(field).Add .RefersToRange.Column, bank
          eZ = Банк_N_Листа(bank) ' r317
          '> >>>
          On Error GoTo 0

        ElseIf bank Like "PART_*" Then ' r317
          On Error Resume Next
          '< <<<
          If Ref_ID(field).Item("OT_OT") And Err.Number = 5 Then ' HotFix!
            Ref_ID(field).Add .RefersToRange.Column, "BO_SF"
            Ref_ID(field).Add .RefersToRange.Column, "KF_SF"
          Else ' after := PRE & "OT"
            Ref_ID(field).Add .RefersToRange.Column, "BO_SF", "OT_OT"
            Ref_ID(field).Add .RefersToRange.Column, "KF_SF", "OT_OT"
          End If
          '> >>>
          On Error GoTo 0
          
        ElseIf bank Like "*_PART" Then ' r317
          Stop '' bank = "PART" Then <- ? НЕ ТРЕБУЕТСЯ, как и PRE ' r317
          On Error Resume Next
          '< <<<
          If Ref_ID(field).Item("STAT_OT") And Err.Number = 5 Then ' HotFix! ' r317
            Ref_ID(field).Add .RefersToRange.Column, PRE & "BO"
            Ref_ID(field).Add .RefersToRange.Column, PRE & "KF"
          Else ' after := PRE & "OT"
            Ref_ID(field).Add .RefersToRange.Column, PRE & "BO", PRE & "OT"
            Ref_ID(field).Add .RefersToRange.Column, PRE & "KF", PRE & "OT"
          End If
          '> >>>
          On Error GoTo 0
          
        ElseIf bank = "ARCH" Or bank = "SUPP" Then ' r317
          HookMsg ActiveWorkbook.Name & ": " & field, vbRetryCancel

          ' Внести реквизиты контрагентов
          If Ref_SUPP.Count = 0 Then Ref_SUPP.Add GetSheetID(.Value, False), "sheet"
          Ref_SUPP.Add .RefersToRange.Column, field
          ' Внести номер строки заголовка реквизитов «HEAD» ' r317
          If Not Ref_SUPP Is Nothing And field = "NameS" Then
            Ref_SUPP.Add .RefersToRange.Row + 1, "line"
            
            With Worksheets(Ref_SUPP("sheet")) ' Данные контрагентов ' r317
              LastRow = .Cells.SpecialCells(xlCellTypeLastCell).Row
              Call SortBySheet(Worksheets(.index), Ref_SUPP("DateD"), Ref_SUPP("NameS"))
              
              While IsEmpty(.Cells(LastRow, Ref_SUPP("NameS")))
                LastRow = LastRow - 1
              Wend
              Ref_SUPP.Add .Range(.Cells(Ref_SUPP("line"), 1), .Cells(LastRow + 1,  _
                .Cells.SpecialCells(xlCellTypeLastCell).Column)).Value2, "Data"
            End With
          End If: eZ = Ref_SUPP("sheet") ' r317
        End If
        
        On Error Resume Next
        '< <<<
        ' Если имя диапазона содержит Date*, То применить формат HotFix!
        With Worksheets(eZ)
          If field Like "Date*" And IS_DEBUG Then ' And Not .ProtectContents Then
            With .Cells(1, objNamed.RefersToRange.Column).EntireColumn
              .NumberFormat = "m/d/yyyy"
              '.Interior.ColorIndex = 44 ' Отключить при открытии файла для редактирования
            End With
          End If
        End With
        
        If Err.Number = 1004 And Not Len(bank) > 0 Then
          HookMsg "В книге """ & ActiveWorkbook.Name & """ поломан именованный" _
            & " диапазон """ & .Name & """, либо лист защищён от записи. " _
            & vbCr & "Для просмотра связей используйте комбинацию кнопок " _
            & "Ctrl+F3. ", vbCritical: End
        End If
        '> >>>
        On Error GoTo 0
      End If
    End With
  Next objNamed: Set objNamed = Nothing
End Sub

Public Sub SortBySheet(ByRef Sh As Worksheet, ByVal FirstKey As Byte, _
  Optional ByVal SecondKey As Byte)
  Attribute SortBySheet.VB_Description = "r317 ¦ Сортировка через метод листа"
  
  If Not Sh.AutoFilterMode Then Sh.Cells(xSUPP("line"), FirstKey).AutoFilter
  With Sh.AutoFilter.Sort
    .SortFields.Clear: .Header = xlYes
    .SortFields.Add Key:=Sh.Cells(xSUPP("line"), FirstKey).Resize(LastRow, 1)
    If SecondKey > 0 Then _
      .SortFields.Add Key:=Sh.Cells(xSUPP("line"), SecondKey).Resize(LastRow, 1)
    .Orientation = xlTopToBottom: .Apply
  End With
End Sub

Public Function GetSupplerRec(ByVal suppName As String, ByVal checkDate As Variant, _
  Optional ByVal isForce_SearchSupp As Boolean = False) As Integer
  Attribute GetSupplerRec.VB_Description = "r317 ¦ Поиск записи контрагента"
  Dim tRelevant As Double, aU As Variant, eZ As Byte, nZ As Integer
  
  If IsNumeric(checkDate) And Not IsEmpty(suppName) Then
    If checkDate > 0 Then
      'suppName = Trim(suppName) ' ВАЖНО! Ещё есть 'лишние пробелы' в Названиях
      aU = Sheets(xSUPP("sheet")).Cells(xSUPP("line"), 1) _
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
            And aU(nZ, xSUPP("DateD")) >= tRelevant Then
              tRelevant = -aU(nZ, xSUPP("DateD")): GetSupplerRec = nZ
            End If
          End If
        End If
      Next nZ
      ' GetSupplerRec = GetSupplerRec + xSUPP("line") - 2 ' ТЕСТ ' r317
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

Public Function GetSheetID(ByVal sheetCodeName As String, _
  Optional ByRef Ref_isThisBook As Boolean = True) As Byte ' "ЭтаКнига" - ДА, эта книга
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

Public Sub HookMsg(ByVal promt As Variant, ByVal style As vbMsgBoxStyle, _
  Optional ByVal title As String)
  Attribute HookMsg.VB_Description = "r317 ¦ Вывод сообщений в окно 'Immediate' в режиме IS_DEBUG"
  Dim comma As Variant
  
  If style > vbRetryCancel And IS_DEBUG Then style = vbRetryCancel
  Select Case style

    Case Is = vbOKCancel
      Debug.Print promt
    
    Case Is = vbRetryCancel
      If IS_DEBUG Then
        promt = Split(Replace(promt, vbCr, vbTab), vbTab)
        Debug.Print "[DEBUG MODE] ";: For Each comma In promt
          Debug.Print comma,: Next comma: Debug.Print ""
        If Asc(Left(title & "@", 1)) < 48 Then Stop ' isTitle
      End If
    
    Case Else
      MsgBox promt, style, IIf(Trip(Len(title)) > 0, title, Application.Name)
    
  End Select
End Sub

Public Sub DeleteModulesAndCode(ByRef Ref_Book As Object)
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
