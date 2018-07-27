Attribute VB_Name = "Execute"
Option Explicit
Option Base 1
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890

Public Sub Shell_Sort(ByRef items As Variant, ByVal column As Byte)
  Attribute SettingsBankID.VB_Description = "r270 Сортировка методом Шелла"
  ' Сортировка методом Шелла Дональда с интервалами длин ЭМ Марцина Циура
  ' http://ru.wikibooks.org/wiki/Примеры_реализации_сортировки_Шелла#VBA
  Dim i As Integer, gap As Integer, pos As Integer, inv_a102549 As Variant
  Dim k As Integer, m As Integer, n As Integer
  
  ReDim tmp(UBound(items, 1)) As Variant
  m = LBound(items, 2) ' Нижняя граница массива
  n = UBound(items, 2)
  If n > m Then ' Если в массиве одна запись, То нет смысла сортировать
    ' Инвертированная эмпирическая последовательность Марцина Циура
    inv_a102549 = Array(1750, 701, 301, 132, 57, 23, 10, 4, 1)
    k = n - m + 1 ' Количество записей в массиве
    n = LBound(inv_a102549) - 1 ' Счётчик последовательности Марцина Циура
    Do
      n = n + 1
      gap = inv_a102549(n) ' Наименьший интервал
    Loop Until k > gap ' Цикл ДО
    Do
      gap = inv_a102549(n) ' Интервал
      For i = (gap + m) To UBound(items, 2)
        pos = i
        For k = LBound(items, 1) To UBound(items, 1)
          tmp(k) = items(k, i)
        Next k
        Do While items(column, pos - gap) > tmp(column) ' Выполнять ПОКА
          For k = LBound(items, 1) To UBound(items, 1)
            items(k, pos) = items(k, pos - gap)
            items(k, pos - gap) = tmp(k)
          Next k
          pos = pos - gap
          If (pos - gap) < m Then Exit Do
        Loop
      Next i
      n = n + 1
    Loop Until gap = 1 ' Цикл ДО
  Else
  '  Debug.Print "Shell_Sort: В массиве ОДНА или НУЛЬ записей"
  End If
End Sub

Public Function GetQuarterNumber(ByVal SetDate As Date, Optional _
  ByVal GetYear As Boolean = False) As String ' Номер квартала
  Attribute SettingsBankID.VB_Description = "r270 Заменить цифру на номер квартала"
  '  Select Case (Month(SetDate) - 1) \ 3 + 1
  '    Case 1: GetQuarterNumber = "I квартал "
  '    Case 2: GetQuarterNumber = "II квартал "
  '    Case 3: GetQuarterNumber = "III квартал "
  '    Case 4: GetQuarterNumber = "IV квартал "
  '  End Select
  '  If GetYear Then GetQuarterNumber = GetQuarterNumber & Year(SetDate) & " г."
  ' Вывод данных через массив медленнее на 25%, чем через Select Case
  Dim Quart As Variant: Quart = Array("I", "II", "III", "IV")
  
  If GetYear Then GetQuarterNumber = Quart((Month(SetDate) - 1) \ 3 + 1) _
    & " квартал " & Year(SetDate) & " г." Else GetQuarterNumber = _
    Quart((Month(SetDate) - 1) \ 3 + 1) & "квартал"
End Function


Public Function ClearSpacesInText(ByVal text As String) As String
  Attribute SettingsBankID.VB_Description = "r270 Удалить опечатки и лишние пробелы"
  text = Replace(text, Chr(160), " ")       ' неразрывный пробел
  ''text = Replace(text, ".", ". ") ' Ошибка при проверке "т.ч."
  text = Replace(text, "т.ч.", "т ч")
  text = Replace(text, """", "")            ' стереть кавычки
  ' МОЖЕТ ЛИШНЕЕ... СМ ' Костыль для ключа #2
  text = Replace(text, " - ", " ")          ' стереть "не-до-тире"
  text = Replace(text, " -", " ")           ' стереть "не-до-тире2"
  
  text = Replace(text, "-", " ")            ' стереть "не-до-тире3"
  
  ''text = Replace(text, " .", ". ")
  ''text = Replace(Replace(text, ",", ", "), " ,", ", ")
  ''text = Replace(Replace(text, "!", "! "), " !", "! ")
  ''text = Replace(Replace(text, "?", "? "), " ?", "? ")
  ''text = Replace(Replace(text, ":", ": "), " :", ": ")
  ''text = Replace(Replace(text, ";", "; "), " :", ": ")
  ''text = Replace(Replace(text, "( ", "("), " )", ")")
  ''text = Replace(text, Chr(10), " ")       ' Перевод строки
  ''text = Replace(text, Chr(13), " ")       ' Перевод каретки
  ''text = Replace(text, Chr(150), Chr(45))  ' Короткое тире
  ''text = Replace(text, Chr(151), Chr(45))  ' Длинное тире
  ''text = Replace(text, Chr(133), "...")    ' Многоточие
  ''text = Replace(text, Chr(172), "")       ' знак переноса
  Do While text Like "*  *" ' Выполнять ПОКА есть двойной пробел
    text = Replace(text, "  ", " ")
  Loop
  ''text = Replace(text, Chr(39), Chr(34))   ' апостроф
  ''text = Replace(text, Chr(171), Chr(34))  ' левые двойные кавычки
  ''text = Replace(text, Chr(187), Chr(34))  ' правые двойные кавычки
  ''text = Replace(text, Chr(147), Chr(34))  ' левые четверть круга кавычки
  ''text = Replace(text, Chr(148), Chr(34))  ' правые четверть круга кавычки
  ''text = Replace(text, Chr(34) & Chr(34), Chr(34))
  ClearSpacesInText = Trim(text)
End Function

Function FindSheet(ByVal FindSheetCodeName As String, Optional ByRef ThisBook _
  As Boolean = False) As Byte ' ThisBook - ДА, эта книга
  Attribute SettingsBankID.VB_Description = "r300 Найти индекс листа по CodeName"
  Dim GetBook As Workbook, GetSheet As Worksheet
  
  Set GetBook = IIf(ThisBook, ThisWorkbook, ActiveWorkbook)
  If InStr(FindSheetCodeName, "!") > 0 Then _
    FindSheetCodeName = Replace(Mid(FindSheetCodeName, 2, InStr( _
      FindSheetCodeName, "!") - 2), "'", "") ' Имя листа должно быть БЕЗ "!"
  For Each GetSheet In GetBook.Worksheets
    If InStr(1, GetSheet.CodeName, FindSheetCodeName, vbTextCompare) _
    Or InStr(1, GetSheet.Name, FindSheetCodeName, vbTextCompare) Then _
      FindSheet = GetSheet.Index: Exit For
  Next GetSheet: Set GetSheet = Nothing: Set GetBook = Nothing
End Function
