Attribute VB_Name = "Execute"
Option Explicit
Option Base 1
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890

Public Sub Shell_Sort(ByRef items As Variant, ByVal column As Byte)
  Attribute Shell_Sort.VB_Description = "r310 Сортировка методом Шелла"
  ' Сортировка методом Шелла Дональда с интервалами длин ЭМ Марцина Циура
  ' http://ru.wikibooks.org/wiki/Примеры_реализации_сортировки_Шелла#VBA
  Dim cnt As Integer, gap As Integer, pos As Integer, inv_a108870 As Variant
  Dim k As Integer, m As Integer, n As Integer
  
  ReDim tmp(UBound(items, 1)) As Variant
  m = LBound(items, 2) ' Нижняя граница массива
  n = UBound(items, 2)
  If n > m Then ' Если в массиве одна запись, То нет смысла сортировать
    ' Инвертированная эмпирическая последовательность Марцина Циура
    inv_a108870 = Array(776591, 345152, 153401, 68178, 30301, 13467, _
      5985, 2660, 1182, 525, 233, 103, 46, 20, 9, 4, 1)
    k = n - m + 1 ' Количество записей в массиве
    n = LBound(inv_a108870) - 1 ' Счётчик последовательности Марцина Циура
    Do
      n = n + 1
      gap = inv_a108870(n) ' Наименьший интервал
    Loop Until k > gap ' Цикл ДО
    Do
      gap = inv_a108870(n) ' Интервал
      For cnt = (gap + m) To UBound(items, 2)
        pos = cnt
        For k = LBound(items, 1) To UBound(items, 1)
          tmp(k) = items(k, cnt)
        Next k
        Do While items(column, pos - gap) > tmp(column) ' Выполнять ПОКА
          For k = LBound(items, 1) To UBound(items, 1)
            items(k, pos) = items(k, pos - gap)
            items(k, pos - gap) = tmp(k)
          Next k
          pos = pos - gap
          If (pos - gap) < m Then Exit Do
        Loop
      Next cnt
      n = n + 1
    Loop Until gap = 1 ' Цикл ДО
  Else
  '  Debug.Print "Shell_Sort: В массиве ОДНА или НУЛЬ записей"
  End If
End Sub

Function GetQuarterNumber(ByVal SetDate As Date, Optional ByVal GetYear _
  As Boolean = False) As String ' Номер квартала
  Attribute GetQuarterNumber.VB_Description = "r310 Заменить цифру на номер квартала"
  GetQuarterNumber = (Month(SetDate) - 1) \ 3 + 1
  Select Case CByte(GetQuarterNumber)
    Case 1 To 3: GetQuarterNumber = String(GetQuarterNumber, "I") & " квартал "
    Case 4: GetQuarterNumber = "IV квартал "
  End Select
  If GetYear Then GetQuarterNumber = GetQuarterNumber & Year(SetDate) & " г."
End Function

Function Trip(ByVal str As String) As String
  Attribute Trip.VB_Description = "r304 Убрать перед/после текста (неразрывные) пробелы и переносы строк"
  Const LF0 As String = "[" & vbLf & " ]*", LF8 As String = "*[" & vbLf & " ]"
  
  While str Like LF0 Or str Like LF8
    If str Like LF0 Then str = Trim(Right(str, Len(str) - 1))
    If str Like LF8 Then str = Trim(Left(str, Len(str) - 1))
  Wend: Trip = str
End Function

Function Tripp(ByVal item_notRange As Variant) As Variant
  Attribute Tripp.VB_Description = "r310 Функция удаления разрывов строк по краям"
  Const CRL0 As String = "[" & vbCrLf & " ]*", LF0 As String = "[" & vbLf & " ]*"
  Const CRL8 As String = "*[" & vbCrLf & " ]", LF8 As String = "*[" & vbLf & " ]"
  Dim str As String, i As Integer, j As Byte
  
  For i = LBound(item_notRange, 2) To UBound(item_notRange, 2)
    For j = LBound(item_notRange, 1) To UBound(item_notRange, 1)
      str = Trim(Replace(item_notRange(j, i), Chr(160), " ")) ' Неразрывный пробел
      While str Like CRL0 Or str Like LF0 Or str Like CRL8 Or str Like LF8
        If str Like CRL0 Or str Like LF0 Then str = Trim(Right(str, Len(str) - 1))
        If str Like CRL8 Or str Like LF8 Then str = Trim(Left(str, Len(str) - 1))
      Wend: item_notRange(j, i) = str
  Next j, i: Tripp = item_notRange
End Function


Function ClearSpacesInText(ByVal text As String) As String
  Attribute ClearSpacesInText.VB_Description = "r270 Удалить опечатки и лишние пробелы"
  text = Replace(text, Chr(160), " ")       ' Убрать: неразрывный пробел
  ''text = Replace(text, ".", ". ") ' Ошибка при проверке "т.ч."
  text = Replace(text, "т.ч.", "т ч")
  text = Replace(text, """", "")            ' стереть кавычки
  ' МОЖЕТ ЛИШНЕЕ... СМ ' HotFix для ключа #2
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
