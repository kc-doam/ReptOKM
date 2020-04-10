Attribute VB_Name = "Execute"
Option Explicit
Option Base 1
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890

Public Sub Shell_Sort(ByRef Items As Variant, ByVal GetColumn As Byte)
  Attribute Shell_Sort.VB_Description = "r311 ¦ Сортировка методом Шелла"
  ' Сортировка методом Шелла Дональда с интервалами длин ЭМ Марцина Циура
  ' http://ru.wikibooks.org/wiki/Примеры_реализации_сортировки_Шелла#VBA
  Dim cnt As Integer, gap As Integer, pos As Integer, inv_a108870 As Variant
  Dim kZ As Integer, mZ As Integer, nZ As Integer
  
  ReDim tmp(UBound(Items, 1)) As Variant
  mZ = LBound(Items, 2) ' Нижняя граница массива
  nZ = UBound(Items, 2)
  If nZ > mZ Then ' Если в массиве одна запись, То нет смысла сортировать
    ' Инвертированная эмпирическая последовательность Марцина Циура
    inv_a108870 = Array(776591, 345152, 153401, 68178, 30301, 13467, _
      5985, 2660, 1182, 525, 233, 103, 46, 20, 9, 4, 1)
    kZ = nZ - mZ + 1 ' Количество записей в массиве
    nZ = LBound(inv_a108870) - 1 ' Счётчик последовательности Марцина Циура
    Do
      nZ = nZ + 1
      gap = inv_a108870(nZ) ' Наименьший интервал
    Loop Until kZ > gap ' Цикл ДО
    Do
      gap = inv_a108870(nZ) ' Интервал
      For cnt = (gap + mZ) To UBound(Items, 2)
        pos = cnt
        For kZ = LBound(Items, 1) To UBound(Items, 1)
          tmp(kZ) = Items(kZ, cnt)
        Next kZ
        Do While Items(GetColumn, pos - gap) > tmp(GetColumn) ' Выполнять ПОКА
          For kZ = LBound(Items, 1) To UBound(Items, 1)
            Items(kZ, pos) = Items(kZ, pos - gap)
            Items(kZ, pos - gap) = tmp(kZ)
          Next kZ
          pos = pos - gap
          If (pos - gap) < mZ Then Exit Do
        Loop
      Next cnt
      nZ = nZ + 1
    Loop Until gap = 1 ' Цикл ДО
  Else
  '  Debug.Print "Shell_Sort: В массиве ОДНА или НУЛЬ записей"
  End If
End Sub

Function GetQuarterNumber(ByVal SetDate As Date, Optional ByVal GetYear _
  As Boolean = False) As String ' Номер квартала
  Attribute GetQuarterNumber.VB_Description = "r310 ¦ Заменить цифру на номер квартала"
  GetQuarterNumber = (Month(SetDate) - 1) \ 3 + 1
  Select Case CByte(GetQuarterNumber)
    Case 1 To 3: GetQuarterNumber = String(GetQuarterNumber, "I") & " квартал "
    Case 4: GetQuarterNumber = "IV квартал "
  End Select
  If GetYear Then GetQuarterNumber = GetQuarterNumber & Year(SetDate) & " г."
End Function

Function Trip(ByVal Text As String) As String
  Attribute Trip.VB_Description = "r311 ¦ Убрать перед/после текста (неразрывные) пробелы и переносы строк"
  Const LF0 As String = "[" & vbLf & " ]*", LF8 As String = "*[" & vbLf & " ]"
  
  While Text Like LF0 Or Text Like LF8
    If Text Like LF0 Then Text = Trim(Right(Text, Len(Text) - 1))
    If Text Like LF8 Then Text = Trim(Left(Text, Len(Text) - 1))
  Wend: Trip = Text
End Function

Function Tripp(ByVal Item_notRange As Variant) As Variant
  Attribute Tripp.VB_Description = "r311 ¦ Функция удаления разрывов строк по краям"
  Const CRL0 As String = "[" & vbCrLf & " ]*", LF0 As String = "[" & vbLf & " ]*"
  Const CRL8 As String = "*[" & vbCrLf & " ]", LF8 As String = "*[" & vbLf & " ]"
  Dim str As String, eZ As Byte, kZ As Integer
  
  For kZ = LBound(Item_notRange, 2) To UBound(Item_notRange, 2)
    For eZ = LBound(Item_notRange, 1) To UBound(Item_notRange, 1)
      str = Trim(Replace(Item_notRange(eZ, kZ), Chr(160), " ")) ' Неразрывный пробел
      While str Like CRL0 Or str Like LF0 Or str Like CRL8 Or str Like LF8
        If str Like CRL0 Or str Like LF0 Then str = Trim(Right(str, Len(str) - 1))
        If str Like CRL8 Or str Like LF8 Then str = Trim(Left(str, Len(str) - 1))
      Wend: Item_notRange(eZ, kZ) = str
  Next eZ, kZ: Tripp = Item_notRange
End Function


Function ClearSpacesInText(ByVal Text As String) As String
  Attribute ClearSpacesInText.VB_Description = "r270 ¦ Удалить опечатки и лишние пробелы"
  Text = Replace(Text, Chr(160), " ")       ' Убрать: неразрывный пробел
  ''text = Replace(text, ".", ". ") ' Ошибка при проверке "т.ч."
  Text = Replace(Text, "т.ч.", "т ч")
  Text = Replace(Text, """", "")            ' стереть кавычки
  ' МОЖЕТ ЛИШНЕЕ... СМ ' HotFix для ключа #2
  Text = Replace(Text, " - ", " ")          ' стереть "не-до-тире"
  Text = Replace(Text, " -", " ")           ' стереть "не-до-тире2"
  
  Text = Replace(Text, "-", " ")            ' стереть "не-до-тире3"
  
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
    Text = Replace(Text, "  ", " ")
  Loop
  ''text = Replace(text, Chr(39), Chr(34))   ' апостроф
  ''text = Replace(text, Chr(171), Chr(34))  ' левые двойные кавычки
  ''text = Replace(text, Chr(187), Chr(34))  ' правые двойные кавычки
  ''text = Replace(text, Chr(147), Chr(34))  ' левые четверть круга кавычки
  ''text = Replace(text, Chr(148), Chr(34))  ' правые четверть круга кавычки
  ''text = Replace(text, Chr(34) & Chr(34), Chr(34))
  ClearSpacesInText = Trim(Text)
End Function
