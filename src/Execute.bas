Attribute VB_Name = "Execute"
Option Explicit
Option Base 1
'123456789012345678901234567890123456h8nor@уа56789012345678901234567890123456789

Public Sub Shell_Sort(ByRef Ref_Items As Variant, ByVal getColumn As Byte)
  Attribute Shell_Sort.VB_Description = "r314 ¦ Сортировка методом Шелла"
  ' Сортировка методом Шелла Дональда с интервалами длин ЭМ Марцина Циура
  ' http://ru.wikibooks.org/wiki/Примеры_реализации_сортировки_Шелла#VBA
  Dim cnt As Integer, gap As Integer, pos As Integer, inv_a108870 As Variant
  Dim kZ As Integer, mZ As Integer, nZ As Integer
  
  ReDim tmp(UBound(Ref_Items, 1)) As Variant
  mZ = LBound(Ref_Items, 2) ' Нижняя граница массива
  nZ = UBound(Ref_Items, 2)
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
      For cnt = (gap + mZ) To UBound(Ref_Items, 2)
        pos = cnt
        For kZ = LBound(Ref_Items, 1) To UBound(Ref_Items, 1)
          tmp(kZ) = Ref_Items(kZ, cnt)
        Next kZ
        Do While Ref_Items(getColumn, pos - gap) > tmp(getColumn) ' Выполнять ПОКА
          For kZ = LBound(Ref_Items, 1) To UBound(Ref_Items, 1)
            Ref_Items(kZ, pos) = Ref_Items(kZ, pos - gap)
            Ref_Items(kZ, pos - gap) = tmp(kZ)
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

Function GetQuarterNumber(ByVal getDate As Date, _
  Optional ByVal getYear As Boolean = False) As String ' Номер квартала
  Attribute GetQuarterNumber.VB_Description = "r314 ¦ Заменить цифру на номер квартала"
  Attribute GetQuarterNumber.VB_ProcData.VB_Invoke_Func = " \n2"
  GetQuarterNumber = (Month(getDate) - 1) \ 3 + 1
  Select Case CByte(GetQuarterNumber)
    Case 1 To 3: GetQuarterNumber = String(GetQuarterNumber, "I") & " квартал "
    Case Is = 4: GetQuarterNumber = "IV квартал "
  End Select
  If getYear Then GetQuarterNumber = GetQuarterNumber & Year(getDate) & " г."
End Function

Function Trip(ByVal text As String) As String
  Attribute Trip.VB_Description = "r314 ¦ Убрать перед/после текста (неразрывные) пробелы и переносы строк"
  Const LF0 As String = "[" & vbLf & " ]*", LF8 As String = "*[" & vbLf & " ]"
  
  While text Like LF0 Or text Like LF8
    If text Like LF0 Then text = Trim(Right(text, Len(text) - 1))
    If text Like LF8 Then text = Trim(Left(text, Len(text) - 1))
  Wend: Trip = text
End Function

Function Tripp(ByVal item_notRange As Variant) As Variant
  Attribute Tripp.VB_Description = "r314 ¦ Функция удаления разрывов строк по краям"
  Const NL0 As String = "[" & vbCrLf & " ]*", LF0 As String = "[" & vbLf & " ]*"
  Const NL8 As String = "*[" & vbCrLf & " ]", LF8 As String = "*[" & vbLf & " ]"
  Dim str As String, eZ As Byte, kZ As Integer
  
  For kZ = LBound(item_notRange, 2) To UBound(item_notRange, 2)
    For eZ = LBound(item_notRange, 1) To UBound(item_notRange, 1)
      str = Trim(Replace(item_notRange(eZ, kZ), Chr(160), " ")) ' Неразрывный пробел
      While str Like NL0 Or str Like LF0 Or str Like NL8 Or str Like LF8
        If str Like NL0 Or str Like LF0 Then str = Trim(Right(str, Len(str) - 1))
        If str Like NL8 Or str Like LF8 Then str = Trim(Left(str, Len(str) - 1))
      Wend: item_notRange(eZ, kZ) = str
  Next eZ, kZ: Tripp = item_notRange
End Function


Function ClearSpacesInText(ByVal text As String) As String
  Attribute ClearSpacesInText.VB_Description = "r314 ¦ Удалить опечатки и лишние пробелы"
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
