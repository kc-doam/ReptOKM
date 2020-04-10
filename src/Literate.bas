Attribute VB_Name = "Literate"
Option Explicit
Option Base 1
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890
Public Enum WordFormType
  wtAsTonne = -4
  wtAsWeek = -3
  wtAsMinute = -2
  wtAsSecond = -1
  wtAsNone = 0
  wtInThousands = 1
  wtInMills = 2 ' Муж и Ср.род > 1
  wtInBills = 3
  wtInTrills = 4
  wtInQuadrills = 5
  wtAsHour = 6
  wtAsDay = 7
  wtAsYear = 8
  wtAsMetre = 9
  wtAsGram = 10
  wtAsRuble = 11
  wtAsQuestion = 12
  wtAsMaterial = 13
End Enum

Public Const ZERO As String = " ноль"

Private Digits() As String, WordForm As New Collection

Function NumberFormatterRU(ByVal Number As Double, ByRef Item As WordFormType, _
  Optional ByVal InWords As Boolean = False) As String '> Число в слово
  Attribute NumberFormatterRU.VB_Description = "r312 ¦ Выбрать количество существительного из списка"
  '' "Один", "Два", "Шесть" Именительный падеж (есть что?)
  '' "Одного", "Двух", "Шести" Родительный падеж (нет чего?)
  '' МУЖСКОЙ и МУЖСКОЙ род (не отработаны)
  Dim Digits_test() As String, Teen As Boolean
  Dim colCount As WordFormType, aU As Variant, eZ As Byte
  
  If WordForm.Count = 0 Then
    WordForm.Add [{" тонна", " тонны", " тонн"}], CStr(wtAsTonne)
    WordForm.Add [{" неделя", " недели", " недель"}], CStr(wtAsWeek) ' Date
    WordForm.Add [{" минута", " минуты", " минут"}], CStr(wtAsMinute) ' Time
    WordForm.Add [{" секунда", " секунды", " секунд"}], CStr(wtAsSecond) ' Time
    WordForm.Add [{"", "", ""}], CStr(wtAsNone) ' HotFix!
    '         "One", "Few", "Many"; Для дробных «Number» нужно применять "Few"
    WordForm.Add [{" тысяча", " тысячи", " тысяч"}], CStr(wtInThousands)
    WordForm.Add [{" миллион", " миллиона", " миллионов"}], CStr(wtInMills)
    WordForm.Add [{" миллиард", " миллиарда", " миллиардов"}], CStr(wtInBills)
    WordForm.Add [{" триллион", " триллиона", " триллионов"}], CStr(wtInTrills)
    WordForm.Add [{" квадриллион", " квадриллиона", " квадриллионов"}], _
      CStr(wtInQuadrills)
    WordForm.Add [{" час", " часа", " часов"}], CStr(wtAsHour) ' Time
    WordForm.Add [{" день", " дня", " дней"}], CStr(wtAsDay) ' Date
    WordForm.Add [{" год", " года", " лет"}], CStr(wtAsYear) ' Date
    WordForm.Add [{" метр", " метра", " метров"}], CStr(wtAsMetre)
    WordForm.Add [{" грам", " грама", " грамов"}], CStr(wtAsGram)
    WordForm.Add [{" рубль", " рубля", " рублей"}], CStr(wtAsRuble)
    WordForm.Add [{" вопрос", " вопроса", " вопросов"}], CStr(wtAsQuestion)
    WordForm.Add [{" материал", " материала", " материалов"}], CStr(wtAsMaterial)
  End If
  
  colCount = (Len(CStr(Number)) + 2) \ 3 - 1 ' Число разрядов если InWords
  If Not (InWords Xor colCount = 0) Then colCount = Item ' Warn!
  ' Разбить число на разряды по 3 цифры
  Digits = Split(Format(Number, IIf(InWords, "0,00", "0")), Chr(160))
  
  ' Item < UnitTyte.Min Xor Item >= WordForm.Count + UnitTyte.Min
  If Item < wtAsTonne Xor Item >= WordForm.Count + wtAsTonne Then ' HotFix!
    Debug.Print "Ошибка ввода: Не найден WordForm с ключом WordFormType#" & Item
    NumberFormatterRU = Number
  Else
    For Each aU In Digits
      Teen = (aU Mod 100 >= 11 And aU Mod 100 < 20)
      
      Select Case -(Not Teen) * aU Mod 10 ' Warn!
        Case Is = 1: aU = 1 ' "One"
        Case 2 To 4: aU = 2 ' "Few"
        Case Else: aU = 3 ' "Many"
      End Select
      
      aU = NumeralRU(CInt(Digits(eZ)), colCount, CByte(aU))
      
      If colCount > wtAsNone And colCount < wtInQuadrills Then ' HotFix!
        If Digits(eZ) Like "000*" Then
          Digits(eZ) = Empty
        Else
          Digits(eZ) = aU
        End If
        
        If colCount > wtInThousands Then _
          colCount = colCount - 1 Else colCount = Item
      Else
        If Digits(eZ) Like "000*" Then
          If aU = "" Then ' " -" ZERO
            Digits(eZ) = IIf(UBound(Digits) > 0, "", ZERO) _
              & WordForm(CStr(colCount))(3)
          End If
        Else
          Digits(eZ) = aU
        End If
      End If: eZ = eZ + 1
    Next aU
    
    NumberFormatterRU = LTrim(Join(Digits, ""))
  End If: Erase Digits
End Function

Function NumeralRU(ByRef Digit As Integer, ByRef Item As WordFormType, _
  state As Byte) As String
  Attribute NumeralRU.VB_Description = "r312 ¦ "
  Dim numeral As Variant, prefix As String, postfix As String
  
  numeral = [{"", " один", " два", " три", " четыр", " пят", " шест", " сем", " восем", " девят"}]
  
  If (Digit \ 10) > 0 Then
    If (Digit \ 10) < 10 Or (Digit Mod 10) = 0 Then
      Select Case (Digit \ 10)
        Case Is < 4
          prefix = numeral((Digit Mod 100) \ 10 + 1) & "дцать"
        Case Is = 4
          prefix = " сорок"
        Case Else
          prefix = numeral((Digit Mod 100) \ 10 + 1) & "ьдесят"
      End Select
    End If
  End If
  
  If Digit >= 10 And Digit <= 19 Then ' -Teen
    numeral(3) = " две" '': If Digit > 14 Then postfix = "ь"
    
    If Digit = 10 Then
      NumeralRU = " десять" & WordForm(CStr(Item))(state)
    Else
      NumeralRU = numeral(Digit Mod 10 + 1) & postfix & "надцать" _
        & WordForm(CStr(Item))(state)
    End If
    
    Exit Function
  End If
  
  If Not (Digit Mod 10) = 0 Then
    ' Существительное: Мужской род
    If (Item > wtInThousands Or Item = wtAsNone) Then
      Select Case state
        Case Is = 2
          If (Digit Mod 10) = 4 Then postfix = "е"
        Case Is = 3
          postfix = "ь"
      End Select
    Else ' Существительное: Женский род
      numeral(2) = " одна": numeral(3) = " две"
      Select Case state
        Case Is = 2
          If (Digit Mod 10) = 4 Then postfix = "е"
        Case Is = 3
          postfix = "ь"
      End Select
    End If
  End If
  
  If Digit > 0 Then ' Если число больше нуля
    NumeralRU = numeral(Digit Mod 10 + 1) & postfix ' Единицы
    
    If Len(Digit) > 1 Then
      NumeralRU = prefix & NumeralRU
      If Len(Digit) > 2 Then
        NumeralRU = numeral(Digit \ 100 + 1) & "сот" & NumeralRU
      End If
    End If
    
    NumeralRU = NumeralRU & WordForm(CStr(Item))(state)
  End If
End Function
