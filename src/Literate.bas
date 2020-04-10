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

Private Gaps() As String, WordForm As New Collection

Function NumberFormatterRU(ByVal Number As Double, ByRef Item As WordFormType, _
  Optional ByVal InWords As Boolean = False) As String '> Число в слово
  Attribute NumberFormatterRU.VB_Description = "r312 ¦ Количество существительного (из списка)"
  '' "Один", "Два", "Шесть" Именительный падеж (есть что?)
  '' "Одного", "Двух", "Шести" Родительный падеж (нет чего?)
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
  Gaps = Split(Format(Number, IIf(InWords, "0,00", "0")), Chr(160))
  
  ' Item < UnitTyte.Min Xor Item >= WordForm.Count + UnitTyte.Min
  If Item < wtAsTonne Xor Item >= WordForm.Count + wtAsTonne Then ' HotFix!
    Debug.Print "Ошибка ввода: Не найден WordForm с ключом WordFormType#" & Item
    NumberFormatterRU = Number
  Else
    For Each aU In Gaps
      Teen = (aU Mod 100) >= 11 And (aU Mod 100) < 20 ' Им.падеж, Мн.число
      
      Select Case -(Not Teen) * aU Mod 10 ' Warn!
        Case Is = 1: aU = 1 ' "One"
        Case 2 To 4: aU = 2 ' "Few"
        Case Else: aU = 3 ' "Many"
      End Select
      
      If InWords Then
        Gaps(eZ) = NumeralRU(CInt(Gaps(eZ)), colCount, CByte(aU)) _
          & WordForm(CStr(colCount))(aU)
        eZ = eZ + 1: If colCount > wtInThousands Then _
          colCount = colCount - 1 Else colCount = Item ' HotFix!
      Else
        Gaps(eZ) = Gaps(eZ) & WordForm(CStr(Item))(aU)
      End If
    Next aU
    
    NumberFormatterRU = LTrim(Join(Gaps, ""))
  End If: Erase Gaps
End Function

Function NumeralRU(ByRef Digit As Integer, ByRef Item As WordFormType, _
  state As Byte) As String
  Attribute NumeralRU.VB_Description = "r312 ¦ Число прописью"
  Dim numeral As Variant, ending As Variant, secondDigit As Byte
  
  ending = [{"а", "", "е", "ь", "и"}]
  numeral = [{"", " один", " дв", " три", " четыр", " пят", " шест", " сем", " восем", " девят"}]
  numeral(1) = Item < wtInMills And Not Item = wtAsNone ' True = Женский род
  
  If Digit > 0 Then ' Если больше нуля
    If state = 1 Then
      If numeral(1) Then numeral(2) = " одна" ' Женский род
    ElseIf state = 2 Then
      Select Case (Digit Mod 10) ' Две последних цифры
        Case Is = 2 ' Мужской род
          If Not numeral(1) Then state = 0 ' state - 2
        Case Is = 3: state = 1 ' state - 1
      End Select
    ElseIf state = 3 Then
      Select Case (Digit Mod 100) ' Две последних цифры
        Case Is < 10: state = 3 ' state ' HotFix!
        Case Is = 12: state = 2 ' state - 1
        Case Is < 20: state = 1 ' state - 2
      End Select
    End If: secondDigit = (Digit Mod 100) \ 10: If (Digit Mod 10) > 0 Then _
      NumeralRU = numeral(Digit Mod 10 + 1) & ending(state + 1) ' Разряд #1
    
    Select Case (Digit Mod 100) ' Две последних цифры
      Case Is > 19
        Select Case secondDigit ' Разряд #2
          Case Is < 4
            NumeralRU = numeral(secondDigit + 1) _
              & IIf(secondDigit = 2, ending(1), "") & "дцать" & NumeralRU
          Case Is = 4
            NumeralRU = " сорок" & NumeralRU
          Case Else
            NumeralRU = numeral(secondDigit + 1) & "ьдесят" & NumeralRU
        End Select
      Case Is > 10
        NumeralRU = NumeralRU & "надцать"
      Case Is = 10
        NumeralRU = " десять"
    End Select
    
    If Digit > 99 Then ' Сотни
      Select Case (Digit \ 100) ' Разряд #3
        Case Is = 1
          NumeralRU = " стo" & NumeralRU
        Case 2 To 4
          NumeralRU = numeral(Digit \ 100 + 1) _
            & IIf(Digit \ 100 = 3, "", ending(3)) & "ст" _
            & IIf(Digit \ 100 = 2, ending(5), ending(1)) & NumeralRU
        Case Else
          NumeralRU = numeral(Digit \ 100 + 1) & "ьсот" & NumeralRU
      End Select
    End If
  ElseIf Item < wtInThousands Or Item > wtInQuadrills Then
    If UBound(Gaps) = 0 Then NumeralRU = " ноль"
  End If
End Function
