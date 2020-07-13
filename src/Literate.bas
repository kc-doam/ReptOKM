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
  Attribute NumberFormatterRU.VB_Description = "r313 ¦ Количество существительного (из списка)"
  Attribute NumberFormatterRU.VB_ProcData.VB_Invoke_Func = " \n7"
  '' "Один", "Два", "Шесть" Именительный падеж (есть что?)
  '' "Одного", "Двух", "Шести" Родительный падеж (нет чего?)
  Dim colCount As WordFormType, Teen As Boolean, aU As Variant, eZ As Byte
  
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
    WordForm.Add [{" Вопрос", " Вопроса", " Вопросов"}], CStr(wtAsQuestion)
    WordForm.Add [{" Материал", " Материала", " Материалов"}], CStr(wtAsMaterial)
  End If
  
  colCount = (Len(CStr(Number)) + 2) \ 3 - 1 ' Число разрядов если InWords
  If Not (InWords Xor colCount = 0) Then colCount = Item ' Warn!
  ' Разбить число на разряды по 3 цифры
  Gaps = Split(Format(Number, IIf(InWords, "0,00", "0")), Chr(160))
  
  ' Item < UnitTyte.MinIndex Xor Item >= WordForm.Count + UnitTyte.MinIndex
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

Private Function NumeralRU(ByRef Digit As Integer, ByRef Item As WordFormType, _
  state As Byte) As String
  Attribute NumeralRU.VB_Description = "r313 ¦ Число прописью"
  Dim numeral As Variant, ending As Variant, secondDigit As Byte
  
  ending = [{"а", "", "е", "ь", "и"}]
  numeral = [{"", " один", " дв", " три", " четыр", " пят", " шест", " сем", " восем", " девят"}]
  numeral(1) = Item < wtInMills And Not Item = wtAsNone ' True = Женский род
  
  If Digit > 0 Then ' Если больше нуля
    Select Case state
      Case Is = 1
        If numeral(1) Then numeral(2) = " одна" ' Женский род
      Case Is = 2
        Select Case (Digit Mod 10) ' Одна последняя цифра
          Case Is = 2 ' Мужской род
            If Not numeral(1) Then state = 0 ' state - 2
          Case Is = 3: state = 1 ' state - 1
        End Select
      Case Is = 3
        Select Case (Digit Mod 100) ' Две последних цифры
          Case Is < 10: state = 3 ' state ' HotFix!
          Case Is = 12: state = 2 ' state - 1
          Case Is < 20: state = 1 ' state - 2
        End Select
    End Select: If (Digit Mod 10) > 0 Then _
      NumeralRU = numeral(Digit Mod 10 + 1) & ending(state + 1) ' Разряд #1
    
    Select Case (Digit Mod 100) ' Две последних цифры
      Case Is > 19: secondDigit = (Digit Mod 100) \ 10
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

Function PorterStemmerRU(ByVal word As String) As String
  Attribute PorterStemmerRU.VB_Description = "r313 ¦ Uni test failed: Стеммер Мартина Портера для русского языка"
  ' Переписано с http://snowball.tartarus.org/algorithms/russian/stemmer.html
  
  '2 Причастие совершенного вида
  Const PERFECTiveGERUNDs As String = "[ыи]вшись [ыи]вши [ыи]в"
  '1 Возвратное
  Const REFLEXive As String = "ся сь"
  '2 Причастие
  Const PARTICIPLEs As String = "ующ [ыи]вш"
  '1 Прилагательное
  Const ADJECTive As String = "[ое]му [ыи]ми [ое]го " _
    & "[яа]я [юуое]ю [ыи]х [ыоие]м [ыоие]й [ыоие]е"
  '2 Глагол
  Const VERBs As String = "[уе]йте уют ишь [ыи]ть ите ены ено ена " _
    & "[ыи]ло [ыи]ли [ыи]ла ует [яыи]т [ыи]м [ыи]л [уе]й ую ен ю"
  '1 Имя существительное
  Const NOUN As String = "иями иях иям [яа]ми ием ией [ьи]я [ьи]ю " _
    & "[яа]х [яоеа]м [оие]й [ие]и [ьи]е [ое]в я ю ы ь у о й и е а"
  '1 Прилагательное привосходной степени
  Const SUPERLATive As String = "ейше ейш"
  '1 Словообразующее окончание в r2
  Const DERIVATIONAL As String = "ость ост"
  Dim rV As Byte, r2 As Byte ' r1 As Byte
  
  If Len(word) > 0 Then word = Replace(LCase(word), "*", "") Else Exit Function
  ' rV - начало области слова после первой гласной (Если гласных нет = 0)
  ' r1 - начало области слова "Гласная-Согласная" с начала слова
  r2 = FindRegions(rV, word) ' r2 - начало области "Гласная-Согласная" после r1
  ' [Шаг 1] Если существует окончание PERFECTIVE GERUND – удалить и завершить
  If Not RemoveEndings(word, Array(Replace(PERFECTiveGERUNDs, "[ыи]", ""), _
    PERFECTiveGERUNDs), rV) Then
    ' Если существует окончание REFLEXIVE – удалить
    RemoveEndings word, REFLEXive, rV
    ' Удалить одно из окончаний и завершить: PARTICIPLE + ADJECTIVE, VERB, NOUN
    If RemoveEndings(word, ADJECTive, rV) Then
      RemoveEndings word, Array("ющ вш нн ем щ", PARTICIPLEs), rV
    Else
      If Not RemoveEndings(word, Array("ешь йте ете нно ны ть " _
        & "ют ет но ло ем ли на ла н л й", VERBs), rV) Then _
        RemoveEndings word, NOUN, rV ' не УЮТ и не МЛЕЮТ, но БЕСЕДуют
    End If
  End If
  ' [Шаг 2] Если слово оканчивается на "и" - удалить
  RemoveEndings word, "и", rV
  ' [ШАГ 3] Если существует окончание DERIVATIONAL в r2 - удалить
  RemoveEndings word, DERIVATIONAL, r2
  ' [ШАГ 4] Удалить одно из окончаний слова: (Н)Н + SUPERLATIVE, (Н)Н, Ь
  RemoveEndings word, SUPERLATive, rV
  If RemoveEndings(word, "нн", rV) Then word = word & "н"
  RemoveEndings word, "ь", rV
  
  PorterStemmerRU = word
End Function

Private Function RemoveEndings(ByRef word As String, ByVal regex As Variant, _
  ByVal region As Byte) As Boolean ' Удалить окончание (самое длинное)
  Attribute PorterStemmerRU.VB_Description = "r313 ¦ Стеммер: удаление окончания"
  Dim rZ As Byte, prefix As String, regMatch As Variant
  
  prefix = Mid(word, 1, IIf(region, region, 1) - 1) ' prefix <- region
  word = Mid(word, Len(prefix) + 1)
  If IsArray(regex) Then
    For Each regMatch In Split(regex(0))
      If word Like "*[яа]" & regMatch Then ' Если найден аффикс
        word = Left(word, Len(word) - Len(regMatch))      
        RemoveEndings = True: Exit For      
      End If
    Next regMatch: regex = regex(1)
  End If
  If Not RemoveEndings Then
    For Each regMatch In Split(regex)
      rZ = InStr(regMatch, "]") + 1 ' rZ - начало области после [list]
      On Error Resume Next
        For region = 2 To rZ - 2
          If rZ < 2 Then region = 1: rZ = 2 ' Если нет [list]
          If word Like "*" & Mid(regMatch, region, 1) & Mid(regMatch, rZ) Then
            regMatch = Mid(regMatch, region, 1) & Mid(regMatch, rZ)
            word = Left(word, Len(word) - Len(regMatch))
            RemoveEndings = True: Exit For
          End If: If region = 1 Then Exit For
        Next region: If RemoveEndings Then Exit For
      On Error GoTo 0
    Next regMatch
  End If: word = prefix & word
End Function

Private Function FindRegions(ByRef rV As Byte, ByVal word As String) As Byte
  Attribute PorterStemmerRU.VB_Description = "r313 ¦ Стеммер: регион r2"
  Dim prevChar As String, Char As String, cnt As Byte, state As Byte
  
  If isVowel(Left(word, 1)) Then rV = 2: state = 1 ' После первой гласной
  For cnt = 2 To Len(word)
    prevChar = Mid(word, cnt - 1, 1): Char = Mid(word, cnt, 1)
    Select Case state
      Case 0: If isVowel(Char) Then rV = cnt + 1: state = 1
      Case 1: If Not isVowel(Char) And isVowel(prevChar) Then state = 2
      Case 2: If Not isVowel(Char) And isVowel(prevChar) Then _
        FindRegions = cnt + 1: Exit For
    End Select
  Next i
End Function

Private Function isVowel(ByVal Char As String) As Boolean
  Attribute PorterStemmerRU.VB_Description = "r313 ¦ Гласная буква"
  Const VOWEL As String = "[ёаеиоуыэюя]"
  
  isVowel = InStr(Mid(VOWEL, 2, Len(VOWEL) - 1), Char)
End Function
