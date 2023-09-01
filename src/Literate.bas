Attribute VB_Name = "Literate"
Option Explicit
Option Base 1
'123456789012345678901234567890123456h8nor@уа56789012345678901234567890123456789

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

Public Function NumberFormatterRU(ByVal number As Double, ByRef Ref_Item As WordFormType, _
  Optional ByVal isNumberText As Boolean = False) As String '> Число в слово
  Attribute NumberFormatterRU.VB_Description = "r317 ¦ Количество существительного (из списка)"
  Attribute NumberFormatterRU.VB_ProcData.VB_Invoke_Func = " \n7"
  '' "Один", "Два", "Шесть" Именительный падеж (есть что?)
  '' "Одного", "Двух", "Шести" Родительный падеж (нет чего?)
  Dim colCount As WordFormType, teenage As Boolean, aU As Variant, eZ As Byte
  
  If WordForm.Count = 0 Then
    WordForm.Add [{" тонна", " тонны", " тонн"}], CStr(wtAsTonne)
    WordForm.Add [{" неделя", " недели", " недель"}], CStr(wtAsWeek) ' Date
    WordForm.Add [{" минута", " минуты", " минут"}], CStr(wtAsMinute) ' Time
    WordForm.Add [{" секунда", " секунды", " секунд"}], CStr(wtAsSecond) ' Time
    WordForm.Add [{"", "", ""}], CStr(wtAsNone) ' HotFix!
    '         "One", "Few", "Many"; Для дробных «number» нужно применять "Few"
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
  
  colCount = (Len(CStr(number)) + 2) \ 3 - 1 ' Число разрядов если isNumberText
  If Not (isNumberText Xor colCount = 0) Then colCount = Ref_Item ' Warn!
  ' Разбить число на разряды по 3 цифры
  Gaps = Split(Format(number, IIf(isNumberText, "0,00", "0")), Chr(160))
  
  ' Item < UnitTyte.MinIndex Xor Item >= WordForm.Count + UnitTyte.MinIndex
  If Ref_Item < wtAsTonne Xor Ref_Item >= WordForm.Count + wtAsTonne Then ' HotFix!
    HookMsg "Ошибка ввода: Не найден WordForm с ключом WordFormType#" & Ref_Item, vbOKCancel
    NumberFormatterRU = number
  Else
    For Each aU In Gaps
      teenage = (aU Mod 100) >= 11 And (aU Mod 100) < 20 ' Им.падеж, Мн.число
      
      Select Case -(Not teenage) * aU Mod 10 ' Warn!
        Case Is = 1: aU = 1 ' "One"
        Case 2 To 4: aU = 2 ' "Few"
        Case Else: aU = 3 ' "Many"
      End Select
      
      If isNumberText Then
        Gaps(eZ) = NumeralRU(CInt(Gaps(eZ)), colCount, CByte(aU)) _
          & WordForm(CStr(colCount))(aU)
        eZ = eZ + 1: If colCount > wtInThousands Then _
          colCount = colCount - 1 Else colCount = Ref_Item ' HotFix!
      Else
        Gaps(eZ) = Gaps(eZ) & WordForm(CStr(Ref_Item))(aU)
      End If
    Next aU
    
    NumberFormatterRU = LTrim(Join(Gaps, ""))
  End If: Erase Gaps
End Function

Private Function NumeralRU(ByRef Ref_Digits As Integer, _
  ByRef Ref_Item As WordFormType, ByVal state As Byte) As String
  Attribute NumeralRU.VB_Description = "r314 ¦ Число прописью"
  Dim numeral As Variant, ending As Variant, secondDigit As Byte
  
  ending = [{"а", "", "е", "ь", "и"}]
  numeral = [{"", " один", " дв", " три", " четыр", " пят", " шест", " сем", " восем", " девят"}]
  numeral(1) = Ref_Item < wtInMills And Not Ref_Item = wtAsNone ' True = Женский род
  
  If Ref_Digits > 0 Then ' Если больше нуля
    Select Case state
      Case Is = 1
        If numeral(1) Then numeral(2) = " одна" ' Женский род
      Case Is = 2
        Select Case (Ref_Digits Mod 10) ' Одна последняя цифра
          Case Is = 2 ' Мужской род
            If Not numeral(1) Then state = 0 ' state - 2
          Case Is = 3: state = 1 ' state - 1
        End Select
      Case Is = 3
        Select Case (Ref_Digits Mod 100) ' Две последних цифры
          Case Is < 10: state = 3 ' state ' HotFix!
          Case Is = 12: state = 2 ' state - 1
          Case Is < 20: state = 1 ' state - 2
        End Select
    End Select: If (Ref_Digits Mod 10) > 0 Then _
      NumeralRU = numeral(Ref_Digits Mod 10 + 1) & ending(state + 1) ' Разряд #1
    
    Select Case (Ref_Digits Mod 100) ' Две последних цифры
      Case Is > 19: secondDigit = (Ref_Digits Mod 100) \ 10
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
    
    If Ref_Digits > 99 Then ' Сотни
      Select Case (Ref_Digits \ 100) ' Разряд #3
        Case Is = 1
          NumeralRU = " стo" & NumeralRU
        Case 2 To 4
          NumeralRU = numeral(Ref_Digits \ 100 + 1) _
            & IIf(Ref_Digits \ 100 = 3, "", ending(3)) & "ст" _
            & IIf(Ref_Digits \ 100 = 2, ending(5), ending(1)) & NumeralRU
        Case Else
          NumeralRU = numeral(Ref_Digits \ 100 + 1) & "ьсот" & NumeralRU
      End Select
    End If
  ElseIf Ref_Item < wtInThousands Or Ref_Item > wtInQuadrills Then
    If UBound(Gaps) = 0 Then NumeralRU = " ноль"
  End If
End Function

Public Function PorterStemmerRU(ByVal word As String) As String
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

Private Function RemoveEndings(ByRef Ref_Word As String, _
  ByVal regex As Variant, ByVal region As Byte) As Boolean ' Удалить окончание (самое длинное)
  Attribute RemoveEndings.VB_Description = "r314 ¦ Стеммер: удаление окончания"
  Dim rZ As Byte, prefix As String, regMatch As Variant
  
  prefix = Mid(Ref_Word, 1, IIf(region, region, 1) - 1) ' prefix <- region
  Ref_Word = Mid(Ref_Word, Len(prefix) + 1)
  If IsArray(regex) Then
    For Each regMatch In Split(regex(0))
      If Ref_Word Like "*[яа]" & regMatch Then ' Если найден аффикс
        Ref_Word = Left(Ref_Word, Len(Ref_Word) - Len(regMatch))      
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
          If Ref_Word Like "*" & Mid(regMatch, region, 1) & Mid(regMatch, rZ) Then
            regMatch = Mid(regMatch, region, 1) & Mid(regMatch, rZ)
            Ref_Word = Left(Ref_Word, Len(Ref_Word) - Len(regMatch))
            RemoveEndings = True: Exit For
          End If: If region = 1 Then Exit For
        Next region: If RemoveEndings Then Exit For
      On Error GoTo 0
    Next regMatch
  End If: Ref_Word = prefix & Ref_Word
End Function

Private Function FindRegions(ByRef Ref_rV As Byte, ByVal word As String) As Byte
  Attribute FindRegions.VB_Description = "r317 ¦ Стеммер: регион r2"
  Dim prevChar As String, char As String, cnt As Byte, state As Byte
  
  If isVowel(Left(word, 1)) Then Ref_rV = 2: state = 1 ' После первой гласной
  For cnt = 2 To Len(word)
    prevChar = Mid(word, cnt - 1, 1): char = Mid(word, cnt, 1)
    Select Case state
      Case Is = 0: If isVowel(char) Then Ref_rV = cnt + 1: state = 1
      Case Is = 1: If Not isVowel(char) And isVowel(prevChar) Then state = 2
      Case Is = 2: If Not isVowel(char) And isVowel(prevChar) Then _
        FindRegions = cnt + 1: Exit For
    End Select
  Next i
End Function

Private Function isVowel(ByVal char As String) As Boolean
  Attribute isVowel.VB_Description = "r314 ¦ Гласная буква"
  Const VOWEL As String = "[ёаеиоуыэюя]"
  
  isVowel = InStr(Mid(VOWEL, 2, Len(VOWEL) - 1), char)
End Function
