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

Function NumberFormatterRU(ByVal Number As Double, ByRef Item As WordFormType, _
  Optional ByVal InWords As Boolean = False) As String '> Число в слово
  Attribute NumberFormatterRU.VB_Description = "r312 ¦ Выбрать количество существительного из списка"
  '' "Один", "Два", "Шесть" Именительный падеж (есть что?)
  '' "Одного", "Двух", "Шести" Родительный падеж (нет чего?)
  '' МУЖСКОЙ и МУЖСКОЙ род (не отработаны)
  ' http://www.unicode.org/cldr/charts/29/supplemental/language_plural_rules.html
  ' Простое решение https://toster.ru/q/554384
  Dim Digits() As String, NN() As Variant, Numeral As New Collection
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
    Numeral.Add Number Mod 100 >= 11 And Number Mod 100 < 20, "-teen"
    NN = [{"", " одна", " две", " три", " четыр", " пят", " шест", " сем", " восем", " девят"}]
    'Numeral.Add NN, "-num"
    
    
    For Each aU In Digits: aU = CInt(aU)
      If InWords And aU > 0 Then ' Если число больше нуля
        'Digits(eZ) = Numeral("-num")(aU Mod 10) ' Единицы
        Digits(eZ) = NN(aU Mod 10 + 1) & "ь" ' Единицы
        
        If Len(aU) > 2 Then
          'Digits(eZ) = Numeral("-num")((aU Mod 100) \ 10) & "десят" & Digits(eZ)
          Digits(eZ) = NN((aU Mod 100) \ 10 + 1) & "десят" & Digits(eZ)
          If Len(aU) > 1 Then
            'Digits(eZ) = Numeral("-num")(aU \ 100) & "ста" & Digits(eZ)
            Digits(eZ) = NN(aU \ 100 + 1) & "сот" & Digits(eZ)
          End If
        End If
      End If
      
      Select Case -(Not Numeral("-teen")) * aU Mod 10 ' Warn!
        Case Is = 1
          Digits(eZ) = Digits(eZ) & WordForm(CStr(colCount))(1)
        Case 2 To 4
          Digits(eZ) = Digits(eZ) & WordForm(CStr(colCount))(2)
        Case Else
          Digits(eZ) = Digits(eZ) & WordForm(CStr(colCount))(3)
      End Select
      
      ' ВАЖНО! Доработать
      If colCount > wtAsNone And colCount < wtInQuadrills Then ' HotFix!
        If Digits(eZ) Like "000*" Then Digits(eZ) = Empty
        
        If colCount > wtInThousands Then _
          colCount = colCount - 1 Else colCount = Item
      Else
        If Digits(eZ) Like "000*" Then Digits(eZ) = Mid(Digits(eZ), 4)
      End If: eZ = eZ + 1
    Next aU
    
    NumberFormatterRU = LTrim(Join(Digits, ""))
  End If: Erase Digits: Set Numeral = Nothing
End Function
