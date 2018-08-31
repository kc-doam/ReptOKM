Attribute VB_Name = "Main"
Option Explicit
Option Base 1
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890
Public Const REV As Integer = &H12C ' ver.2 rev.300

' Коллекция с именами папок статистик
Public DirName As New Collection
' Коллекция файлов статистик менеджеров (индекс совпадает с DirName)
Public FileName As New Collection
' Коллекция с именами менеджеров (индекс совпадает с DirName)
Public Manager As New Collection
' Коллекции: ключи коллекции BankID, рабочие листы и колонки
Public BankID As New Collection
' Коллекция: реквизиты и контакты контрагентов, цены
Public BankSUPP As New Collection


Private Sub Auto_Open()
  Attribute SettingsBankID.VB_Description = "r302 Автозапуск"
  ' Папки для поиска статистик менеджеров
  ' ПРИМЕЧАНИЕ: Путь необходимо указывать с косой чертой в конце строки,
  '             а также добавить маску в процедуру GetWorkbooks
  Const Dir_ss As String = "#Finansist\YCHET\"
  Const Dir_az As String = Dir_ss & "Азбука права\" ' rev.302
  Const Dir_uv As String = Dir_ss & "ИБ Юридическая пресса\"
  Const Dir_br As String = Dir_ss & "Авторы-бренды\"
  Const Dir_so As String = Dir_ss & "Интернет-статьи\" ' rev.240
  Const Dir_sa As String = Dir_ss & "Рецензирование ИБ Финансист\" ' rev.260
  Const Dir_vopr As String = Dir_ss & "Вопросы под заказ\Базы\" ' rev.230
  
  ' Очищаем коллекции с именами папок и с именами файлов
  Set DirName = Nothing: Set FileName = Nothing: Set Manager = Nothing
  ' Назначаем коллекцию листов для каждого Банка
  GetWorkbooks Dir_ss: GetWorkbooks Dir_br: GetWorkbooks Dir_uv
  GetWorkbooks Dir_vopr, True ' rev.230
  GetWorkbooks Dir_so: GetWorkbooks Dir_az ' Вопросы вперёд! rev.240
  GetWorkbooks Dir_sa ' rev.260
  
  '-> NEXT
End Sub


Private Sub GetWorkbooks(ByVal strDir As String, Optional ByVal QuestStat _
  As Boolean = False) ' Все статистики
  Attribute SettingsBankID.VB_Description = "r302 Запись найденных баз/статистик в коллекцию"
  Dim strName As String: strName = GetMainPath & strDir
  
  ' Возвращаем в strName первый найденный файл по маске *.xl*
  On Error GoTo ErrDir
    If (GetAttr(strName) And vbDirectory) = vbDirectory Then
      With ThisWorkbook
        If DirName.Count = 0 And Right(.Name, 2) = "sm" Then _
          RecLog GetMainPath & "#Finansist\YCHET\Архив\", _
            IIf(.ReadOnly, "Чтение", "Запись")
      End With
      ' Возвращаем в strName первый найденный файл по маске *.xl*
      strName = Dir(strName & "*.xl*", vbNormal)
      Do While strName <> vbNullString ' Выполнять ПОКА
        ' Применяем дополнительную маску для выборки файлов
        If Not strName Like "*отдел*" And Not strName Like "*[Кк]опия*" _
        And Not strName Like "*.lnk" And (IIf(QuestStat, strName _
        Like BASE_VOPR, strName Like "[Сс]татистика*")) Then
          DirName.Add GetMainPath & strDir: FileName.Add strName
        End If: strName = Dir
      Loop
    End If
  Exit Sub

  ErrDir:
  Select Case Err.Number
    Case Is = 53: strName = "Файл не найден: "
    Case Is = 75: strName = "Нет доступа к файлу: "
    Case Is = 457: Exit Sub
    Case Else: strName = "Проверьте сетевой путь. Нет доступа к директории: "
  End Select: MsgBox strName & vbCrLf & GetMainPath & IIf( _
    Err.Number = 76, "", strDir) & IIf(Err.Number = 53, "#FILE", ""), vbCritical
  If Err.Number = 53 Or Err.Number = 76 Then End
End Sub
