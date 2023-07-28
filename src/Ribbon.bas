Attribute VB_Name = "Ribbon"
Option Explicit
Option Base 1
'123456789012345678901234567890123456h8nor@уа56789012345678901234567890123456789

Private Sub Ribbon_GetEnabledMacro(ByVal objRC As IRibbonControl, ByRef Enabled)
  Attribute Ribbon_GetEnabledMacro.VB_Description = "r315 ¦ Подключение макросов пользовательского меню"
  Debug.Print "Sub Ribbon_GetEnabledMacro"
End Sub

Private Sub Ribbon_GetVisibleMenu(ByVal objRC As IRibbonControl, ByRef Visible)
  Attribute Ribbon_GetVisibleMenu.VB_Description = "r315 ¦ Видимость объектов пользовательского меню"
  Debug.Print "Sub Ribbon_GetVisibleMenu"
End Sub

Private Sub Ribbon_Initialize(ByVal objRibbonUI As IRibbonUI) ' currentUI.onLoad
  Attribute Ribbon_Initialize.VB_Description = "r315 ¦ Подключение пользовательского меню"
  Debug.Print "Sub Ribbon_Initialize"
End Sub
