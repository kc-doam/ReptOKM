Attribute VB_Name = "Ribbon"
Option Explicit
Option Base 1
Option Private Module
'123456789012345678901234567890123456h8nor@уа56789012345678901234567890123456789

Private Sub Ribbon_GetEnabledMacro(ByVal objRC As IRibbonControl, ByRef Enabled)
  Attribute Ribbon_GetEnabledMacro.VB_Description = "r316 ¦ Подключение макросов пользовательского меню"
  HookMsg "Sub Ribbon_GetEnabledMacro", vbRetryCancel
End Sub

Private Sub Ribbon_GetVisibleMenu(ByVal objRC As IRibbonControl, ByRef Visible)
  Attribute Ribbon_GetVisibleMenu.VB_Description = "r316 ¦ Видимость объектов пользовательского меню"
  HookMsg "Sub Ribbon_GetVisibleMenu", vbRetryCancel
End Sub

Private Sub Ribbon_Initialize(ByVal objRibbonUI As IRibbonUI) ' currentUI.onLoad
  Attribute Ribbon_Initialize.VB_Description = "r316 ¦ Подключение пользовательского меню"
  HookMsg "Sub Ribbon_Initialize", vbRetryCancel
End Sub
