Attribute VB_Name = "mRibbonSVC"
 Option Explicit
  
 Public isFirstOptionQ As Boolean
 Public vIsSuppresPressed As Boolean
 
  Public vIRibbonUI As IRibbonUI
 
Sub p_RefreshRibbonNow()
On Error Resume Next
  Call p_restoreOptions
   vIRibbonUI.Invalidate
    If Err.Number > 0 Then
         Err.Clear
    End If
    isFirstOptionQ = True
End Sub

Sub p_OnRibbonLoad(vRibbon As IRibbonUI)

 
 Application.MultiThreadedCalculation.Enabled = True
 Application.AutoRecover.Time = 7
 Application.EnableEvents = True
 
    Set vIRibbonUI = vRibbon
    vItWasOtlPage = False
      vConnName = ""
      vAppName = ""
      vDbName = ""
      vUserName = ""
      vPassword = ""
      vFriendlyName = ""
      vIsSVEnabled = False
      vIsSuppresPressed = False
      vCurrEnv = 0
      X = HypSetMenu(False)
      vIsFirstRetrive = True
      vIsUseNameDefault = True
      isFirstOptionQ = True
      vCurrXMLGlobal = ""
 
End Sub

 
 







