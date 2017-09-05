Attribute VB_Name = "mSVCommonSVC"
Option Explicit

Public vConnName As String
Public vServerName As String
Public vAPSServerName As String
Public vAppName As String
Public vDbName As String
Public vUserName As String
Public vPassword As String
Public vFriendlyName As String

Public vConnName_stored As String
Public vServerName_stored  As String
Public vAPSServerName_stored As String
Public vAppName_stored As String
Public vDbName_stored As String
Public vUserName_stored As String
Public vPassword_stored As String
Public vFriendlyName_stored As String
Public vArrFormulas() As Variant

Public vCurrConnectQ As String
 
 Sub p_callDefaulConnectionsMenu()
  'fDefaultConnection.Show (1)
  Call p_checkConnectionError
  If Not HypConnected(Empty) Then
     Call p_ConnectionByDefault
  End If
 End Sub
 
 Sub p_checkInternalConnect()
    Dim sngEnd As Single
    Dim sngElapsed As Single
    sngEnd = Timer
    
    ActiveSheet.AutoFilterMode = False
    ActiveWindow.FreezePanes = False
     
    Call hideTextBox("ConnectQ")
    
    If Not isFirstOptionQ Then
     Call p_RefreshRibbonNow
    End If
 
 If Not HypConnected(Empty) Then

    Call p_callDefaulConnectionsMenu
    
 End If
   
  If Err.Number <> 0 Then
    Err.Clear
  End If
  
  
 End Sub
 

 Sub p_WriteLoginInformation()
 Dim vCurrConnPrefix
   On Error GoTo ErrorHandler
  vCurrConnPrefix = ActiveSheet.Name & ActiveWorkbook.Name
 Call p_WriteGlobalProperty(vCurrConnPrefix & "vConnName", vConnName)
 Call p_WriteGlobalProperty(vCurrConnPrefix & "vServerName", vServerName)
 Call p_WriteGlobalProperty(vCurrConnPrefix & "vAPSServerName", vAPSServerName)
 Call p_WriteGlobalProperty(vCurrConnPrefix & "vAppName", vAppName)
 Call p_WriteGlobalProperty(vCurrConnPrefix & "vDbName", vDbName)
 Call p_WriteGlobalProperty(vCurrConnPrefix & "vUserName", vUserName)
 Call p_WriteGlobalProperty(vCurrConnPrefix & "vPassword", vPassword)
 Call p_WriteGlobalProperty(vCurrConnPrefix & "vFriendlyName", vFriendlyName)
 
  
 Call p_CreateTextBox("ConnectQ", vAppName & "." & vDbName & "@" & vServerName)
 vIsFirstRetrive = True
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(X, "p_WriteLoginInformation")
 End Sub
 
 Sub p_ReadLoginInformation()
 On Error GoTo ErrorHandler
 Dim vCurrConnPrefix
  
 vCurrConnPrefix = ActiveSheet.Name & ActiveWorkbook.Name
 vConnName_stored = f_ReadGlobalProperty(vCurrConnPrefix & "vConnName")
 vServerName_stored = f_ReadGlobalProperty(vCurrConnPrefix & "vServerName")
 vAPSServerName_stored = f_ReadGlobalProperty(vCurrConnPrefix & "vAPSServerName")

 vAppName_stored = f_ReadGlobalProperty(vCurrConnPrefix & "vAppName")
 vDbName_stored = f_ReadGlobalProperty(vCurrConnPrefix & "vDbName")
 vUserName_stored = f_ReadGlobalProperty(vCurrConnPrefix & "vUserName")
 vPassword_stored = f_ReadGlobalProperty(vCurrConnPrefix & "vPassword")
 vFriendlyName_stored = f_ReadGlobalProperty(vCurrConnPrefix & "vFriendlyName")
 
  vCurrConnectQ = getTextBoxValue("ConnectQ")
  
l_exit:
    Exit Sub
ErrorHandler:
vConnName_stored = ""
vAppName_stored = ""
vDbName_stored = ""
vServerName_stored = ""
 vAPSServerName_stored = ""
 'Call p_ErrorHandler(x, "p_ReadLoginInformation")
 End Sub
 


Function getErrorText(vErrNum As Long) As String
 
    getErrorText = vbNewLine & "SmartView error num is " & vErrNum & " : " & GetReturnCodeMessage(vErrNum) & vbNewLine
    If vErrNum = 41 Then
      getErrorText = vbNewLine & vbNewLine & "Check members in the retive slice and database connection "
    End If
    
End Function
 

 

