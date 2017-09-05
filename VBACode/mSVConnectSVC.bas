Attribute VB_Name = "mSVConnectSVC"
Option Explicit

Public Sub p_checkConnectionsArray()
Dim J
On Error Resume Next
  J = UBound(fArrQuickConnections)
 If Err.Number <> 0 Then
   Call p_ReadCurrConnections
   
  Err.Clear
 End If
End Sub

Function getConnectionID(ByVal vConectionsName As String, Optional ByVal vSVR As String = "")

On Error GoTo ErrorHandler

Dim i, J, vCurrDBString, jCount
Dim vCurrArrayLine() As String
Dim vIsMyServer

Call p_checkConnectionsArray
   On Error Resume Next
  For i = 0 To UBound(fArrQuickConnections, 1)
If (Len(fArrQuickConnections(i, 1)) > 3) Then
  If (("" = vSVR Or ((InStr(UCase(fArrQuickConnections(i, 5)), UCase(vSVR))) > 0))) Then
   vIsMyServer = 0
      vCurrArrayLine() = Split(fArrQuickConnections(i, 0), "'")
    If ((InStr(UCase(vCurrArrayLine(2)), UCase(vSVR))) > 0) Then '
      vIsMyServer = 1
    End If
    vCurrArrayLine() = Split(vConectionsName, "@")
    If vCurrArrayLine(1) = "0" Then
     If Err.Number = 0 Then
      vIsMyServer = 1
      End If
      Err.Clear
    End If
    
   vCurrDBString = UCase(fArrQuickConnections(i, 1) & "." & fArrQuickConnections(i, 2))
    
    If (vIsMyServer = 1) Then
        If InStr(vConectionsName, vCurrDBString) > 0 Then
          jCount = jCount + 1
          J = i
          Exit For
        End If
     End If
  End If
 End If
  
  Next
    On Error GoTo ErrorHandler
  If jCount <> 1 Then
     MsgBox " Can't find database connection  automaticaly. 1" & vbNewLine & " Please chose default connection with using Quick Connect Menu & p_makeConnectionsByID ", vbExclamation
     End
  End If
   
getConnectionID = J
 
 Exit Function
    
ErrorHandler:
Call p_ErrorHandler(X, "getConnectionID")
 
End Function


Sub p_makeConnectionsByID(ByVal vConectionsName As String, Optional ByVal vSVR As String = "")

On Error GoTo ErrorHandler
Dim J
  J = getConnectionID(vConectionsName, vSVR)
   
 Call p_ConnectionByMenuId(J, False)
  p_RefreshRibbonNow
 Exit Sub
    
ErrorHandler:
Call p_ErrorHandler(X, "p_makeConnectionsByID")

End Sub

 

Sub p_connnectByServerAndDatabaseName(ByVal vConectionsName As String, Optional ByVal vSVR As String = "")     ' App.DB
On Error GoTo ErrorHandler

Dim i, J, isDefault, vCurrDBString, jCount

Dim vCurrConectionNameOnSheet
 
 vCurrConectionNameOnSheet = getCurrSheetConnectionName()
If vCurrConectionNameOnSheet = "" Then
   Call p_makeConnectionsByID(vConectionsName, vSVR)
    Exit Sub
End If

 If Not HypConnected(Empty) Then
 If InStr(UCase(vCurrConectionNameOnSheet), UCase(vConectionsName)) > 0 Then
          Call p_makeConnectionsByID(vCurrConectionNameOnSheet, "")
      Exit Sub
    Else
     If vSVR = "" Then
      MsgBox " This Excel Sheet is already connected to  " & vCurrConectionNameOnSheet & " Please check App.DBName or create new connections ", vbExclamation
     End
     End If
    End If
 
 Call p_makeConnectionsByID(vConectionsName, vSVR)
 vCurrConnectQ = getTextBoxValue("ConnectQ")
End If

   
l_exit:
    Exit Sub
    
ErrorHandler:
Call p_ErrorHandler(X, "p_connnectByServerAndDatabaseName")

End Sub


Public Function f_createConnectionLink(ByVal vtUser As Variant, ByVal vtPassword As Variant _
                                    , ByVal vtProviderURL As Variant _
                                    , ByVal serverName As Variant, ByVal vtApp As Variant _
                                    , ByVal vtDB As Variant) As Long
  On Error GoTo ErrorHandler
 Dim vtFriendlyName As String
 '
 vtFriendlyName = CRC16HASH(vtDB & "'" & vtApp & "'" & serverName & "'" & vtUser & "'" & vtProviderURL)
 
  X = HypDisconnect(Empty, True)
  
 Dim isCreatedConnection As Boolean
 
 isCreatedConnection = HypConnectionExists(vtFriendlyName)
 
 If isCreatedConnection Then
   X = HypDisconnectAll()
   X = HypInvalidateSSO()
   X = HypRemoveConnection(vtFriendlyName)
 End If
 
 X = HypCreateConnection(Empty, vtUser, vtPassword, HYP_ESSBASE, _
    "http://" & vtProviderURL & ":13080/aps/SmartView", serverName, vtApp, _
    vtDB, vtFriendlyName, "")
      
 If X <> 0 Then GoTo ErrorHandler
  
   X = HypDisconnect(Empty, True)
   f_createConnectionLink = 0
  vCurrConnectQ = getTextBoxValue("ConnectQ")
l_exit:
    Exit Function
ErrorHandler:
f_createConnectionLink = HypRemoveConnection(vtFriendlyName)
Call p_ErrorHandler(X, "f_createConnection")
    
End Function

Function getCurrSheetConnectionName() As String
Dim vResult
 
 getCurrSheetConnectionName = ""
If isTextBoxPresent("ConnectQ") Then
'  vResult = Split(getTextBoxValue("ConnectQ"), "@")
  getCurrSheetConnectionName = getTextBoxValue("ConnectQ")
End If

 getCurrSheetConnectionName = getClearString(getCurrSheetConnectionName)
 
End Function

 
 

Sub p_SetEnvID(vSheetName As String)     ' vCurrEnv
 Dim vResultStr
 Dim vTextBoxStr
 
 If ActiveSheet.Name = vSheetName Then
 vResultStr = "@" & vCurrEnv
 
 
If isTextBoxPresent("ConnectQ") Then
  vTextBoxStr = Split(getTextBoxValue("ConnectQ"), "@")
  vResultStr = vTextBoxStr(0) & vResultStr
  If vTextBoxStr(0) = "" Then
   Call p_deleteAllTextBox("ConnectQ")
  Else
    Call p_CreateTextBox("ConnectQ", "" & vResultStr)
  End If
End If

 End If
 
End Sub

  Sub p_ConnectionByDefault()
  On Error GoTo ErrorHandler
  p_svcDisconnectSheet
 Dim vConnName
  Call p_ReadLoginInformation
  
  If vConnName_stored = "" Then
  
   vConnName = getCurrSheetConnectionName()
   
  If vConnName = "" Then
    If Len(vUserName) = 0 Then
      MsgBox "Please choose default connection with using Quick Connect Menu", vbExclamation
      Exit Sub
    End If
    If f_SVLoginComplex(vUserName, vPassword, vFriendlyName, vServerName, vAppName, vDbName, vAPSServerName) Then
             MsgBox " Connected by default to " & vAppName & " : " & vDbName
    End If
   Else
    Call p_connnectByServerAndDatabaseName(vConnName)
      MsgBox " Connected  by stored name :  " & vConnName
     ' End
   End If
     
 Else
  
  vConnName = vConnName_stored
  vAppName = vAppName_stored
  vServerName = vServerName_stored
  vAPSServerName = vAPSServerName_stored
  vDbName = vDbName_stored
  vUserName = vUserName_stored
  vPassword = vPassword_stored
  vFriendlyName = vFriendlyName_stored
  
   If f_SVLoginComplex(vUserName, vPassword, vFriendlyName, vServerName, vAppName, vDbName, vAPSServerName) Then
       MsgBox " Connected from last QC : " & vAppName & " : " & vDbName
       End
   End If
    
 End If
   ' p_RefreshRibbonNow
   vCurrConnectQ = getTextBoxValue("ConnectQ")
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(X, " Create connection failed on p_ConnectionByDefault")
 p_svcDisconnectSheet
End Sub
  
Sub p_ConnectionByMenuId(ByVal vMenuId As Variant, Optional ByVal iSshowMsgBox As Boolean = True)

  On Error GoTo ErrorHandler
  
 Dim vCurrArrayLine() As String
   
   'vIsConnected = HypConnected(Empty)
  
 If HypConnected(Empty) Then
   X = HypDisconnect(Empty, True)
   X = HypDisconnect(Empty, True)
 End If
 
  vConnName = fArrQuickConnections(vMenuId, 0)
  vAppName = fArrQuickConnections(vMenuId, 1)
  vDbName = fArrQuickConnections(vMenuId, 2)
  vUserName = fArrQuickConnections(vMenuId, 3)
  vPassword = f_XORDecryption(fArrQuickConnections(vMenuId, 4), f_XOREncryption(vCurrPasswordLine, VBA.Environ("Computername") & VBA.Environ("Username")))
  vFriendlyName = CRC16HASH(vConnName)
  
   vCurrArrayLine() = Split(vConnName, "'")
   vServerName = vCurrArrayLine(2)
   vAPSServerName = vCurrArrayLine(4)
 
   If f_SVLoginComplex(vUserName, vPassword, vFriendlyName, vServerName, vAppName, vDbName, vAPSServerName) Then
      If iSshowMsgBox Then
        MsgBox " Connected  to " & vAppName & " : " & vDbName
      End If
      Call p_WriteLoginInformation
   End If
  
 'p_RefreshRibbonNow
 vCurrConnectQ = getTextBoxValue("ConnectQ")
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(X, " Create connection failed on p_ConnectionByMenuId")
 p_svcDisconnectSheet
End Sub

Sub p_checkConnectionError()
p_svcDisconnectSheet
If X = 1001 Then
     MsgBox " Lost connection by TimeOut reason  "
     Call p_ConnectionByDefault
 End If

  If X = 4 Then
     MsgBox "You are need to create connection first"
     Call p_ConnectionByDefault
  End If
End Sub


 Function f_SVLoginComplex(vUserNameCurr As Variant, vPasswordCurr As Variant, vFriendlyNameCurr As Variant, vServerNameCurr As Variant, vAppNameCurr As Variant, vDbNameCurr As Variant, vAPSServerNameCurr As Variant) As Boolean
 
  On Error GoTo ErrorHandler
  
  If vIsSVEnabled Then
     vIsSVEnabled = False
     X = HypSetMenu(False)
  End If
  
  Dim ErrMsg
     X = HypDisconnect(Empty, True)
      p_svcDisconnectSheet
      
     X = HypConnect(Empty, vUserNameCurr, vPasswordCurr, vFriendlyNameCurr)
     
     If X = -15 Then
        X = f_createConnectionLink(vUserNameCurr, vPasswordCurr, vAPSServerNameCurr, vServerNameCurr, vAppNameCurr, vDbNameCurr)
        X = HypConnect(Empty, vUserNameCurr, vPasswordCurr, vFriendlyNameCurr)
     End If
     
    If X = 100000 Or X = 4 Then
     p_svcDisconnectSheet
       Call p_ErrorHandler(0, " Error in Creating connection. Please check Essbase server is running ")
    End
    End If
     
    If X = -523077684 Or X = -522029176 Then
      X = HypDisconnect(Empty, True)
      p_svcDisconnectSheet
        Call p_ErrorHandler(0, " check login and password ")
    End
    End If
 
    If X = -522029289 Then '-522029289
       X = HypDisconnect(Empty, True)
       p_svcDisconnectSheet
       Call p_ErrorHandler(0, " Error in Creating connection. Please copy retrive slice to the new sheet and try again.")
    End
    End If

    
    If X = -4 Then
       X = HypDisconnect(Empty, True)
       p_svcDisconnectSheet
       X = HypConnect(Empty, vUserNameCurr, vPasswordCurr, vFriendlyNameCurr)
    End If
 
   If X <> 0 Then
         ErrMsg = 1
         GoTo ErrorHandler
   End If
  

    X = HypSetActiveConnection(vFriendlyNameCurr)
  
    If X <> 0 Then
     X = HypDisconnect(Empty, True)
     p_svcDisconnectSheet
       Call p_ErrorHandler(0, " Error in setting default connection. Please logon again.")
     End
    End If
     
     X = HypSetAsDefault(vFriendlyNameCurr)
     
   If X <> 0 Then
     ErrMsg = 3
     GoTo ErrorHandler
    End If
 

    Call p_WriteLoginInformation
    
   vCurrConnectQ = getTextBoxValue("ConnectQ")
l_exit:
f_SVLoginComplex = True
    Exit Function
ErrorHandler:
 X = HypDisconnect(Empty, True)
Call p_ErrorHandler(X, "f_SVLoginComplex " & ErrMsg)
f_SVLoginComplex = False
 End Function


