VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fManageConnectrionLinks 
   Caption         =   "Manage SmartView Quick Links"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4875
   OleObjectBlob   =   "fManageConnectrionLinks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fManageConnectrionLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vCurrTempEnv


 'er@essbase.ru
Private Sub fCancel_Click()
   vCurrEnv = vCurrTempEnv
   Call p_ReadCurrConnectionsINT
   Unload Me
End Sub

Private Sub FManageConectionLinks_Terminate()

    Call fCancel_Click
    
End Sub
 

Private Sub fDelete_Click()
 On Error GoTo ErrorHandler
Dim iCount As Integer, i As Integer
Dim vCurrArrayLine() As String
Dim vCurrStr
iCount = Me.fListLinks.ListCount - 1
For i = 0 To iCount
X = iCount - i
    If Me.fListLinks.Selected(X) Then
      vCurrArrayLine() = Split(fListLinks.List(X), "|")
     Call f_DeleteLineFromCfg(vCurrArrayLine(1))
    End If
Next i

  Call p_ReadCurrConnectionsINT
  Call UserForm_Initialize
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, " fDelete_Click")

End Sub

 
Private Sub LoadRecord(vId As Variant)
 On Error Resume Next
  Dim vCurrArrayLine() As String
  
   vCurrArrayLine() = Split(fArrQuickConnections(vId, 0), "'")
    Me.txtAppName.Text = UCase(vCurrArrayLine(1))
    Me.txtAPSName.Text = vCurrArrayLine(4)
    Me.txtDbName.Text = UCase(vCurrArrayLine(0))
    Me.txtUser.Text = vCurrArrayLine(3)
    Me.txtEssBase = vCurrArrayLine(2)
   'Basic'Sample'localhost'hypadmin'localhost|Sample|Basic|hypadmin|
  If Err.Number <> 0 Then
   Err.Clear
  End If
   
End Sub

 

Private Sub fLoad_Click()
 On Error GoTo ErrorHandler


If Me.fListLinks Is Nothing Then
  MsgBox " Please select line "
  Exit Sub
End If

Dim vStr, vStr2

Dim iCount As Integer, i As Integer, J As Integer

iCount = Me.fListLinks.ListCount - 1
For i = 0 To iCount
X = iCount - i
    If Me.fListLinks.Selected(X) Then
     Dim vCurrArrayLine() As String
   
     vCurrArrayLine() = Split(fListLinks.List(X), "|")
     vStr2 = vCurrArrayLine(1)
       For J = 0 To UBound(fArrQuickConnections, 1)
        If fArrQuickConnections(J, 0) <> "" Then
        vStr = CRC16HASH(fArrQuickConnections(J, 0))
         If InStr(vStr2, vStr) > 0 Then
               Call LoadRecord(J)
         End If
        End If
      Next J
    End If
Next i
  
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, "Manage Conections fLoad_Click")
 
End Sub
Private Function f_testConnect() As Boolean
 On Error GoTo VBAErrorHandler
 Dim VErrMsg
 
 Dim vtFriendlyName As String
 
 f_testConnect = False
 
 If (Me.txtUser.Text = "") Or (Me.txtPass.Text = "") Or (Me.txtAPSName.Text = "") Or (Me.txtEssBase.Text = "") Or (Me.txtAppName.Text = "") Or (Me.txtDbName.Text = "") Then
   MsgBox "Please provide all required information ", vbExclamation
   Exit Function
 End If
 
 
 
 X = f_createConnectionLink((Me.txtUser.Text), Me.txtPass.Text, Me.txtAPSName.Text, Me.txtEssBase.Text, getCamelName(Me.txtAppName.Text), getCamelName(Me.txtDbName.Text))
   If X <> 0 Then
      VErrMsg = "Create connection failed  "
      GoTo ErrorHandler
   End If
   
    vtFriendlyName = CRC16HASH(getCamelName(Me.txtDbName.Text) & "'" & getCamelName(Me.txtAppName.Text) & "'" & Me.txtEssBase.Text & "'" & Me.txtUser.Text & "'" & Me.txtAPSName.Text)
   
 X = HypConnect(Empty, Me.txtUser.Text, Me.txtPass.Text, vtFriendlyName)
 
     If X <> 0 Then
      VErrMsg = "Test to login failed " & vbCrLf & "Check login, password,server name,application name "
      VErrMsg = VErrMsg & vbCrLf & " Try to check login and password using SmartView Panel "
      If X = 1 Then HypRemoveConnection (vtFriendlyName)
      GoTo ErrorHandler
     End If
     
  X = HypDisconnect(Empty, True)
   
    If X <> 0 Then
      VErrMsg = "Error in disconnection "
      GoTo ErrorHandler
     End If
     
    f_testConnect = True
   
l_exit:
    Exit Function
VBAErrorHandler:
    MsgBox Err & ": " & Error(Err) & vbCrLf & " Error Line: " & Erl & vbCrLf & " fTestConnect_Click ", vbCritical
     f_testConnect = False
    Exit Function
ErrorHandler:
    MsgBox VErrMsg & vbCrLf & vbCrLf, vbExclamation
    ' MsgBox VErrMsg & vbCrLf & vbCrLf & " f_testConnect  " & getErrorText(x), vbExclamation
    f_testConnect = False
End Function

Private Sub fSave_Click()
 On Error GoTo ErrorHandler

    Dim vArrOfStrings() As String, vCurrStr As String
    Dim i As Long, J As Long
    Dim vConnName As String
    
     If Not f_testConnect() Then
       MsgBox "Test Failed"
      Exit Sub
     End If
    
    Call GetEssbaseRibonnConnectionFileName
    SetAttr vRibbonSetFileName, vbNormal
    'f_createConnectionLink(getCamelName(Me.txtUser.Text), Me.txtPass.Text, Me.txtAPSName.Text, getCamelName(Me.txtEssBase.Text), getCamelName(Me.txtAppName.Text), getCamelName(Me.txtDbName.Text))
     vConnName = getCamelName(Me.txtDbName.Text) & "'" & getCamelName(Me.txtAppName.Text) & "'" & getCamelName(Me.txtEssBase.Text) & "'" & getCamelName(Me.txtUser.Text) & "'" & Me.txtAPSName.Text
     'Call f_DeleteLineFromCfg(vConnName)
     Call p_deleteRecordByName(Me.txtAppName.Text, Me.txtEssBase.Text)
    
    Open vRibbonSetFileName For Append As #1
   
    Print #1, vConnName & "|" & _
                getCamelName(Me.txtAppName.Text) & "|" & _
                getCamelName(Me.txtDbName.Text) & "|" & _
                getCamelName(Me.txtUser.Text) & "|" & _
                f_XOREncryption(Me.txtPass.Text, f_XOREncryption(vCurrPasswordLine, VBA.Environ("Computername") & VBA.Environ("Username")))
  
    Close #1
   SetAttr vRibbonSetFileName, vbHidden
   Call p_ReadCurrConnectionsINT
   Call updateListLinks
   
   MsgBox "Saved."
    Call p_setRegNetworkValues
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, "Manage Conections fSave_Click")
End Sub

 
Private Sub fTestConnect_Click()
 On Error Resume Next
  
 If f_testConnect() Then
    MsgBox "Test Passed", vbInformation
  Else
    MsgBox " Test Failed ", vbExclamation
 End If
 
 End Sub


 
  

Private Sub initialaseHeader()
    Me.txtAPSName.Text = ""
    Me.txtAppName.Text = ""
    Me.txtDbName.Text = ""
    Me.txtUser.Text = ""
    Me.txtEssBase = ""
    Call LoadRecord(50)
End Sub
Private Sub updateListLinks()
 On Error GoTo ErrorHandler
 Dim i As Long, J As Long
 Dim vCurrServer() As String
 
 Dim vCurrSpace
    fListLinks.Clear
     On Error Resume Next
     Call p_ReadCurrConnectionsINT
 
    For i = 0 To UBound(fArrQuickConnections)
        If fArrQuickConnections(i, 0) <> "" Then 'Exit Sub
          vCurrServer() = Split(fArrQuickConnections(i, 5), ".")
          vCurrSpace = LCase(vCurrServer(0)) & Chr(9) & LCase(fArrQuickConnections(i, 1)) & Chr(9) & "." & LCase(fArrQuickConnections(i, 2))
         
         Me.fListLinks.AddItem vCurrSpace & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "|" & CRC16HASH(fArrQuickConnections(i, 0))     ' fArrQuickConnections(i, 0)
       
       End If
    Next i
    
    If Err.Number <> 0 Then
     Err.Clear
    End If
    
   On Error GoTo ErrorHandler
   
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, "Manage Conections updateListLinks")
End Sub

Private Sub p_deleteRecordByName(vCubeName, vSvrName)
 On Error GoTo ErrorHandler
 Dim i
     Call p_ReadCurrConnectionsINT
    For i = 0 To UBound(fArrQuickConnections)
        If InStr(fArrQuickConnections(i, 0), vCubeName) > 0 And InStr(fArrQuickConnections(i, 5), vSvrName) > 0 Then
         f_DeleteLineFromCfg (CRC16HASH(fArrQuickConnections(i, 0)))
       End If
    Next i
l_exit:
    Exit Sub
ErrorHandler:
  'Call p_ErrorHandler(0, "p_deleteRecordByName")
  Err.Clear
End Sub



Private Sub UserForm_Initialize()
 On Error GoTo ErrorHandler
 Dim i As Long, J As Long
 Dim vCurrArrayLine() As String
 Dim vCurrSpace
 
   vCurrTempEnv = vCurrEnv
 
   vCurrEnv = 0
   
   Call updateListLinks
   Call initialaseHeader
   Call p_ReadCurrConnectionsINT
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, "Manage Conections UserForm_Initialize")
End Sub


 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
  'er@essbase.ru

