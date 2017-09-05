VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fExecuteMDX 
   Caption         =   "Execute MDX"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   OleObjectBlob   =   "fExecuteMDX.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fExecuteMDX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub p_ExecuteCurrentMDX(vCurrMDX As String)
 On Error Resume Next
  Call p_deleteAllTextBox("MDXq")
  Call p_CreateTextBox("MDXq", "" & vCurrMDX)
  Call hideTextBox("MDXq")
  
  If InStr(vCurrMDX, "MEMBER_NAME") <> 0 Then
   vIsUseNameDefault = False
  End If
  
  If InStr(vCurrMDX, "ALIAS") <> 0 Then
   vIsUseNameDefault = True
  End If
 
  Unload Me
  Call p_ExecStoredMDXINT
 
 If X = -15 Then
   Call p_ErrorHandler(X, " Check MDX Statment in the qeury box  (bExecuteMDX_Click)")
   X = 0
 End If
 
 If X <> 0 Then
    Call p_ErrorHandler(X, " p_ExecuteCurrentMDX  ")
     X = 0
  End If
   

 If Err.Number <> 0 Then
      Err.Clear
 End If
   
 Call p_setExcelCalcOn
   
   
End Sub

Private Sub p_ExecuteMDXBooks(vGlobalMDX As String)
  On Error GoTo ErrorHandler
  Dim vCurrMDX As String
  Dim vIsAsk As Boolean
  Dim vIsAlias  As Boolean
  Unload Me
   Dim vCurrentSheet As Worksheet
   X = 0
   Call p_setExcelCalcOff
   
  vIsAsk = True
  If InStr(vGlobalMDX, "NOASK") <> 0 Then
   vIsAsk = False
  End If
  
  If InStr(vGlobalMDX, "MEMBER_NAME") <> 0 Then
   vIsUseNameDefault = False
  End If
  
  If InStr(vGlobalMDX, "ALIAS") <> 0 Then
   vIsUseNameDefault = True
  End If
    
   vIsAlias = vIsUseNameDefault
   
For Each vCurrentSheet In Worksheets
 If (InStr(UCase(vCurrentSheet.Name), "OTL") = 0 And vCurrentSheet.Visible) Then
 X = 1
 If (vIsAsk) Then
  X = MsgBox("Execute MDX on " & vCurrentSheet.Name & "?", vbOKCancel, "Global MDX")
 End If
       If X = 1 Then
        X = 0
        vCurrentSheet.Activate
        vCurrMDX = getTextBoxValue("MDXq")
        If Len(vCurrMDX) > 10 Then
          vIsUseNameDefault = vIsAlias
          If InStr(vCurrMDX, "MEMBER_NAME") <> 0 Then
             vIsUseNameDefault = False
            End If
            
            If InStr(vCurrMDX, "ALIAS") <> 0 Then
             vIsUseNameDefault = True
            End If
            
         Call p_ExecStoredMDXINT '  MsgBox vCurrMDX
        End If
               If X = -15 Then
                   Call p_ErrorHandler(X, " Check MDX Statment in the qeury box  (bExecuteMDX_Click)")
                   X = 0
                   Call p_setExcelCalcOn
                   Exit Sub
                 End If
                
                
                If X <> 0 Then
                  Call p_ErrorHandler(X, " p_ExecuteMDXBooks  ")
                   Call p_setExcelCalcOn
                 Exit Sub
       End If
  End If
 
 If Err.Number <> 0 Then
      Err.Clear
 End If
 
 End If
  
   '
 
Next

 On Error Resume Next
 ActiveWorkbook.Sheets("OTL").Activate
 Call p_setExcelCalcOn
 Exit Sub
  
ErrorHandler:
 Call p_ErrorHandler(0, "p_ExecuteMDXBooks")
 Call p_setExcelCalcOn
  End
End Sub


Private Sub bExecuteMDX_Click()
 On Error Resume Next
    
     
  If ActiveSheet Is Nothing Then
        MsgBox "active sheet is not determinated "
        Exit Sub
    End If
    
 Dim vCurrMDX
  vCurrMDX = UCase(Me.mdxTextBox.Text)
 If InStr(vCurrMDX, "FROM") = 0 Or InStr(vCurrMDX, "SELECT") = 0 Then
    MsgBox " Check MDX Statment "
   Exit Sub
 End If
 
 If InStr(vCurrMDX, "ALL_SHEETS") <> 0 Then
  If InStr(UCase(ActiveSheet.Name), "OTL") = 0 Then
   MsgBox ("You can execute MDX from all sheets on only from page named 'OTL' ")
    Unload Me
   End
 End If
   X = MsgBox("Execute all MDX on this Books  ?", vbOKCancel, "Global MDX")
       If X <> 1 Then
         Unload Me
         End
       End If
         Call p_ExecuteMDXBooks("" & vCurrMDX)
  Else
    Call p_ExecuteCurrentMDX("" & vCurrMDX)
 End If
End Sub

 

Private Sub cb_name_Click()
  vIsUseNameDefault = cb_name.value
End Sub

Private Sub fCancel_Click()
 Unload Me
End Sub

Private Sub fSave_Click()
 
 Dim vCurrMDX
  vCurrMDX = UCase(Me.mdxTextBox.Text)
 If InStr(vCurrMDX, "FROM") = 0 Or InStr(vCurrMDX, "SELECT") = 0 Then
    MsgBox " Check MDX Statment "
   Exit Sub
 End If
 
  Call p_deleteAllTextBox("MDXq")
  Call p_CreateTextBox("MDXq", "" & vCurrMDX)
  Call hideTextBox("MDXq")
 Unload Me
 
End Sub

Private Sub UserForm_Initialize()
 On Error GoTo ErrorHandler
 p_RefreshRibbonNow
  'Me.Describes.Text = "Enter a valid MDX statement."
 cb_name.value = vIsUseNameDefault
  ' Me.mdxTextBox.Text =
   With Me.mdxTextBox
    .value = getTextBoxValue("MDXq")
    .SetFocus
   ' .SelStart = 0
   ' .SelLength = Len(.Text)
End With
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, "ExecuteMDX: UserForm_Initialize ")
  Unload Me
End Sub

 
