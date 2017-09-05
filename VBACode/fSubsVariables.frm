VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fSubsVariables 
   Caption         =   "Substitution Variables"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   OleObjectBlob   =   "fSubsVariables.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fSubsVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public vtApplicationName As Variant
 Public vtDatabaseName As Variant

Private Sub cmdbCancel_Click()
   Unload Me
End Sub

Private Sub cmdbEdit_Click()
 On Error GoTo ErrorHandler
Dim iCount As Integer, i As Integer

iCount = Me.ListVariables.ListCount - 1
Dim vCurrArrayString() As String
For i = 0 To iCount
X = iCount - i
    If Me.ListVariables.Selected(X) Then
    ' Call f_DeleteLineFromCfg(fListLinks.List(x))
      vCurrArrayString() = Split(Me.ListVariables.List(X), Chr(9))
    Me.VariableName.Text = vCurrArrayString(1)
    Me.VariableValue.Text = vCurrArrayString(3)
    
    End If
Next i
 
l_exit:
    Exit Sub
ErrorHandler:
  MsgBox Err & ": " & Error(Err) & vbCrLf & " Error Line: " & Erl & vbCrLf & " fDelete_Click "
End Sub

 

Private Sub cmdbNewSave_Click()
  On Error GoTo ErrorHandler
 
 If Me.VariableName.Text = "" Then
  Exit Sub
 End If
 
      X = HypSetSubstitutionVariable(Empty, vtApplicationName, vtDatabaseName, Replace(Me.VariableName.Text, "&", ""), Me.VariableValue.Text)
  If X <> 0 Then
    GoTo ErrorHandler
  End If
  Call UserForm_Initialize
  
l_exit:
    Exit Sub
ErrorHandler:
    Call p_ErrorHandler(X, " Subs Variables cmdbNewSave_Click")
    
End Sub

Private Sub UserForm_Initialize()
 On Error GoTo ErrorHandler
 Dim i As Long, J As Long
 Dim vCurrArrayLine() As String
 
    ListVariables.Clear
   
 Dim vtSheetName As Variant

Dim vtVariableNames As Variant
Dim vtVariableValues As Variant

Dim vTempVariableNames As String
Dim vTempVariableValues As String

  Call p_setExcelCalcOff
  
 vtApplicationName = vAppName_stored
 vtDatabaseName = vDbName_stored
 
 If vtApplicationName = "" Then
   vtApplicationName = vAppName
   vtDatabaseName = vDbName
 End If
 
  Call p_checkInternalConnect
  
      If ActiveSheet Is Nothing Then
        MsgBox "active sheet is not determinated "
        GoTo l_exit
    End If
    
 vIsConnected = HypConnected(Empty)
 If vIsConnected Then
 
      X = HypGetSubstitutionVariable(Empty, vtApplicationName, vtDatabaseName, Empty, vtVariableNames, vtVariableValues)
      
      If X <> 0 Then
        GoTo ErrorHandler
      End If
      
    Me.TextAppName.Text = "Application:" & vtApplicationName & " Database:" & vtDatabaseName
    
    For J = 0 To UBound(vtVariableNames)
        For i = 0 To UBound(vtVariableNames) - 1
              If vtVariableNames(i) > vtVariableNames(i + 1) Then
                vTempVariableNames = vtVariableNames(i)
                vTempVariableValues = vtVariableValues(i)
                vtVariableNames(i) = vtVariableNames(i + 1)
                vtVariableValues(i) = vtVariableValues(i + 1)
                vtVariableNames(i + 1) = vTempVariableNames
                vtVariableValues(i + 1) = vTempVariableValues
              End If
        Next
    Next
     
        For i = 0 To UBound(vtVariableNames)
           If Len(vtVariableNames(i)) > 10 Then
               Me.ListVariables.AddItem i + 1 & Chr(9) & "&" & vtVariableNames(i) & Chr(9) & Chr(9) & vtVariableValues(i)
            Else
               Me.ListVariables.AddItem i + 1 & Chr(9) & "&" & vtVariableNames(i) & Chr(9) & Space(12) & Chr(9) & vtVariableValues(i)
            End If
        Next
 End If
   Call p_setExcelCalcOn_INT
l_exit:
    Exit Sub
ErrorHandler:
    Call p_ErrorHandler(X, " Subs Variables UserForm_Initialize" & "SV Error Code :" & X)
End Sub

