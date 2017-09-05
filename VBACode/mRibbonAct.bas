Attribute VB_Name = "mRibbonAct"
Option Explicit
 
 
 Public vIsPriorConnect As Boolean
 Public StartExcelTime   As Single
 Public vIsSVEnabled As Boolean
 Public isMDXSlice As Boolean


Public Sub p_CheckConnectionINT()

 If ActiveSheet Is Nothing Then
      MsgBox "active sheet is not determinated ", vbExclamation
    End
 End If

If Not isConnectPresent Then
    MsgBox "Can't find connection link. Please create Quick Connect", vbExclamation
    Call p_RefreshRibbonNow
  End
End If

 

End Sub

Public Sub p_CheckConnection()

If isSheetOTL Then
   MsgBox "You can't connect from OTL page", vbExclamation
   End
Else
 Call p_CheckConnectionINT
 Call p_checkInternalConnect
 
End If


End Sub

Public Sub p_LastConnectNew(vIRibbonControl As IRibbonControl)

  If InStr(UCase(ActiveSheet.Name), "QUERY") Then
      MsgBox "You can't connect from Query page. Please use other sheets", vbExclamation
    End
  End If
  
   Call p_CheckConnection
   Call p_ReadLoginInformation
   
   Call p_SheetInfo(vIRibbonControl)
End Sub


 Public Sub p_SubsVariables(vIRibbonControl As IRibbonControl)
  On Error GoTo ErrorHandler
  
  ActiveSheet.Cells(1, 1).Select

 Call p_CheckConnection
 
 Call p_ReadLoginInformation
 
 If vAppName_stored = "" And vAppName = "" Then
  MsgBox "Can't find current connection. Please reconect", vbExclamation
  End
 End If
 

    fSubsVariables.Show (1)

 
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(X, "p_SubsVariables")
 End Sub

 
 Public Sub p_qConnectAction(vIRibbonControl As IRibbonControl)
 
  On Error GoTo ErrorHandler
  Dim vtFriendlyName As String
  Dim isConnection
  
  isConnection = True
 
  If InStr(UCase(ActiveSheet.Name), "QUERY") Then
      MsgBox "You can't connect from Query page. Please use other sheets", vbExclamation
    End
  End If
  
  
  
    If vIRibbonControl.ID = "b_EditQConnect" Then
      fManageConnectrionLinks.Show (1)
      isConnection = False
    End If
    
    If isConnection And vModeAnalyse = 0 Then
      Call p_ConnectionByMenuId(Replace(vIRibbonControl.ID, "b_qConnect", ""))
    End If
    
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(X, "p_qConnectAction")
 

End Sub

 Public Sub p_EditQConnect(vIRibbonControl As IRibbonControl)
 
  On Error GoTo ErrorHandler
 
      fManageConnectrionLinks.Show (1)
   
    
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(X, "p_EditQConnect")
 

End Sub



Sub p_SVon(ByVal vIRibbonControl As IRibbonControl)
 On Error GoTo ErrorHandler

  X = HypSetMenu(True)
    If X <> 0 Then GoTo ErrorHandler
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(X, "p_SVon")
 
End Sub
 
 
           
 Sub p_Connections(ByVal vIRibbonControl As IRibbonControl)  '(ByVal vIRibbonControl As IRibbonControl, ByRef vPressed)
 
On Error Resume Next
 vIsSVEnabled = Not vIsSVEnabled
 HypSetMenu (vIsSVEnabled)
 X = HypExecuteMenu(Empty, "Smart View->Panel")
 X = 0
 Err.Clear
 
End Sub

 Sub p_MemberInfo(ByVal vIRibbonControl As IRibbonControl)  '(ByVal vIRibbonControl As IRibbonControl, ByRef vPressed)
 
On Error Resume Next
 
 X = HypExecuteMenu(Empty, "EssBase->Member Information")
 X = 0
 Err.Clear
 
End Sub

Sub p_Options(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
ActiveSheet.Cells(1, 1).Select

    If ActiveSheet Is Nothing Then
        MsgBox "active sheet is not determinated "
        GoTo l_exit
    End If

     X = HypMenuVOptions()
       If X <> 0 And X <> -55 Then GoTo ErrorHandler
l_exit:
    Exit Sub
ErrorHandler:
'Call p_ErrorHandler(x, "p_Options")
 
End Sub




Sub p_Disconnect(ByVal vIRibbonControl As IRibbonControl)

   p_svcDisconnect
 
  Call p_SheetInfo(vIRibbonControl)
   
End Sub


Sub p_DisconnectOld(ByVal vIRibbonControl As IRibbonControl)
  
  X = HypDisconnect(Empty, True)
  
  Call p_SheetInfo(vIRibbonControl)
End Sub


Sub p_DisconnectAll(ByVal vIRibbonControl As IRibbonControl)
  
         Dim WS_Count As Integer
         Dim i As Integer
         
         WS_Count = ActiveWorkbook.Worksheets.Count
 
         For i = 1 To WS_Count
             X = HypDisconnect(ActiveWorkbook.Worksheets(i).Name, True)
         Next i
    Call p_SheetInfo(vIRibbonControl)
End Sub
 Sub p_setPOVe(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

ActiveSheet.Cells(1, 1).Select

  X = HypExecuteMenu(Empty, "Essbase->POV")
l_exit:
    Exit Sub
ErrorHandler:
 If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
 Else
   Call p_ErrorHandler(X, "p_setPOVe")
 End If

 
End Sub
Sub p_SheetInfo(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

vIsSVEnabled = False
X = HypSetMenu(False)

 'Call p_CheckConnectionINT
 'ActiveSheet.Cells(1, 1).Select

        X = HypExecuteMenu(Empty, "Smart View->Sheet Info")
l_exit:
    Exit Sub
ErrorHandler:
If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
 Else
   Call p_ErrorHandler(X, "p_SheetInfo")
 End If
 
End Sub


Sub p_setAliasTable(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

 Call p_CheckConnection
 
   Call p_setExcelCalcOff
     X = HypExecuteMenu(Empty, "Essbase->Change Alias")
  Call p_setExcelCalcOn
l_exit:
    Exit Sub
ErrorHandler:
If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
 Else
   Call p_ErrorHandler(X, "p_setAliasTable")
 End If
 
End Sub

Sub p_Pivot(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
Dim bErrLine

bErrLine = 1
 Call p_CheckConnection
bErrLine = bErrLine + 1

 Call p_setExcelCalcOff
 bErrLine = bErrLine + 1
    X = HypMenuVPivot()
    bErrLine = bErrLine + 1
 Call p_setExcelCalcOn
 bErrLine = bErrLine + 1
 
 
 isMDXSlice = False
  X = HypShowPov(False)
  bErrLine = bErrLine + 1
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(X, "p_Pivot" & bErrLine)
  
End Sub

Sub p_ZoomOut(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

 Call p_CheckConnection
  
 Call p_setExcelCalcOff
      X = HypMenuVZoomOut()
 Call p_setExcelCalcOn
    
    If X <> 0 Then GoTo ErrorHandler
 
  X = HypShowPov(False)
   isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:
'Call p_ErrorHandler(x, "p_ZoomOut")
   X = 0
End Sub

Sub p_ZoomIn(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

 Call p_CheckConnection
 
  Call p_setExcelCalcOff
         X = HypMenuVZoomIn()
  Call p_setExcelCalcOn
  
     If X <> 0 Then GoTo ErrorHandler
  X = HypShowPov(False)
   isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:
'Call p_ErrorHandler(x, "p_ZoomIn")
 X = 0
End Sub

Sub p_KeepOnly(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
 Call p_CheckConnection
 
  Call p_setExcelCalcOff
      X = HypMenuVKeepOnly()
   Call p_setExcelCalcOn
     
     If X <> 0 Then GoTo ErrorHandler
  X = HypShowPov(False)
   isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:
'Call p_ErrorHandler(x, "p_KeepOnly")
 X = 0
End Sub

Sub p_RemoveOnly(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

 Call p_CheckConnection
 
Call p_setExcelCalcOff
 X = HypMenuVRemoveOnly()

Call p_setExcelCalcOn
     If X <> 0 Then GoTo ErrorHandler
  X = HypShowPov(False)
  isMDXSlice = False
  
l_exit:
    Exit Sub
ErrorHandler:
'Call p_ErrorHandler(x, "p_RemoveOnly")
   X = 0
End Sub

Sub p_MemberSelect(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

 
 Call p_CheckConnection
     
     X = HypExecuteMenu(Empty, "Essbase->Member Selection")
     If X <> 0 Then GoTo ErrorHandler
     

  isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:

If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
 Else
  ' Call p_ErrorHandler(x, "p_MemberSelect")
 End If
 
End Sub
Sub p_CellComments(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
 
  If Not HypConnected(Empty) Then
      Call p_CheckConnection
  End If
  
     X = HypExecuteMenu(Empty, "Essbase->Linked Objects")
     If X <> 0 Then GoTo ErrorHandler
     
  isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:

If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
End If
 
End Sub

Sub p_Attributes(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
 
  If Not HypConnected(Empty) Then
      Call p_CheckConnection
  End If
  
     X = HypExecuteMenu(Empty, "Essbase->Insert Attributes")
     If X <> 0 Then GoTo ErrorHandler
     
  isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:

If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
End If
 
End Sub

Sub p_QueryDesigner(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
 
  If Not HypConnected(Empty) Then
      Call p_CheckConnection
  End If
  
     X = HypExecuteMenu(Empty, "Essbase->Query Designer")
     If X <> 0 Then GoTo ErrorHandler
     

  isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:

  If (X = -15) And InStr(UCase(ActiveSheet.Name), "QUERY") Then
      'MsgBox "You can't connect from Query page. Please use other sheets", vbExclamation
       X = 0
    End
  End If


If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
End If
 
End Sub

Sub p_Retrieve(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler


 'Call p_CheckConnection
 
   If Not HypConnected(Empty) Then
      Call p_CheckConnection
  End If
  
 Call p_RefreshRibbonNow
 
 
   
    Call p_setExcelCalcOff
        
        X = HypShowPov(False)
        
     If vIsFirstRetrive Then
      If vIsUseNameDefault Then
       X = HypSetAliasTable(Empty, "Default")
      Else
       X = HypSetAliasTable(Empty, "none")
      End If
       vIsFirstRetrive = False
     Else
        X = HypMenuVRefresh()
       ' x = HypRetrieve(ActiveSheet.Name)
     End If
         If X <> 0 Then GoTo ErrorHandler
        X = HypShowPov(False)
   Call p_setExcelCalcOn
     If X <> 0 Then GoTo ErrorHandler

   isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:
'Call p_ErrorHandler(x, "p_Retrieve")
 X = 0
  Call p_setExcelCalcOn
End Sub

Sub p_Undo(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
   
  Call p_CheckConnection
 
Call p_setExcelCalcOff
       X = HypMenuVUndo()
Call p_setExcelCalcOn

     If X <> 0 Then GoTo ErrorHandler
 
  isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:
'Call p_ErrorHandler(x, "p_Undo")
    X = 0
End Sub

Sub p_Redo(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
     
 Call p_CheckConnection
Call p_setExcelCalcOff

      X = HypMenuVRedo()
Call p_setExcelCalcOn

     If X <> 0 Then GoTo ErrorHandler
 
  isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:
'Call p_ErrorHandler(x, "p_Redo")
X = 0
End Sub
'

Sub p_SplitReports(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
     

 Call p_CheckConnection
  
   Call p_setExcelCalcOff
    X = HypExecuteMenu(Empty, "Essbase->Visualize in Excel")
  Call p_setExcelCalcOn

l_exit:
    Exit Sub
ErrorHandler:
If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
 Else
  ' Call p_ErrorHandler(x, "p_CalculationEssBase")
  X = 0
 End If
 
End Sub

Sub p_CalculationEssBase(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next

 Dim ObjectWsh As Object
Set ObjectWsh = VBA.CreateObject("WScript.Shell")
Dim waitOnReturn As Boolean: waitOnReturn = False
Dim windowStyle  As Integer: windowStyle = 1

Dim vCurrConectionNameOnSheet As String
Dim vPassword As String
Dim vMenuId
 
 
vCurrConectionNameOnSheet = getCurrSheetConnectionName()
vMenuId = getConnectionID(vCurrConectionNameOnSheet)

  vPassword = f_XORDecryption(fArrQuickConnections(vMenuId, 4), f_XOREncryption(vCurrPasswordLine, VBA.Environ("Computername") & VBA.Environ("Username")))
 
Dim vConnSTR
'"d:\Users\Rasyukeg\AppData\Roaming\Microsoft\AddIns\EssBaseWF.hta"  "conn=rulyubimir'rulyubimir'http://wedcb786.frmon.danet:13080/aps/SmartView'wedcb785.frmon.danet:1424'Kz1pnl'Kz1pnl|rulyubimir'rulyubimir'http://wedcb786.frmon.danet:13080/aps/SmartView'wedcb785.frmon.danet:1424'Kz1ttcst'Kz1ttcst"
Dim vStrCMD

 vStrCMD = "%userprofile%\AppData\Roaming\Microsoft\AddIns\EssBaseWF.hta"
  vStrCMD = vStrCMD & " "" conn="
  
  Dim i

 For i = 0 To UBound(fArrQuickConnections)
   vConnSTR = Split(fArrQuickConnections(i, 0), "'")
   If (InStr(fArrQuickConnections(vMenuId, 5), fArrQuickConnections(i, 5)) > 0) Then
        vStrCMD = vStrCMD & fArrQuickConnections(i, 3)
         vPassword = f_XORDecryption(fArrQuickConnections(i, 4), f_XOREncryption(vCurrPasswordLine, VBA.Environ("Computername") & VBA.Environ("Username")))
        vStrCMD = vStrCMD & Chr(39) & vPassword
        vStrCMD = vStrCMD & Chr(39) & "http://" & vConnSTR(4) & ":13080/aps/SmartView" 'aps
        vStrCMD = vStrCMD & Chr(39) & fArrQuickConnections(i, 5)  'esb
        vStrCMD = vStrCMD & Chr(39) & fArrQuickConnections(i, 1)  'app
        vStrCMD = vStrCMD & Chr(39) & fArrQuickConnections(i, 2) & "|" ' db
   End If
 Next

  vStrCMD = vStrCMD & """"
  
  Dim vEssBaseCalc
 If (isTextBoxPresent("EssBaseCalc")) Then
    vStrCMD = vStrCMD & " "" csc="
    vEssBaseCalc = getTextBoxValue("EssBaseCalc")
    vEssBaseCalc = getStringWOComments(vEssBaseCalc)
    vEssBaseCalc = Replace(vEssBaseCalc, ";", "|")
    vStrCMD = vStrCMD & vEssBaseCalc
    vStrCMD = vStrCMD & """"
 End If

 
  

           
  ObjectWsh.Run vStrCMD, windowStyle, waitOnReturn
   
  Set ObjectWsh = Nothing
  
l_exit:
    Exit Sub
ErrorHandler:
 
 Call p_ErrorHandler(0, "p_CalculationEssBase" & Err.Number & Err.Description & vStrCMD)
     
End Sub

Sub p_CalculationEssBaseOld(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

 Call p_CheckConnection
 
    
    X = HypExecuteMenu(Empty, "Essbase->Calculate")

l_exit:
    Exit Sub
ErrorHandler:
If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
 Else
  ' Call p_ErrorHandler(x, "p_CalculationEssBase")
  X = 0
 End If
 
     
End Sub


 
Sub p_SubmitData(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

 Call p_CheckConnection
 
  
   X = MsgBox(" Upload data ?", vbOKCancel, "Essbase Save Data")
       If X = 1 Then
         
         Call p_setExcelCalcOff
            X = HypMenuVSubmitData() 'HypExecuteMenu(ActiveSheet.Name, "Essbase->Submit Data") ' HypSubmitData(Empty) ' 'HypMenuVSubmitData()
         Call p_setExcelCalcOn
       Else
         X = 0
       End If
       If X <> 0 Then GoTo ErrorHandler
  
  isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:

 ' Call p_ErrorHandler(x, "p_SubmitData")
 
End Sub
 
Sub p_SubmitDataVORefresh(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

 Call p_CheckConnection
 
  
   X = MsgBox(" Upload data ?", vbOKCancel, "Essbase Save Data")
       If X = 1 Then
         Call p_setExcelCalcOff
            X = HypSubmitSelectedRangeWithoutRefresh(Null, False, True, True)
         Call p_setExcelCalcOn
       Else
         X = 0
       End If
       If X <> 0 Then GoTo ErrorHandler
  
  isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:

 
End Sub

Sub p_Export(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

Dim VBComp, vStrCMD

 vStrCMD = "%userprofile%\AppData\Roaming\Microsoft\AddIns\"

 

l_exit:
    Exit Sub
ErrorHandler:

  Call p_ErrorHandler(X, "p_SubmitData")
End Sub

Sub p_About(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

    MsgBox "Danone.Essbase  v3 2017.09.03" & vbNewLine _
          & vbNewLine & " developer: er@essbase.ru "
   ActiveWorkbook.FollowHyperlink Address:="https://danone.facebook.com/groups/315738245470777/", NewWindow:=True
   ActiveWorkbook.FollowHyperlink Address:="https://github.com/er77/EssBase.ACT/issues", NewWindow:=True
   
     X = HypMenuVAbout
 
l_exit:
    Exit Sub
ErrorHandler:
  Call p_ErrorHandler(0, "p_About")
End Sub


               

