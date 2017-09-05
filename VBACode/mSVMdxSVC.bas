Attribute VB_Name = "mSVMdxSVC"
Option Explicit
 

Sub p_clearUsedRange()
On Error Resume Next
ActiveSheet.Cells(1, 1).Select
 ActiveSheet.UsedRange.Clear
    If Err.Number > 0 Then
         Err.Clear
    End If
End Sub
 
 Sub p_ExecStoredMDXINT()

On Error GoTo ErrorHandler

Dim vCurrMDX As String
Dim vCurrDBString

ActiveSheet.Cells(1, 1).Select

 
   Dim vExecutedMDX

   Dim DataObj As New msforms.DataObject
   
   vCurrMDX = getTextBoxValue("MDXq")
   vExecutedMDX = getMdxWithOutExcelVaribles(vCurrMDX)
    DataObj.SetText vExecutedMDX
    DataObj.PutInClipboard
   
   Dim vCurrArrayLine() As String
   vCurrArrayLine() = Split(vExecutedMDX, "FROM")
   vCurrArrayLine() = Split(vCurrArrayLine(1), "/")
    vCurrArrayLine(0) = getClearString(vCurrArrayLine(0))
   
  '/* CONNECT TO VSVR=%VSVR% */
  
    Dim vSVRArr() As String
    Dim vSVR As String
     vSVRArr() = Split(vExecutedMDX, "VSVR=")
     
   Dim J
    On Error Resume Next
      J = UBound(vSVRArr)
     If Err.Number <> 0 Then
       vSVR = ""
      Err.Clear
      Else
      vSVR = getStringWOComments(vSVRArr(1))
     End If
   On Error GoTo ErrorHandler
   
   Call p_connnectByServerAndDatabaseName(vCurrArrayLine(0), vSVR)
   Call p_CheckConnection
 
 
        isMDXSlice = True
            Call p_clearUsedRange
            Call p_setExcelCalcOff
            
  If vIsUseNameDefault Then
    X = HypSetAliasTable(Empty, "Default")
  Else
   X = HypSetAliasTable(Empty, "None")
  End If
  
   X = HypExecuteQuery(Empty, vExecutedMDX)
   
    If X <> 0 Then GoTo ErrorHandler
   X = HypShowPov(False)
  
   ActiveSheet.Cells.EntireRow.Hidden = False
    
  Call p_deleteAllTextBox("MDXq")
  Call p_CreateTextBox("MDXq", "" & vCurrMDX)
  Call hideTextBox("MDXq")
  
  Call p_setExcelCalcOn

l_exit:
    Exit Sub
    
ErrorHandler:

 If X = -15 Then
   MsgBox "There is some errors happens . Probably it is because MDX syntax error. Now starts Syntax checker", vbExclamation
   Call p_ExecuteMdxfromMenuINT
 Else
   If X = 10000 Then
     MsgBox "There are no data found. Check suppress missing clause in the MDX statement", vbExclamation
  Else
   If X = -4 Then
     MsgBox "The connection have retaired. Please make new one", vbExclamation
    End
   End If
    Call p_ErrorHandler(X, " p_ExecStoredMDX " & vExecutedMDX)
     End If
 End If
 
  Call p_deleteTextBox("MDXq") ' delete if MDXq is more than 1
  Call p_setExcelCalcOn

 
  End
 

End Sub
 
 
 

Sub p_ExecStoredMDX(vIRibbonControl As IRibbonControl)

 On Error GoTo ErrorHandler

isMDXSlice = True
ActiveSheet.Cells(1, 1).Select
 
   Call fExecuteMDX.Show(1)
  ' Call p_freezePanel
  ' Call p_RefreshRibbonNow
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, "p_ExecuteMdxNative")
End Sub

 

Sub p_ExecuteMdxfromMenu(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

  isMDXSlice = True
  
 
 
 Call p_CheckConnection
 
 Call p_ExecuteMdxfromMenuINT
 
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, "p_ExecuteMdxfromMenu")
 
End Sub

Sub p_ExecuteMdxfromMenuINT()
On Error GoTo ErrorHandler

 
 
    Call p_setExcelCalcOff
    
       X = HypExecuteMenu(Empty, "Essbase->Execute Mdx")
      ' Call p_freezePanel
       
    Call p_setExcelCalcOn
   
     If X <> 0 Then GoTo ErrorHandler
 
  
  X = HypShowPov(False)
   isMDXSlice = False


l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, "p_ExecuteMdxfromMenu")
 
End Sub



Function getMdxWithOutExcelVaribles(ByVal vCurrMDX As String)
On Error GoTo ErrorHandler
         Dim vWS_Count As Integer
         Dim i, J
         Dim vCurrSheetName, vCurrTextBox, vResultMdx, vCurVarName
         Dim vArrCurrTexBox, vArrLine
      
         vWS_Count = ActiveWorkbook.Worksheets.Count
         
         If InStr(vCurrMDX, "%") < 1 Then
           getMdxWithOutExcelVaribles = vCurrMDX
          Exit Function
         End If
         
         vResultMdx = UCase(vCurrMDX)
         vResultMdx = Replace(vResultMdx, UCase("CurrSheet"), ActiveSheet.Name)
         ' Begin the loop.
         For i = 1 To vWS_Count
            vCurrSheetName = ActiveWorkbook.Worksheets(i).Name
            If InStr(UCase(vCurrMDX), "%" & UCase(vCurrSheetName) & ".") > 0 Then
              If Not isTextBoxPresentOnSheetID("MDXVaribales", i + 0) Then
                MsgBox "Current MDX Query have sheet link " & vCurrSheetName & " , but on this page have not any textBox ~MDXVaribales~ ", vbExclamation
               End
              End If
              vCurrTextBox = getTextBoxValueOnSheetID("MDXVaribales", i + 0)
              vCurrTextBox = getClearString(vCurrTextBox)
               
              ' loop for text lines
               vArrCurrTexBox = Split(vCurrTextBox, ";")

               For J = 0 To UBound(vArrCurrTexBox)
                If Len(vArrCurrTexBox(J) > 3) Then
                   vArrLine = Split(getStringWOComments(vArrCurrTexBox(J)), "=")
                    If (UBound(vArrLine) = 1) Then
                       vCurVarName = UCase("%" & vCurrSheetName & "." & vArrLine(0) & "%")
                       vResultMdx = Replace(vResultMdx, vCurVarName, Trim(vArrLine(1)))
                    End If
                 End If
                Next
            End If
            
         Next i
         getMdxWithOutExcelVaribles = vResultMdx
l_exit:
    Exit Function
ErrorHandler:
Call p_ErrorHandler(0, "getMdxWithOutExcelVaribles")
End Function
