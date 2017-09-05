Attribute VB_Name = "mSheetRepair"
Option Explicit


Function ConvertToLetter(ByVal iCol As Integer) As String
 On Error GoTo ErrorHandler
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
     Exit Function
ErrorHandler:
   Call p_setExcelCalcOn
   Call p_ErrorHandler(X, "ConvertToLetter" & Err.Number & Err.Description & Err.HelpContext)
   End
End Function

Function DoubleChar(ByVal iCol As Integer) As String
    DoubleChar = "." & LCase(ConvertToLetter(iCol)) & Int(iCol * Rnd + iCol)
End Function

Function isCheckName(ByVal vSheetName As String) As Boolean
 On Error GoTo ErrorHandler
     Dim vWS_Count, i
         vWS_Count = ActiveWorkbook.Worksheets.Count
         isCheckName = True
         For i = 1 To vWS_Count
            If InStr(ActiveWorkbook.Worksheets(i).Name, vSheetName) > 0 Then
                  isCheckName = False
            End If
         Next i
     Exit Function
ErrorHandler:
   Call p_setExcelCalcOn
   Call p_ErrorHandler(X, "isCheckName" & Err.Number & Err.Description & Err.HelpContext)
   End
End Function

Function NewName() As String
 On Error GoTo ErrorHandler
     Dim vArrName() As String
     Dim vNumber
     Dim vSec
      vArrName() = Split(ActiveSheet.Name, ".")
      vArrName() = Split(vArrName(0), " ")
      vArrName() = Split(vArrName(0), ",")
      
       Dim isNameCorrect
       isNameCorrect = False
       vNumber = 1
      While Not isNameCorrect
        vNumber = ActiveWorkbook.Worksheets.Count * vNumber
        vSec = Second(Now)
        vNumber = vNumber - (Round(vNumber / vSec))   '(vNumber * vSec) Mod vNumber
        vNumber = vNumber * (vNumber * Second(Now))
        vNumber = vNumber - (Round(vNumber / vSec))
        vNumber = Int(vNumber - (Round(vNumber / 100)))
        vNumber = vSec + vNumber - 100 * Int(vNumber / 100)
        vNumber = Abs(vNumber)
        vNumber = vNumber * (vNumber * vSec - vSec * Int(vNumber * vSec / (vSec * 10)))
        vNumber = Int(vNumber * vSec - 100 * Int(vNumber * vSec / 100))
        vNumber = Left("" & vNumber * vSec, 3) * Right("" & vNumber * Second(Now), 3) * Right("" & vNumber * Second(Now), 1) * Left("" & vNumber * Second(Now), 1)
        vNumber = Left("" & vNumber * Second(Now), 1) * Right("" & vNumber * Second(Now), 1) + Right("" & vNumber * Second(Now), 2) + Left("" & vNumber * Second(Now), 2)
        vNumber = Left("" & vNumber * Second(Now), 2)
        If (vNumber > 10) Then
            NewName = Left(vArrName(0), 3) & DoubleChar(vNumber)
            isNameCorrect = isCheckName(NewName)
        End If
      Wend
     Exit Function
ErrorHandler:
   Call p_setExcelCalcOn
   Call p_ErrorHandler(X, "NewName" & Err.Number & Err.Description & Err.HelpContext)
   End
End Function

Sub p_CreateSheet(vOldSheetName As String, vNewSheetName As String)
 On Error GoTo ErrorHandler
   ActiveSheet.Cells(1, 1).Select
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = vNewSheetName
    Worksheets(vNewSheetName).Move _
       after:=Worksheets(vOldSheetName)
       
l_exit:
    Exit Sub
ErrorHandler:

   Call p_setExcelCalcOn
   Call p_ErrorHandler(X, "p_CreateSheet" & Err.Number & Err.Description & Err.HelpContext)
   End
End Sub
 
Sub p_CopySheet(vOldSheetName As String, vNewSheetName As String)
 On Error GoTo ErrorHandler
    Dim vUsedRange As Range
    ActiveSheet.Cells(1, 1).Select
    Worksheets(vOldSheetName).UsedRange.Copy
    Worksheets(vNewSheetName).Range("A1").PasteSpecial xlPasteValues
    Worksheets(vNewSheetName).Range("A1").PasteSpecial xlPasteFormulas
l_exit:
    Exit Sub
ErrorHandler:

   Call p_setExcelCalcOn
   Call p_ErrorHandler(X, "p_CopySheet " & Err.Number & Err.Description & Err.HelpContext)
   End
End Sub

Sub p_copyMDXcommnads(vOldSheetName As String, vNewSheetName As String)
 On Error GoTo ErrorHandler
        Dim vCurrMDX As String
   Dim vCurrConnectQ As String
   Dim vVariableBox As String
   
   Worksheets(vOldSheetName).Activate
      vCurrMDX = getTextBoxValue("MDXq")
 vCurrConnectQ = getTextBoxValue("ConnectQ")
 vVariableBox = getTextBoxValue("MDXVaribales")
  
    Worksheets(vNewSheetName).Activate
    
    If (Len(vCurrMDX) > 0) Then
     Call p_CreateTextBox("MDXq", vCurrMDX)
    End If
    
    If (Len(vVariableBox) > 0) Then
    Call p_CreateTextBox("MDXVaribales", vVariableBox)
    End If
    
    If (Len(vCurrConnectQ) > 0) Then
      Call p_CreateTextBox("ConnectQ", vCurrConnectQ)
    End If
    
l_exit:
    Exit Sub
ErrorHandler:

   Call p_setExcelCalcOn
   Call p_ErrorHandler(X, "p_copyMDXcommnads " & Err.Number & Err.Description & Err.HelpContext)
   End
End Sub

Public Sub p_CopySheetMain(vOldSheetName As String, vNewSheetName As String)
   
 On Error GoTo ErrorHandler
 
   Call p_CreateSheet(vOldSheetName, vNewSheetName)
   Call p_CopySheet(vOldSheetName, vNewSheetName)
   Call p_copyMDXcommnads(vOldSheetName, vNewSheetName)
  
   
l_exit:
    Exit Sub
ErrorHandler:

   Call p_setExcelCalcOn
   Call p_ErrorHandler(X, "p_CopySheetMain " & Err.Number & Err.Description & Err.HelpContext)
   
   End
End Sub

Public Sub p_RenewSheet()

  If ActiveSheet Is Nothing Then
        MsgBox "active sheet is not determinated "
        End
    End If
    
  ActiveSheet.Cells(1, 1).Select
    
On Error GoTo ErrorHandler
    
   Call p_setExcelCalcOff
   Dim vOldSheetName As String
   Dim vNewSheetName As String
   vOldSheetName = ActiveSheet.Name
   vNewSheetName = NewName()
   
   Call p_CopySheetMain(vOldSheetName, vNewSheetName)
   
    Application.DisplayAlerts = False
    Worksheets(vOldSheetName).Delete
    Application.DisplayAlerts = True
    
     Worksheets(vNewSheetName).Activate
     ActiveSheet.Name = vOldSheetName

   Call p_setExcelCalcOn
   
l_exit:
    Exit Sub
ErrorHandler:

   Call p_setExcelCalcOn
   Call p_ErrorHandler(X, "p_RenewSheet " & Err.Number & Err.Description & Err.HelpContext)
   End
    
End Sub

Public Sub p_RepairRetrive(vIRibbonControl As IRibbonControl)

   Call p_RenewSheet
   
l_exit:
    Exit Sub
ErrorHandler:

   Call p_setExcelCalcOn
   Call p_ErrorHandler(X, "p_RepairRetrive " & Err.Number & Err.Description & Err.HelpContext)
   
   End
End Sub

Public Sub p_copySheetUI(vIRibbonControl As IRibbonControl)

  If ActiveSheet Is Nothing Then
        MsgBox "active sheet is not determinated "
        End
    End If
    
  ActiveSheet.Cells(1, 1).Select
    
On Error GoTo ErrorHandler
    
   Call p_setExcelCalcOff
   
   Call p_CopySheetMain(ActiveSheet.Name, NewName())
   Call p_setExcelCalcOn
   
  ActiveSheet.Cells(1, 1).Select
   
l_exit:
    Exit Sub
ErrorHandler:

   Call p_setExcelCalcOn
   Call p_ErrorHandler(X, "p_copySheetUI " & Err.Number & Err.Description & Err.HelpContext)
   
   End
End Sub
