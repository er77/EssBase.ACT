Attribute VB_Name = "tTest"
Option Explicit

Public Sub p_test(vIRibbonControl As IRibbonControl)
  On Error GoTo ErrorHandler
  
 
 Dim vString

   X = HypGetCalcScript(Null, "ACT_CALC_ALL", 1, vString)

   ActiveSheet.Cells(1, 1).value = vString
 

 
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(X, "p_test")
End Sub
