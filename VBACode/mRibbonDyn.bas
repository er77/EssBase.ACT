Attribute VB_Name = "mRibbonDyn"
Option Explicit

Public vCurrXMLGlobal

Function isSheetOTL() As Boolean
isSheetOTL = False
  If (InStr(UCase(ActiveSheet.Name), "OTL") > 0) Then
   isSheetOTL = True
  End If
  
End Function


Function isConnectPresent() As Boolean
 isConnectPresent = isTextBoxPresent("ConnectQ")
End Function

 Sub p_getEnabled(ByVal vIRibbonControl As IRibbonControl, ByRef vReturnValue)
  vReturnValue = isConnectPresent
 End Sub
 
  Sub p_IsVisible(ByVal vIRibbonControl As IRibbonControl, ByRef vReturnValue)
  
    vReturnValue = False
  
  If vModeAnalyse = 0 Then
    Select Case vIRibbonControl.ID
        Case "grp_RData"
           vReturnValue = True
        Case "b_SheetInfo"
            vReturnValue = True
        Case "grp_Options"
            vReturnValue = True
        Case "grp_Main0"
            vReturnValue = True
        Case "grp_Refresh"
          vReturnValue = True
     End Select
  End If
     
 End Sub

Private Function f_makeEssConnectMenu() As String
  Dim vCurrXML  As String
  Dim vCurrSvr  As String
  Dim vCurrName As String
  Dim vCurrServer() As String
  Dim vCurrServerSTR As String
  Dim ISHugeList  As Boolean
  Dim i As Integer
  Dim J As Integer
  Dim vMCount As Integer
  
   If vCurrXMLGlobal = "" Then
 
    vCurrXML = ""
    vCurrSvr = ""
    vMCount = 0
         ' DoEvents
        On Error Resume Next
        ISHugeList = False
       For i = 0 To UBound(fArrQuickConnections, 1)
            If Not fArrQuickConnections(i, 0) = "" Then
            
             vCurrServer() = Split(fArrQuickConnections(i, 5), ".")
             vCurrServerSTR = LCase(vCurrServer(0)) & "." & UCase(Left(fArrQuickConnections(i, 1), 3))
             
             If (vCurrSvr <> vCurrServerSTR) Then
               If vCurrSvr = "" Then
                  vCurrSvr = vCurrServerSTR
               Else
                  ISHugeList = True
                  Exit For
               End If
              End If
            End If
        Next
        
    vCurrXML = ""
    vCurrSvr = ""
    vMCount = 0
     
         ' DoEvents
        For i = 0 To UBound(fArrQuickConnections, 1)
        
            If Not fArrQuickConnections(i, 0) = "" Then
            vMCount = 1
            
           
             vCurrServer() = Split(fArrQuickConnections(i, 5), ".")
             vCurrServerSTR = LCase(vCurrServer(0)) & "." & UCase(Left(fArrQuickConnections(i, 1), 3))
              If ISHugeList Then
                If (UCase(vCurrSvr) <> UCase(vCurrServerSTR)) Then
                 If vCurrSvr <> "" Then
                     vCurrXML = vCurrXML & "</menu>"
                 End If
                    vCurrSvr = vCurrServerSTR
                    vCurrXML = vCurrXML & "<menu id=""" & vCurrSvr & """ label=""" & vCurrSvr & """    insertBeforeMso=""Cut""  imageMso=""ExportMoreMenu"">"
                End If
              End If
            vCurrName = LCase(fArrQuickConnections(i, 1)) & "." & Chr(9) & LCase(fArrQuickConnections(i, 2))
            
             If ISHugeList Then
            
             
               vCurrXML = vCurrXML & "<button id=""b_qConnect" & i & """ label=""" & vCurrName & """ " _
                                   & "onAction=""p_qConnectAction"" imageMso=""DatabasePermissionsMenu""   />" '
                vCurrXML = vCurrXML & "<menuSeparator id=""b_qConnectSep" & i & """  />"
             Else
                 
                 vCurrName = UCase(vCurrName)
                 vCurrXML = vCurrXML & "<button id=""b_qConnect" & i & """ label=""" & vCurrName & """ " _
                                     & "onAction=""p_qConnectAction"" imageMso=""DatabasePermissionsMenu""   />" '
                 vCurrXML = vCurrXML & "<menuSeparator id=""b_qConnectSep" & i & """  />"
             
             End If
            End If
        Next
       
       On Error GoTo ErrorHandler
       If vMCount = 1 And ISHugeList Then
         vCurrXML = vCurrXML & "</menu>"
       End If
 
f_makeEssConnectMenu = vCurrXML
 vCurrXMLGlobal = vCurrXML
' ActiveSheet.Cells(1, 1).value = vCurrXMLGlobal

  Else
  f_makeEssConnectMenu = vCurrXMLGlobal
 ' ActiveSheet.Cells(1, 1).value = vCurrXMLGlobal
  End If
l_exit:
    Exit Function
ErrorHandler:
 Call p_ErrorHandler(0, "f_makeEssConnectMenu")
 
End Function

 
 Public Sub p_QuickConnect(vIRibbonControl As IRibbonControl, ByRef vXMLMenu)
 On Error GoTo ErrorHandler
    
    Dim vCurrXML As String
   Call p_setExcelCalcOff
   
  ActiveSheet.Cells(1, 1).Select

  p_checkConnectionsArray
   
 vCurrXML = "<menu xmlns=""" & _
           "http://schemas.microsoft.com/office/2006/01/customui"">" & vbCrLf
   


 vCurrXML = vCurrXML & " <button id=""b_LastConnect""  label=""ReConnect Sheet""    onAction=""p_LastConnectNew""     imageMso=""RecurrenceEdit"" /> "
  
     vCurrXML = vCurrXML & f_makeEssConnectMenu
  
   vCurrXML = vCurrXML & "<menuSeparator id=""b_qConnectSepZ3""  />"
 
 
    vCurrXML = vCurrXML & " <button id=""b_EditQConnect""  label=""Manage QC""    onAction=""p_EditQConnect""      imageMso=""AddOrRemoveAttendees"" /> "
 

    vCurrXML = vCurrXML & "</menu>"
    
    vXMLMenu = vCurrXML

   p_RefreshRibbonNow
   
l_exit:
    Exit Sub
ErrorHandler:
 If Err.Number = 91 Then
         MsgBox "Please open workbook or create new one ", vbExclamation
         End
   Else
     Call p_ErrorHandler(0, "p_QuickConnect")
 End If
 
Call p_setExcelCalcOn_INT

  
End Sub

 

 


