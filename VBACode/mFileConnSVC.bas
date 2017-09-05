Attribute VB_Name = "mFileConnSVC"
Option Explicit
Option Compare Text

Public vRibbonSetFileName As String
Public fArrQuickConnections() As String

Public fArrQCProd() As String

 
Sub GetEssbaseRibonnConnectionFileName()
 On Error Resume Next
 
Dim objFolders As Object
Set objFolders = CreateObject("WScript.Shell").SpecialFolders
 
       vRibbonSetFileName = objFolders("mydocuments") & "\essribon.cfg"
     

' MsgBox vRibbonSetFileName
  If (Dir(vRibbonSetFileName) = "") Then
    Open vRibbonSetFileName For Output As #1
        Write #1, ""
        Close #1
  End If
 SetAttr vRibbonSetFileName, vbNormal
 
 Set objFolders = Nothing
 
 If Err.Number <> 0 Then
   Err.Clear
 End If
 
End Sub
 


Public Sub f_DeleteLineFromCfg(vDeletedString As String)
 On Error GoTo ErrorHandler

    Dim vArrOfStrings() As String, vCurrStr As String
    Dim i As Long, J As Long
    
    Call GetEssbaseRibonnConnectionFileName
      SetAttr vRibbonSetFileName, vbNormal
    Open vRibbonSetFileName For Input As 1
    i = 0
    J = 0
     'DoEvents
    Do Until EOF(1)
        J = J + 1
        Line Input #1, vCurrStr
         Dim vCurrArrayLine() As String
          vCurrArrayLine() = Split(vCurrStr, "|")
        If (UBound(vCurrArrayLine) > 3) Then
            If (InStr(CRC16HASH(vCurrArrayLine(0)), vDeletedString) = 0) Then
                i = i + 1
                ReDim Preserve vArrOfStrings(1 To i)
                vArrOfStrings(i) = vCurrStr
            End If
        End If
    Loop
    Close #1
    J = i
    
    'Write array to file
    Open vRibbonSetFileName For Output As 1
    
    For i = 1 To J
      'DoEvents
        Print #1, vArrOfStrings(i)
    Next i
    Close #1

l_exit:
    SetAttr vRibbonSetFileName, vbHidden
    Exit Sub
ErrorHandler:
  Call p_ErrorHandler(0, "f_DeleteLineFromCfg")
    
End Sub

 

 

Public Sub p_ReadCurrConnections()
 On Error GoTo ErrorHandler
     
    p_ReadCurrConnectionsINT
     
l_exit:
    Exit Sub
ErrorHandler:
 
   Err.Clear
End Sub


Public Sub p_ReadCurrConnectionsINT()
 On Error GoTo ErrorHandler

    Dim vCurrStr As String
    Dim fCurrConnections(200, 5) As String
    Dim i, J, q
    Dim vCurrArrayLine() As String
    Dim vTempArrLine(1, 5) As String
    Call GetEssbaseRibonnConnectionFileName
    
    SetAttr vRibbonSetFileName, vbNormal
         
    If Dir(vRibbonSetFileName) = "" Then
      Exit Sub
    End If
     Dim fCurrConnections2() As String
     
     Call SetAttr(vRibbonSetFileName, vbNormal)
     
    Open vRibbonSetFileName For Input As 1
    On Error GoTo ErrorHandler
    
    i = 0
     'DoEvents
       Dim k
        k = 0
        
    On Error Resume Next
    Do Until EOF(1)
        Line Input #1, vCurrStr
        vCurrArrayLine() = Split(vCurrStr, "|")
        If UBound(vCurrArrayLine) = 4 Then
           For J = 0 To 4
                fCurrConnections(i, J) = vCurrArrayLine(J)
            Next
          vCurrArrayLine() = Split(vCurrArrayLine(0), "'")
          fCurrConnections(i, 5) = vCurrArrayLine(2)
          i = i + 1
          k = k + 1
       End If
    Loop
    Close #1
     On Error GoTo ErrorHandler
     
    If k < 1 Then
      ReDim Preserve fCurrConnections2(0 To 0, 0 To 5)
      fArrQuickConnections = fCurrConnections2
        With New FileSystemObject
           If .FileExists(vRibbonSetFileName) Then
               .DeleteFile vRibbonSetFileName
           End If
        End With
        vCurrXMLGlobal = ""
     Exit Sub
    End If
    
    Call f_DeleteLineFromCfg("XXX$%^")  ' delete trash
  
    Call SetAttr(vRibbonSetFileName, vbHidden)
    
   ReDim Preserve fCurrConnections2(0 To k - 1, 0 To 5)
  J = 0
      For i = 0 To k
        If fCurrConnections(i, 0) <> "" Then
               For q = 0 To 5
                    fCurrConnections2(J, q) = fCurrConnections(i, q)
                Next
            J = J + 1
        End If
      Next
     
      'DoEvents
    ' buble sort
   
      For i = 0 To UBound(fCurrConnections2)
        For J = 0 To (UBound(fCurrConnections2) - 1)
         If (fCurrConnections2(J, 0) <> "") Then
           If (UCase(fCurrConnections2(J, 5) + fCurrConnections2(J, 1) + fCurrConnections2(J, 2)) > UCase(fCurrConnections2(J + 1, 5) + fCurrConnections2(J + 1, 1) + fCurrConnections2(J + 1, 2))) Then
                For q = 0 To 5
                    vTempArrLine(1, q) = fCurrConnections2(J, q)
                Next
                For q = 0 To 5
                    fCurrConnections2(J, q) = fCurrConnections2(J + 1, q)
                Next
                For q = 0 To 5
                    fCurrConnections2(J + 1, q) = vTempArrLine(1, q)
                Next
          End If
         End If
        Next
      Next
      
     fArrQuickConnections = fCurrConnections2
     
     vCurrXMLGlobal = ""
    
l_exit:
    Exit Sub
ErrorHandler:
' Call p_ErrorHandler(0, "p_ReadCurrConnectionsINT")
 ReDim Preserve fCurrConnections2(0 To 0, 0 To 5)
     fArrQuickConnections = fCurrConnections2
      vCurrXMLGlobal = ""
      Close #1
     Err.Clear
End Sub

 




 


