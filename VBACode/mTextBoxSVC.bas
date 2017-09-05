Attribute VB_Name = "mTextBoxSVC"
Option Explicit

Sub p_checkMultiplySheets()
 ActiveSheet.Select
End Sub

Function getClearString(ByVal vCurrStr As String)
  Dim i As Long, vCodesToClean As Variant
  vCodesToClean = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
                       21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 96, 126, 127, 127, 129, 141, 143, 144, 157, 160)
  For i = LBound(vCodesToClean) To UBound(vCodesToClean)
    If InStr(vCurrStr, Chr(vCodesToClean(i))) Then vCurrStr = Replace(vCurrStr, Chr(vCodesToClean(i)), "")
  Next
  
  For i = 128 To 255
    If InStr(vCurrStr, Chr(i)) Then vCurrStr = Replace(vCurrStr, Chr(i), "")
  Next
    vCurrStr = Application.WorksheetFunction.Clean(vCurrStr)
    getClearString = Trim(vCurrStr)
End Function

Function getStringWOComments(ByVal vCurrStr As String)
        Dim strPattern As String: strPattern = "[^a-zA-Z0-9=;.&_]" 'The regex pattern to find special characters
        Dim strReplace As String: strReplace = "" 'The replacement for the special characters
        Dim regEx
        Set regEx = CreateObject("vbscript.regexp") 'Initialize the regex object
        ' Configure the regex object
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        ' Perform the regex replacement
        getStringWOComments = regEx.Replace(vCurrStr, strReplace)
End Function


Sub p_CreateTextBox(vNameOfTextBox As String, vText As String)
On Error Resume Next
Call p_checkMultiplySheets

Call p_deleteAllTextBox(vNameOfTextBox)
  ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 200, 50).Name = vNameOfTextBox
  ActiveSheet.Shapes(vNameOfTextBox).TextFrame.Characters.Text = vText
Call hideTextBox(vNameOfTextBox)

   If Err.Number > 0 Then
         Err.Clear
    End If
End Sub

 

Sub p_deleteAllTextBox(vNameOfTextBox As String)  ' delete All
On Error Resume Next
Dim oTextBox As TextBox
 
For Each oTextBox In ActiveSheet.TextBoxes

  If InStr(UCase(oTextBox.Name), UCase(vNameOfTextBox)) > 0 Then
    oTextBox.Delete
  End If
Next oTextBox
Set oTextBox = Nothing
'DoEvents
   If Err.Number > 0 Then
         Err.Clear
    End If
End Sub

Sub p_deleteTextBox(vNameOfTextBox As String)  ' delete if MDXq is more than 1

Dim oTextBox As TextBox
Dim i


Call p_checkMultiplySheets
i = 0

For Each oTextBox In ActiveSheet.TextBoxes

  If InStr(UCase(oTextBox.Name), UCase(vNameOfTextBox)) > 0 Then
   i = i + 1
  End If
  
  If i > 1 Then
    oTextBox.Delete
  End If
  
Next oTextBox
Set oTextBox = Nothing
'DoEvents
End Sub

 
Public Function isTextBoxPresent(vNameOfTextBox As String) As Boolean
On Error Resume Next

    isTextBoxPresent = False

    isTextBoxPresent = (Len(Trim(ActiveSheet.Shapes(vNameOfTextBox).TextFrame.Characters.Text)) > 0)
    If Err.Number > 0 Then
         Err.Clear
    End If
End Function
Public Function isTextBoxPresentOnSheetID(vNameOfTextBox As String, vSheetId As Integer) As Boolean
On Error Resume Next
    isTextBoxPresentOnSheetID = False

    isTextBoxPresentOnSheetID = (Len(Trim(ActiveWorkbook.Worksheets(vSheetId).Shapes(vNameOfTextBox).TextFrame.Characters.Text)) > 0)
    If Err.Number > 0 Then
         Err.Clear
    End If
End Function

Function getTextBoxValue(vNameOfTextBox As String) As String
On Error Resume Next
getTextBoxValue = ""
    If isTextBoxPresent(vNameOfTextBox) Then
       getTextBoxValue = UCase(ActiveSheet.Shapes(vNameOfTextBox).TextFrame.Characters.Text)
    End If
 Call p_deleteTextBox(vNameOfTextBox) ' delete if vNameOfTextBox  is more than 1
 Call hideTextBox(vNameOfTextBox)
    If Err.Number > 0 Then
         Err.Clear
    End If
End Function

Function getTextBoxValueOnSheetID(vNameOfTextBox As String, vSheetId As Integer) As String
On Error Resume Next
getTextBoxValueOnSheetID = ""

    If isTextBoxPresentOnSheetID(vNameOfTextBox, vSheetId) Then
       getTextBoxValueOnSheetID = UCase(ActiveWorkbook.Worksheets(vSheetId).Shapes(vNameOfTextBox).TextFrame.Characters.Text)
    End If
    
    If Err.Number > 0 Then
         Err.Clear
    End If
    
End Function


Sub hideTextBox(vNameOfTextBox As String)
Dim vIsMyBox
On Error Resume Next
vIsMyBox = False


   If InStr(UCase(vNameOfTextBox), UCase("ConnectQ")) > 0 Then
     vIsMyBox = True
    End If
     
     If InStr(UCase(vNameOfTextBox), UCase("CalcQ")) > 0 Then
       vIsMyBox = True
     End If
     
     If InStr(UCase(vNameOfTextBox), UCase("SqlQ")) > 0 Then
        vIsMyBox = True
     End If
     
    If InStr(UCase(vNameOfTextBox), UCase("MDXq")) > 0 Then
        vIsMyBox = True
     End If
     
    If vIsMyBox Then
        ActiveSheet.Shapes(vNameOfTextBox).Width = 1
        ActiveSheet.Shapes(vNameOfTextBox).Height = 1
        ActiveSheet.Shapes(vNameOfTextBox).Left = 50000
        ActiveSheet.Shapes(vNameOfTextBox).Top = 50000
     End If
     
    If Err.Number > 0 Then
         Err.Clear
    End If
End Sub

Sub p_HideTextBox(vIRibbonControl As IRibbonControl)
Dim oTextBox As TextBox
Dim i
Call p_checkMultiplySheets
i = 0

For Each oTextBox In ActiveSheet.TextBoxes
 
  hideTextBox (oTextBox.Name)
  
Next oTextBox
Set oTextBox = Nothing
End Sub

Sub p_ShowTextBox(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
Call p_checkMultiplySheets
Dim oTextBox As TextBox
Dim i

If Not (isTextBoxPresent("ConnectQ")) Then
  Call p_CreateTextBox("ConnectQ", "")
End If

If Not (isTextBoxPresent("MDXQ")) Then
  Call p_CreateTextBox("MDXQ", "")
End If

If Not (isTextBoxPresent("CalcQ")) Then
  Call p_CreateTextBox("CalcQ", "")
End If

If Not (isTextBoxPresent("SqlQ")) Then
  Call p_CreateTextBox("SqlQ", "")
End If

i = 0

For Each oTextBox In ActiveSheet.TextBoxes
 Call showTextBox(oTextBox.Name, i)
Next oTextBox



Set oTextBox = Nothing
End Sub

Sub showTextBox(vNameOfTextBox As String, ByVal i As Integer)

Call p_checkMultiplySheets
 On Error Resume Next
    '    ActiveSheet.Shapes(vNameOfTextBox).Width = 150
    '    ActiveSheet.Shapes(vNameOfTextBox).Height = 150
    '    ActiveSheet.Shapes(vNameOfTextBox).Left = 400
    '    ActiveSheet.Shapes(vNameOfTextBox).Top = 400
 
 
     If InStr(UCase(vNameOfTextBox), UCase("ConnectQ")) > 0 Then
        ActiveSheet.Shapes(vNameOfTextBox).Width = 100
        ActiveSheet.Shapes(vNameOfTextBox).Height = 20
        ActiveSheet.Shapes(vNameOfTextBox).Left = 10
        ActiveSheet.Shapes(vNameOfTextBox).Top = 10
      End If
      
     
     If InStr(UCase(vNameOfTextBox), UCase("CalcQ")) > 0 Then
        ActiveSheet.Shapes(vNameOfTextBox).Width = 150
        ActiveSheet.Shapes(vNameOfTextBox).Height = 150
       ActiveSheet.Shapes(vNameOfTextBox).Left = 70
        ActiveSheet.Shapes(vNameOfTextBox).Top = 70
     End If
     
     If InStr(UCase(vNameOfTextBox), UCase("SqlQ")) > 0 Then
        ActiveSheet.Shapes(vNameOfTextBox).Width = 150
        ActiveSheet.Shapes(vNameOfTextBox).Height = 150
        ActiveSheet.Shapes(vNameOfTextBox).Left = 100
        ActiveSheet.Shapes(vNameOfTextBox).Top = 100
     End If
     
    If InStr(UCase(vNameOfTextBox), UCase("MDXq")) > 0 Then
        ActiveSheet.Shapes(vNameOfTextBox).Width = 150
        ActiveSheet.Shapes(vNameOfTextBox).Height = 150
        ActiveSheet.Shapes(vNameOfTextBox).Left = 250
        ActiveSheet.Shapes(vNameOfTextBox).Top = 150
   With ActiveSheet.Shapes(vNameOfTextBox)
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
        .Solid
    End With

     End If
     If Err.Number > 0 Then
         Err.Clear
    End If
End Sub
