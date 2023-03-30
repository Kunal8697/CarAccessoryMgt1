Attribute VB_Name = "BasicModule1"
Option Explicit
Public cn As New ADODB.Connection
Public str As String
Public Dsc As String
Public iRptCaller As Integer
'Connectivity function from VIsual Bascic To MS Access
Public Function Con()
str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\CarAccessoryMgt1\AccessoryMgt.mdb;Persist Security Info=False"
cn.Open str
End Function
'To Clear The Contents Of Text Boxes After Save
Public Function ClearText(f As Form)
Dim i As Integer
For i = 0 To f.Controls.Count - 1
    If TypeName(f.Controls(i)) = "TextBox" Or TypeName(f.Controls(i)) = "ComboBox" Or TypeName(f.Controls(i)) = "MaskEdBox" Then
    f.Controls(i).Text = ""
     If TypeName(f.Controls(i)) = "OptionButton" Then
      f.Controls(i).Clear
      End If
          End If
    Next i
End Function
Public Sub LogError(ByVal errD As String)
On Error GoTo ErrHandler
        Open App.Path & "\log.txt" For Append As #1
        Write #1, Now & " " & errD
        Close #1
Exit Sub
ErrHandler:
Dsc = Err.Number & "" & Err.Description
LogError (Dsc)
MsgBox Dsc, , "Car Accessproes System"
End Sub
' Function is use when user insert name that time number or any special character are not allowed
Public Sub CheckName(ByRef key As Integer) ' checkname is name of function key as parameter
On Error GoTo ErrHandler
            If key = 32 Or key = 8 Then
               Exit Sub
            End If
                If key < 65 Or key > 91 And key < 97 Or key > 122 Then
                    key = 0
                End If
                    Exit Sub
ErrHandler:
Dsc = Err.Number & "" & Err.Description
LogError (Dsc)
MsgBox Dsc, , "Car Accessory System"
End Sub
'Conversion if number wrongly entered in integer & to enter null value if the input is not required
Public Function IntoStr(v As Variant) As String
    If IsNull(v) Then
        IntoStr = " "
    ElseIf IsDate(v) Then
        IntoStr = CStr(v)
    ElseIf IsNumeric(v) Then
        IntoStr = CStr(v)
    ElseIf CStr(v) = "" Then
        IntoStr = " "
    Else
        IntoStr = v
    End If
    Exit Function
ErrHandler:
    Dsc = Err.Number & "" & Err.Description
    LogError (Dsc)
    MsgBox Dsc, , "car Accessorysystem"
End Function
'To center the form in the screen
Public Sub CenterInScreen(f As Form)
    f.Top = (FrmMdi.Height - f.Height) / 6
    f.Left = (FrmMdi.Width - f.Width) / 2
End Sub

