Attribute VB_Name = "access"
Option Explicit

Public Function access97(txtFileName As String) As String
On Error GoTo errHandler
    Dim ch(18) As Byte, x As Integer
    Dim sec
    If Trim(txtFileName) = "" Then Exit Function
    sec = Array(0, 134, 251, 236, 55, 93, 68, 156, 250, 198, 94, 40, 230, 19, 182, 138, 96, 84)
    Open txtFileName For Binary Access Read As #1 Len = 18
    Get #1, &H42, ch
    Close #1
    For x = 1 To 17
        access97 = access97 & Chr(ch(x) Xor sec(x))
    Next x
    Exit Function
errHandler:
    MsgBox "ERROR occcured:" & vbCrLf & Err.Number & ":  " & Err.Description, vbCritical, "ERROR"
    Exit Function
End Function

