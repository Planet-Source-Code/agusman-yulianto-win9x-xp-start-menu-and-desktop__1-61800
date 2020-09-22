Attribute VB_Name = "mdlDesktop"
Global Mulai As Boolean
Global NShortcut As Integer
Global ShareFolder As String
Function NomorBolong() As Integer
On Error GoTo Errb
    NomorBolong = 999
    For i = 1 To NShortcut
        X = Desktop.ShortCut(i).Tag
    Next i
    Exit Function
Errb:
    NomorBolong = i
End Function

