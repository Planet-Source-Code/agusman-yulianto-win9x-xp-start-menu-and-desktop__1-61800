Attribute VB_Name = "mdlMain"

Public Declare Function JFPCloseDevice Lib "fpapiop.dll" (ByVal hDevice As Long) As Long
Public Declare Function JFPInitDevice Lib "fpapiop.dll" (ByVal nType As Integer) As Long
Public Declare Function JFPGetFinger Lib "fpapiop.dll" (ByVal pFinger As Long, ByRef pRawImage As Byte, ByRef pMinutiae As Byte, ByVal isAuto As Integer) As Long
Public Declare Function JFPDrawRawImage Lib "fpapiop.dll" (ByVal hWnd As Long, ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByRef pRawImage As Byte) As Long
Public Declare Function JFPMatchFinger Lib "fpapiop.dll" (ByRef pMinutiae1 As Byte, ByRef pMinutiae2 As Byte, ByVal level As Integer) As Long

Global hDevice As Long
Global Minutiae1(256) As Byte
Global Minutiae2(256) As Byte

Global VerifiedBy As String

Type Fingers
    F1T1 As String
    F1T2 As String
    F2T1 As String
    F2T2 As String
    NIP As String
End Type
Global DBFingers(1 To 800) As Fingers
Global NKaryawan As Integer

Sub LoadDatabase()
On Error Resume Next
Dim MyRst As New ADODB.Recordset
    MyRst.Open "Select NIP,minutiaeF1T1,minutiaeF1T2,minutiaeF2T1,minutiaeF2T2 From Karyawan", g_objConn, adOpenDynamic, adLockOptimistic
    NKaryawan = 1
    While Not MyRst.EOF
        With DBFingers(NKaryawan)
            .F1T1 = MyRst("minutiaeF1T1").GetChunk(MyRst("minutiaeF1T1").ActualSize)
            .F1T2 = MyRst("minutiaeF1T2").GetChunk(MyRst("minutiaeF1T2").ActualSize)
            .F2T1 = MyRst("minutiaeF2T1").GetChunk(MyRst("minutiaeF2T1").ActualSize)
            .F2T2 = MyRst("minutiaeF2T2").GetChunk(MyRst("minutiaeF2T2").ActualSize)
            .NIP = MyRst("NIP")
        End With
        NKaryawan = NKaryawan + 1
        MyRst.MoveNext
    Wend
    MyRst.Close
    Set MyRst = Nothing
End Sub
