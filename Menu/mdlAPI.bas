Attribute VB_Name = "mdlAPI"
Option Explicit

' API Profile String functions:
#If Win32 Then
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#Else
Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
#End If

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Const GWL_STYLE = (-16)
Const ES_NUMBER = &H2000&
Const ES_UPPERCASE = &H8

Function NameOfPC(MachineName As String) As Long
    Dim NameSize As Long
    Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
End Function

Function ReadINI(strSectionHeader As String, strVariableName As String, strFileName As String) As String
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    ReadINI = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function

Function WriteINI(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    WriteINI = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function

Public Sub SetNumber(NumberText As TextBox, Flag As Boolean)
    Dim curstyle As Long, NewStyle As Long

    curstyle = GetWindowLong(NumberText.hWnd, GWL_STYLE)

    If Flag Then
       curstyle = curstyle Or ES_NUMBER
    Else
       curstyle = curstyle And (Not ES_NUMBER)
    End If

    NewStyle = SetWindowLong(NumberText.hWnd, GWL_STYLE, curstyle)
    NumberText.Refresh
End Sub

Public Sub SetUppercase(UpperText As TextBox, Flag As Boolean)
    Dim curstyle As Long, NewStyle As Long

    curstyle = GetWindowLong(UpperText.hWnd, GWL_STYLE)

    If Flag Then
       curstyle = curstyle Or ES_UPPERCASE
    Else
       curstyle = curstyle And (Not ES_UPPERCASE)
    End If

    NewStyle = SetWindowLong(UpperText.hWnd, GWL_STYLE, curstyle)
    UpperText.Refresh
End Sub

Function GetLocation() As String
    GetLocation = ReadINI("Location Info", "Location", AddPath(INICONN_FN))
End Function

Public Function WinDir() As String
    Dim strBuff As String * 255
    
    strBuff = String(255, Chr(0))
    GetWindowsDirectory strBuff, 255
    
    WinDir = Left(strBuff, InStr(strBuff, Chr(0)) - 1)
End Function
