VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDBSetting 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Connection Setting"
   ClientHeight    =   4410
   ClientLeft      =   1710
   ClientTop       =   720
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDBSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5550
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   120
      Top             =   3930
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "Connection"
      Height          =   765
      Left            =   90
      TabIndex        =   10
      Top             =   60
      Width           =   5295
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000013&
         Caption         =   "SQL Server"
         Height          =   345
         Left            =   2190
         TabIndex        =   12
         Top             =   300
         Width           =   1965
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000013&
         Caption         =   "Microsoft Access"
         Height          =   345
         Left            =   150
         TabIndex        =   11
         Top             =   300
         Value           =   -1  'True
         Width           =   1965
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   3930
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin VB.Frame fraSQLServer 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   120
      TabIndex        =   5
      Top             =   810
      Width           =   5295
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2430
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   17
         Top             =   2340
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "&Test"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3750
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   16
         Top             =   2340
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.TextBox txtServerName 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1620
         TabIndex        =   3
         Top             =   1800
         Width           =   3315
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1620
         TabIndex        =   0
         Top             =   360
         Width           =   3315
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1620
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   3315
      End
      Begin VB.TextBox txtDBName 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1620
         TabIndex        =   2
         Top             =   1320
         Width           =   3315
      End
      Begin VB.Label lblLogin 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server Name:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   270
         TabIndex        =   9
         Top             =   1860
         Width           =   1245
      End
      Begin VB.Label lblLogin 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   8
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label lblLogin 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   540
         TabIndex        =   7
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lblLogin 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DB Name:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   540
         TabIndex        =   6
         Top             =   1380
         Width           =   975
      End
   End
   Begin VB.Frame fraMsAccess 
      BackColor       =   &H80000013&
      Height          =   1245
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   5295
      Begin VB.CommandButton cmdBrowseFile 
         Caption         =   "&&"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4860
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   20
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   360
      End
      Begin VB.CommandButton cmdSaveAccess 
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2370
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   19
         Top             =   750
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.CommandButton cmdTestAccess 
         Caption         =   "&Test"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3690
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   18
         Top             =   750
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.TextBox txtFileName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1350
         TabIndex        =   14
         Top             =   210
         Width           =   3465
      End
      Begin VB.Label lblLogin 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Name:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   4
         Left            =   270
         TabIndex        =   15
         Top             =   270
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmDBSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowseFile_Click()
On Error GoTo diCancel
    Dialog1.CancelError = True
    Dialog1.Filter = "*.mdb|*.mdb"
    Dialog1.ShowOpen
    txtFileName.Text = Dialog1.FileName
diCancel:
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    WriteINI "Database", "UserName", txtUserName, AddPath(INICONN_FN)
    WriteINI "Database", "Password", Enkrip(txtPassword), AddPath(INICONN_FN)
    WriteINI "Database", "Connection", "SQL SERVER", AddPath(INICONN_FN)
    WriteINI "Database", "DatabaseName", txtDBName, AddPath(INICONN_FN)
    WriteINI "Database", "ServerName", txtServerName, AddPath(INICONN_FN)
    cmdClose.SetFocus
    cmdSave.Enabled = False
End Sub

Private Sub cmdTest_Click()
On Error GoTo Test_Error

    Dim strConnString    As String
    Dim objConn          As ADODB.Connection

    If (txtUserName = "") Then
        ShowMsg "User Name should not be emptied", False
        txtUserName.SetFocus
        Exit Sub
    End If
    If txtPassword = "" Then
        ShowMsg "Password should not be emptied", False
        txtPassword.SetFocus
        Exit Sub
    End If
    If txtDBName = "" Then
        ShowMsg "DB Name should not be emptied", False
        txtDBName.SetFocus
        Exit Sub
    End If
    If txtServerName = "" Then
        ShowMsg "Server Name should not be emptied", False
        txtServerName.SetFocus
        Exit Sub
    End If

    
    Screen.MousePointer = 11 'change cursor to hourglass

    strConnString = "Provider=SQLOLEDB;" & _
                        "Data Source=" & txtServerName & ";" & _
                        "Initial Catalog=" & txtDBName & ";" & _
                        "User ID=" & txtUserName & ";" & _
                        "Password=" & txtPassword.Text

    Set objConn = New ADODB.Connection
    With objConn
       .ConnectionString = strConnString
       .ConnectionTimeout = 5
       .Open
    End With
    objConn.Close
    Set objConn = Nothing

    ShowMsg "Success to connect into database.", True

    Screen.MousePointer = 0

    cmdTest.Enabled = False
    cmdSave.Enabled = True
    cmdSave.SetFocus

    Exit Sub

Test_Error:
    Screen.MousePointer = 0
    Select Case Err.Number
        Case -2147217843
            ShowMsg "User name or Password is incorrect.", False, "Connection test Error"
            txtUserName.SetFocus
        Case -2147467259
            ShowMsg "Can not open database or server requested." & vbCr & "Login failed.", False, "Connection test Error"
            txtUserName.SetFocus
        Case Else
            ShowError Err.Description, "Connection Test"
            txtUserName.SetFocus
    End Select
End Sub

Private Sub cmdTestAccess_Click()
On Error GoTo Test_Error

    Dim strConnString    As String
    Dim objConn          As ADODB.Connection

    If (txtFileName = "") Then
        ShowMsg "File Name should not be emptied", False
        txtFileName.SetFocus
        Exit Sub
    End If

    Screen.MousePointer = 11 'change cursor to hourglass

    strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                    "Data Source=" & txtFileName.Text & ";" & _
                    "Persist Security Info=True"

    Set objConn = New ADODB.Connection
    With objConn
       .ConnectionString = strConnString
       .ConnectionTimeout = 5
       .Open
    End With
    objConn.Close
    Set objConn = Nothing

    ShowMsg "Success to connect into database.", True

    Screen.MousePointer = 0

    cmdTestAccess.Enabled = False
    cmdSaveAccess.Enabled = True
    cmdSaveAccess.SetFocus

    Exit Sub

Test_Error:
    Screen.MousePointer = 0
    Select Case Err.Number
        Case -2147217843
            ShowMsg "User name or Password is incorrect.", False, "Connection test Error"
            txtUserName.SetFocus
        Case -2147467259
            ShowMsg "Can not open database." & vbCr & "Login failed.", False, "Connection test Error"
            txtUserName.SetFocus
        Case Else
            ShowError Err.Description, "Connection Test"
            txtUserName.SetFocus
    End Select
End Sub

Private Sub cmdSaveAccess_Click()
    WriteINI "Database", "Connection", "MS ACCESS", AddPath(INICONN_FN)
    WriteINI "Database", "DatabaseName", txtFileName, AddPath(INICONN_FN)
    cmdClose.SetFocus
    cmdSave.Enabled = False
End Sub

Private Sub Form_Load()
    Me.Top = 2750
    Me.Left = (Screen.Width - Me.Width) \ 2
    txtUserName = ReadINI("Database", "UserName", AddPath(INICONN_FN))
    txtPassword = Dekrip(ReadINI("Database", "Password", AddPath(INICONN_FN)))
    txtServerName = ReadINI("Database", "ServerName", AddPath(INICONN_FN))
    m_Conn = ReadINI("Database", "Connection", AddPath(INICONN_FN))
    If UCase(m_Conn) = "SQL SERVER" Then
        fraMsAccess.Visible = False
        fraSQLServer.Visible = True
        txtDBName = ReadINI("Database", "DatabaseName", AddPath(INICONN_FN))
        Me.Height = 4785
        cmdClose.Top = 3930
    Else
        fraMsAccess.Visible = True
        fraSQLServer.Visible = False
        txtFileName = ReadINI("Database", "DatabaseName", AddPath(INICONN_FN))
        Me.Height = 3100
        cmdClose.Top = 2250
    End If
    cmdTest.Enabled = False
    cmdSave.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_Encrypt = Nothing
End Sub

Private Sub Option1_Click()
    fraSQLServer.Visible = False
    fraMsAccess.Visible = True
    Me.Height = 3100
    cmdClose.Top = 2250
End Sub

Private Sub Option2_Click()
    fraSQLServer.Visible = True
    fraMsAccess.Visible = False
    Me.Height = 4785
    cmdClose.Top = 3930
End Sub

Private Sub txtDBName_Change()
    ChangeButtonStatus True
End Sub

Private Sub txtDBName_GotFocus()
    txtDBName.SelStart = 0
    txtDBName.SelLength = Len(txtDBName)
End Sub

Private Sub txtDBName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFileName_Change()
    cmdTestAccess.Enabled = (Trim(txtFileName) <> "")
End Sub

Private Sub txtPassword_Change()
    ChangeButtonStatus True
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtServerName_Change()
    ChangeButtonStatus True
End Sub

Private Sub txtServerName_GotFocus()
    txtServerName.SelStart = 0
    txtServerName.SelLength = Len(txtServerName)
End Sub

Private Sub txtServerName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtUserName_Change()
    ChangeButtonStatus True
End Sub

Private Sub ChangeButtonStatus(blnFlag As Boolean)
    If cmdTest.Enabled = Not blnFlag Then cmdTest.Enabled = blnFlag
    If cmdSave.Enabled = blnFlag Then cmdSave.Enabled = Not blnFlag
End Sub

Private Sub txtUserName_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName)
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
