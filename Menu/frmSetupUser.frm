VERSION 5.00
Begin VB.Form frmSetupUser 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSetupUser.frx":0000
   ScaleHeight     =   4770
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save Setting"
      Height          =   375
      Left            =   6780
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "UnSelect All"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4230
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select All"
      Height          =   375
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4230
      Width           =   1215
   End
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      Height          =   795
      Left            =   5580
      Picture         =   "frmSetupUser.frx":6D1B
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1350
      Width           =   885
   End
   Begin VB.CommandButton cmdCari 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find"
      Height          =   795
      Left            =   4670
      Picture         =   "frmSetupUser.frx":75E5
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1350
      Width           =   885
   End
   Begin VB.CommandButton cmdHapus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      Height          =   795
      Left            =   3760
      Picture         =   "frmSetupUser.frx":8427
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1350
      Width           =   885
   End
   Begin VB.CommandButton cmdBatal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   795
      Left            =   2850
      Picture         =   "frmSetupUser.frx":8CF1
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1350
      Width           =   885
   End
   Begin VB.CommandButton cmdSimpan 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      Height          =   795
      Left            =   1940
      Picture         =   "frmSetupUser.frx":95BB
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1350
      Width           =   885
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit"
      Height          =   795
      Left            =   1030
      Picture         =   "frmSetupUser.frx":9E85
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1350
      Width           =   885
   End
   Begin VB.CommandButton cmdTambah 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New"
      Height          =   795
      Left            =   120
      Picture         =   "frmSetupUser.frx":A74F
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1350
      Width           =   885
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   900
      Width           =   1725
   End
   Begin VB.TextBox txtUser 
      Height          =   345
      Left            =   1560
      TabIndex        =   3
      Top             =   510
      Width           =   1725
   End
   Begin VB.ListBox List1 
      Height          =   1860
      Left            =   150
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   2250
      Width           =   7875
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Setup User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   30
      TabIndex        =   5
      Top             =   60
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   1
      Top             =   990
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   0
      Top             =   540
      Width           =   675
   End
   Begin VB.Image Image11 
      Height          =   450
      Left            =   0
      Picture         =   "frmSetupUser.frx":B019
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8190
   End
End
Attribute VB_Name = "frmSetupUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim USER_ID As String
Private Sub EnableButton(Tambah As Boolean, Edit As Boolean, Hapus As Boolean, Simpan As Boolean, Batal As Boolean, Cari As Boolean, Setting As Boolean, keluar As Boolean)
    cmdTambah.Enabled = Tambah
    cmdEdit.Enabled = Edit
    cmdSimpan.Enabled = Simpan
    cmdBatal.Enabled = Batal
    cmdCari.Enabled = Cari
    cmdHapus.Enabled = Hapus
    cmdTutup.Enabled = keluar
End Sub
Private Sub EnableText(bl As Boolean)
    txtUser.Enabled = bl
    txtPassword.Enabled = bl
    If Not bl Then
        txtPassword.BackColor = &H8000000F
    Else
        txtPassword.BackColor = vbWhite
    End If
End Sub
Sub ClearAllText()
    txtUser.Text = ""
    txtPassword.Text = ""
End Sub
Private Sub ShowRecord()
'On Error Resume Next
    txtUser.Text = rs("USER_ID") & ""
    txtPassword.Text = Dekrip(Trim(rs("Password")))
    Haknya (txtUser.Text)
End Sub
Private Sub cmdBatal_Click()
On Error Resume Next
    EnableButton True, True, True, False, False, True, True, True
    rs.Close
    rs.Open "Select * From USER_MASTER where USER_ID='" & USER_ID & "'", g_objConn
    If Not (rs.EOF And rs.BOF) Then
        ShowRecord
    Else
        ClearAllText
    End If
    EnableText False
End Sub
Private Sub cmdCari_Click()
On Error Resume Next
    Y = Search("USER_MASTER", "USER_ID", "", "User ID", "Cari User")
    rs.Close
    rs.Open "Select * From USER_MASTER where USER_ID='" & Y & "'", g_objConn
    If Not (rs.EOF And rs.BOF) Then ShowRecord
End Sub

Private Sub cmdClose_Click()

End Sub

Private Sub cmdEdit_Click()
    USER_ID = txtUser.Text
    EnableButton False, False, False, True, True, False, False, True
    EnableText True
    txtUser.Enabled = False
End Sub

Private Sub cmdHapus_Click()
On Error GoTo Err2
    g_objConn.Execute "Delete From USER_MASTER where USER_ID ='" & txtUser.Text & "'"
    g_objConn.Execute "Delete From USER_MENU where USER_ID ='" & txtUser.Text & "'"
    Set rs = g_objConn.OpenResultSet("Select * From USER_MASTER")
    If Not (rs.EOF And rs.BOF) Then
        ShowRecord
    Else
        ClearAllText
    End If
    Exit Sub
Err2:
    ClearAllText
    EnableButton True, False, False, False, False, False, False, True
End Sub

Private Sub cmdSimpan_Click()
Dim rs As New ADODB.Recordset
    rs.Open "Select * From User_Master Where USER_ID='" & txtUser.Text & "'", g_objConn, adOpenDynamic, adLockOptimistic
    If (rs.EOF And rs.BOF) Then
        rs.AddNew
    Else
        Y = MsgBox("User already exist. Replace ?", 36, "")
        If Y <> 6 Then Exit Sub
    End If
    rs.Fields("USER_ID") = txtUser.Text
    rs.Fields("Password") = Enkrip(txtPassword.Text)
    rs.Update
    rs.Close
    Set rs = Nothing
    EnableButton True, True, True, False, False, True, True, True
    EnableText False
End Sub
Private Sub cmdTambah_Click()
    USER_ID = txtUser.Text
    EnableButton False, False, False, True, True, False, False, True
    ClearAllText
    EnableText True
    If txtUser.Enabled Then txtUser.SetFocus
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub
Private Sub Haknya(ByVal USER_ID As String)
Dim rs As New ADODB.Recordset
    rs.Open "Select * From USER_MENU Where USER_ID='" & UserUD & "'", g_objConn
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        While Not rs.EOF
            For i = 0 To List1.ListCount - 1
                If List1.List(i) = rs!MenuName & "" Then
                    List1.Selected(i) = True
                End If
            Next i
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Command1_Click()
    g_objConn.Execute "Delete From USER_MENU Where USER_ID='" & txtUser.Text & "'"
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
            g_objConn.Execute "INSERT INTO USER_MENU (USER_ID,MenuName,AppName) " & _
                             " VALUES ('" & txtUser.Text & "','" & List1.List(i) & "','" & App.EXEName & "')"

        End If
    Next i
    SetupMenu txtUser.Text
    Unload Me
End Sub
Private Sub Command2_Click()
    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = True
    Next i
    List1.ListIndex = 0
End Sub

Private Sub Command3_Click()
    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = False
    Next i
    List1.ListIndex = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Width = 8190
    
    For Each X In Desktop.Controls
        If TypeOf X Is MyMenu Then
            For i = 1 To X.Count
                List1.AddItem X.Name & " | " & X.GetMenuItem(i)
            Next i
        End If
    Next X
    rs.Open "Select * From USER_MASTER", g_objConn, adOpenDynamic, adLockOptimistic
    If Not (rs.EOF And rs.BOF) Then
        EnableButton True, True, True, False, False, True, True, True
        ClearAllText
    Else
        EnableButton True, False, True, False, False, True, True, True
    End If
    rs.MoveLast
    ShowRecord
    EnableText False
End Sub


Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub


Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Call DiEnter(KeyAscii)
End Sub

Private Sub txtUser_Change()
Dim rsHak As New ADODB.Recordset
    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = False
    Next i
    rsHak.Open "Select * From USER_MENU where USER_ID='" & txtUser.Text & "'", g_objConn
    If Not (rsHak.EOF And rsHak.BOF) Then
        rsHak.MoveFirst
        While Not rsHak.EOF
            For i = 0 To List1.ListCount - 1
                If UCase(List1.List(i)) = UCase(rsHak("MenuName")) Then
                    List1.Selected(i) = True
                End If
            Next i
            rsHak.MoveNext
        Wend
    End If
    If List1.ListCount > 0 Then List1.ListIndex = 0
End Sub
