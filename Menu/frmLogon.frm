VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLogon 
   BackColor       =   &H80000013&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2955
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5445
   ControlBox      =   0   'False
   Icon            =   "frmLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogon.frx":0BC2
   ScaleHeight     =   2955
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDBSettings_1 
      Caption         =   "&&"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5070
      TabIndex        =   9
      Top             =   600
      Width           =   345
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   735
      Left            =   4110
      Picture         =   "frmLogon.frx":78DD
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2070
      Width           =   1305
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   735
      Left            =   2760
      Picture         =   "frmLogon.frx":81A7
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2070
      Width           =   1305
   End
   Begin VB.TextBox txtLoginName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   960
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Data :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Logon"
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
      TabIndex        =   3
      Top             =   60
      Width           =   660
   End
   Begin VB.Image Image11 
      Height          =   405
      Left            =   0
      Picture         =   "frmLogon.frx":8A71
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5730
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public bCancelled As Boolean
Private Sub cmdCancel_Click()
    bCancelled = True
    If Not Mulai Then Unload Me Else End
End Sub

Private Sub cmdDBSettings_Click()
On Error GoTo diCancel
    Dialog1.CancelError = True
    Dialog1.InitDir = App.Path
    Dialog1.Filter = "*.mdb|*.mdb"
    Dialog1.ShowOpen
    txtFileName.Text = Dialog1.FileName
    txtLoginName.SetFocus
diCancel:
End Sub

Function FieldAda() As Boolean
On Error GoTo Err10
Dim rs As New ADODB.Recordset
    FieldAda = True
    rs.Open "Select Bidang From UnitDinas", g_objConn, adOpenDynamic, adLockOptimistic
    Exit Function
Err10:
    FieldAda = False
End Function
Sub TambahField()
On Error Resume Next
Dim dB As DAO.Database
Dim tb As TableDef
    Set dB = DBEngine.OpenDatabase(g_strDBName)
    dB.TableDefs("UnitDinas").Fields.Append dB.TableDefs("UnitDinas").CreateField("Bidang", dbText, 50)
    dB.TableDefs("UnitDinas").Fields.Append dB.TableDefs("UnitDinas").CreateField("PemegangKas", dbText, 50)
    dB.TableDefs("UnitDinas").Fields.Append dB.TableDefs("UnitDinas").CreateField("JabatanPemegangKas", dbText, 120)
    dB.TableDefs("UnitDinas").Fields.Append dB.TableDefs("UnitDinas").CreateField("NipPemegangKas", dbText, 20)
    dB.TableDefs("UnitDinas").Fields.Append dB.TableDefs("UnitDinas").CreateField("BagianBelanja1", dbText, 120)
    dB.TableDefs("UnitDinas").Fields.Append dB.TableDefs("UnitDinas").CreateField("BagianBelanja2", dbText, 120)
    dB.TableDefs("Indikator").Fields.Append dB.TableDefs("Indikator").CreateField("LokasiKegiatan", dbText, 120)
    dB.TableDefs("Uraian").Fields.Append dB.TableDefs("Uraian").CreateField("DasarHukum", dbText, 255)
    dB.TableDefs("UraianPerubahan").Fields.Append dB.TableDefs("UraianPerubahan").CreateField("DasarHukum", dbText, 255)
    dB.TableDefs("TempUraian").Fields.Append dB.TableDefs("TempUraian").CreateField("DasarHukum", dbText, 255)
    
    
    Set tb = dB.CreateTableDef("Uraian")
    tb.Fields("DasarHukum").AllowZeroLength = True
    dB.TableDefs.Append tb
    
    Set tb = dB.CreateTableDef("UraianPerubahan")
    tb.Fields("DasarHukum").AllowZeroLength = True
    dB.TableDefs.Append tb
    
    Set tb = dB.CreateTableDef("TempUraian")
    tb.Fields("DasarHukum").AllowZeroLength = True
    dB.TableDefs.Append tb
    
    
    Set tb = dB.CreateTableDef("TempLaporanUraian")
    With tb
        .Fields.Append .CreateField("URAIAN", dbText, 225)
        .Fields.Append .CreateField("KEGIATAN", dbText, 225)
        .Fields.Append .CreateField("VOL", dbLong)
        .Fields.Append .CreateField("SAT", dbText, 15)
        .Fields.Append .CreateField("HGSAT", dbCurrency)
        .Fields.Append .CreateField("JLF", dbCurrency)
        .Fields.Append .CreateField("KodeRek", dbText, 21)
        .Fields.Append .CreateField("KodeUnitDinas", dbText, 5)
        .Fields.Append .CreateField("NoUrut", dbSingle)
        .Fields.Append .CreateField("KDA", dbText, 21)
        .Fields.Append .CreateField("KDB", dbText, 21)
        .Fields.Append .CreateField("KDC", dbText, 21)
        .Fields.Append .CreateField("KDD", dbText, 21)
        .Fields.Append .CreateField("KDE", dbText, 21)
        .Fields.Append .CreateField("TH", dbText, 4)
        .Fields.Append .CreateField("URAIANPERUBAHAN", dbText, 225)
        .Fields.Append .CreateField("VOLPERUBAHAN", dbLong)
        .Fields.Append .CreateField("SATPERUBAHAN", dbText, 15)
        .Fields.Append .CreateField("HGSATPERUBAHAN", dbCurrency)
        .Fields.Append .CreateField("JLFPERUBAHAN", dbCurrency)
    End With
    tb.Fields("SATPERUBAHAN").AllowZeroLength = True
    tb.Fields("URAIANPERUBAHAN").AllowZeroLength = True
    tb.Fields("SAT").AllowZeroLength = True
    
    dB.TableDefs.Append tb
    
    dB.Close
End Sub
Sub AmbilNama()
'On Error Resume Next
Dim rs As New ADODB.Recordset
    rs.Open "Select * From UnitDinas", g_objConn, adOpenDynamic, adLockReadOnly
    If Not (rs.EOF And rs.BOF) Then
        g_UnitDinas = rs!KodeUnitDinas & ""
        g_NamaUnitDinas = rs!NamaUnitDinas & ""
        g_NamaKepalaDinas = rs!KepalaDinas & ""
        g_NIPKepalaDinas = rs!NIP & ""
        g_Jabatan = rs!Jabatan & ""
        g_Bidang = rs!Bidang & ""
        g_PemegangKas = rs!PemegangKas & ""
        g_JabatanPemegangKas = rs!JabatanPemegangKas & ""
        g_nipPemegangKas = rs!NipPemegangKas & ""
        
        If g_NamaKepalaDinas = "" Then
            MsgBox "Informasi Dinas/Unit Kerja belum lengkap. Silakan tekan ENTER untuk melengkapi informasi dinas/unit kerja", vbInformation, "Info"
            frmSettingUnitKerja.Show 1
        End If
        rs.Close
        rs.Open "Select * From NamaTim", g_objConn, adOpenDynamic, adLockOptimistic
        If Not (rs.EOF And rs.BOF) Then
            For i = 1 To 9
                NIPAnggaran(i) = rs.Fields("nip" & i) & ""
            Next i
            For i = 1 To 9
                NMAnggaran(i) = rs.Fields("nm" & i) & ""
            Next i
            For i = 1 To 9
                JBTAnggaran(i) = rs.Fields("jbt" & i) & ""
            Next i
        Else
            ClearText Me
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmdDBSettings_1_Click()
On Error GoTo diCancel
    Dialog1.CancelError = True
    Dialog1.InitDir = App.Path
    Dialog1.Filter = "*.mdb|*.mdb"
    Dialog1.ShowOpen
    txtFileName.Text = Dialog1.FileName
    txtLoginName.SetFocus
diCancel:
End Sub

Private Sub cmdLogin_Click()
Dim dB As DAO.Databases
    If Dir(txtFileName.Text) = "" Then
        MsgBox "File tidak ada!", vbInformation, "Info"
        Exit Sub
    End If

     If Trim(txtFileName.Text) = "" Then
         ShowMsg "File Data harus diisi"
         txtFileName.SetFocus
         Exit Sub
     End If
     
     If Trim(txtLoginName.Text) = "" Then
         ShowMsg "Login Name harus diisi"
         txtLoginName.SetFocus
         Exit Sub
     End If
    
     If Trim(txtPassword.Text) = "" Then
         ShowMsg "Password harus diisi"
         txtPassword.SetFocus
         Exit Sub
     End If
    
    g_strDBName = txtFileName.Text
    Screen.MousePointer = 11
    strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                    "Data Source=" & txtFileName.Text & ";" & _
                    "Persist Security Info=True"
    Set g_objConn = New ADODB.Connection
    With g_objConn
       .ConnectionString = strConnString
       .ConnectionTimeout = 5
       .Open
    End With
    Screen.MousePointer = 0
    
     If ValidateLogin(txtLoginName.Text, txtPassword.Text) Then
        g_strUserName = txtLoginName.Text
        strDBPassword = txtPassword.Text
        'If Not FieldAda Then
        TambahField
        '    Set db = DBEngine.OpenDatabase(g_strDBName)
            
        'End If
        g_objConn.Execute "Delete from Uraian where KDE='x'"
        Mulai = False
        Call AmbilNama
        If Not TableExist("TempUraian") Then
            Call CreateTables
        End If
        SetupMenu g_strUserName
        Unload Me
     Else
         MsgBox "Login Incorrect"
         txtLoginName.SetFocus
     End If
End Sub

Private Sub mnuSettings_Click()
   frmDBSetting.Show vbModal
End Sub

Private Sub Form_Load()
    'Remove these lines
    '----------------------
    txtFileName.Text = App.Path + "\Start.mdb"
    txtPassword.Text = "adm"
    txtLoginName.Text = "adm"
    '----------------------
    
    Label3.Caption = "Logon to " & UCase(App.EXEName) & " "
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub isButton1_Click()
 bCancelled = True
    If Not Mulai Then Unload Me Else End
End Sub

Private Sub txtFileName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then cmdDBSettings_Click
End Sub

Private Sub txtLoginName_GotFocus()
    txtLoginName.SelStart = 0
    txtLoginName.SelLength = Len(txtLoginName)
End Sub

Private Sub txtLoginName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Trim(txtPassword.Text) = "" Then
            txtPassword.SetFocus
        Else
            cmdLogin_Click
        End If
    End If
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Trim(txtLoginName.Text) = "" Then
            txtLoginName.SetFocus
        Else
            cmdLogin_Click
        End If
    End If
End Sub
Private Function ValidateLogin(ByVal strUserName As String, ByVal strPassword As String) As Boolean
Dim rs As New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "Select * From User_Master Where USER_ID='" & strUserName & "' and PASSWORD='" & Enkrip(strPassword) & "'", g_objConn, adOpenDynamic, adLockOptimistic
    If (rs.EOF And rs.BOF) Then
        ValidateLogin = False
    Else
        ValidateLogin = True
    End If
    rs.Close
    Set rs = Nothing
End Function
