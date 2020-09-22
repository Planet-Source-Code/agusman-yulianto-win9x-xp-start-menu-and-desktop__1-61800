VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmMasterPropinsi 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6690
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMasterPropinsi.frx":0000
   ScaleHeight     =   2685
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport Report1 
      Left            =   5550
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtNama 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1800
      TabIndex        =   10
      Top             =   1200
      Width           =   2565
   End
   Begin VB.TextBox txtKode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1800
      TabIndex        =   0
      Top             =   675
      Width           =   765
   End
   Begin DesktopTemplate.isButton cmdSave 
      Height          =   615
      Left            =   1725
      TabIndex        =   1
      Top             =   1800
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1085
      Icon            =   "frmMasterPropinsi.frx":6D1B
      Style           =   9
      Caption         =   "&Save"
      IconAlign       =   3
      CaptionAlign    =   4
      iNonThemeStyle  =   10
      Tooltiptitle    =   "Simpan"
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DesktopTemplate.isButton cmdEdit 
      Height          =   615
      Left            =   900
      TabIndex        =   2
      Top             =   1800
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1085
      Icon            =   "frmMasterPropinsi.frx":75F5
      Style           =   9
      Caption         =   "&Edit"
      IconAlign       =   3
      CaptionAlign    =   4
      iNonThemeStyle  =   10
      Tooltiptitle    =   "Simpan"
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DesktopTemplate.isButton cmdFind 
      Height          =   615
      Left            =   3375
      TabIndex        =   3
      Top             =   1800
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1085
      Icon            =   "frmMasterPropinsi.frx":7ECF
      Style           =   9
      Caption         =   "&Find"
      IconAlign       =   3
      CaptionAlign    =   4
      iNonThemeStyle  =   10
      Tooltiptitle    =   "Simpan"
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DesktopTemplate.isButton cmdDelete 
      Height          =   615
      Left            =   4200
      TabIndex        =   4
      Top             =   1800
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1085
      Icon            =   "frmMasterPropinsi.frx":8D21
      Style           =   9
      Caption         =   "&Delete"
      IconAlign       =   3
      CaptionAlign    =   4
      iNonThemeStyle  =   10
      Tooltiptitle    =   "Simpan"
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DesktopTemplate.isButton cmdPrint 
      Height          =   615
      Left            =   5025
      TabIndex        =   5
      Top             =   1800
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1085
      Icon            =   "frmMasterPropinsi.frx":95FB
      Style           =   9
      Caption         =   "&Print"
      IconAlign       =   3
      CaptionAlign    =   4
      iNonThemeStyle  =   10
      Tooltiptitle    =   "Simpan"
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DesktopTemplate.isButton cmdClose 
      Height          =   615
      Left            =   5850
      TabIndex        =   6
      Top             =   1800
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1085
      Icon            =   "frmMasterPropinsi.frx":9ED5
      Style           =   9
      Caption         =   "&Close"
      IconAlign       =   3
      CaptionAlign    =   4
      iNonThemeStyle  =   10
      Tooltiptitle    =   "Simpan"
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DesktopTemplate.isButton cmdCancel 
      Height          =   615
      Left            =   2550
      TabIndex        =   7
      Top             =   1800
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1085
      Icon            =   "frmMasterPropinsi.frx":A7AF
      Style           =   9
      Caption         =   "&Cancel"
      IconAlign       =   3
      CaptionAlign    =   4
      iNonThemeStyle  =   10
      Tooltiptitle    =   "Simpan"
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DesktopTemplate.isButton cmdNew 
      Height          =   615
      Left            =   75
      TabIndex        =   8
      Top             =   1800
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1085
      Icon            =   "frmMasterPropinsi.frx":B089
      Style           =   9
      Caption         =   "&New"
      IconAlign       =   3
      CaptionAlign    =   4
      iNonThemeStyle  =   10
      Tooltiptitle    =   "Simpan"
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Master Propinsi"
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
      Left            =   75
      TabIndex        =   12
      Top             =   75
      Width           =   1635
   End
   Begin VB.Image Image11 
      Height          =   450
      Left            =   0
      Picture         =   "frmMasterPropinsi.frx":B963
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Propinsi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   11
      Top             =   1275
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Propinsi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   9
      Top             =   750
      Width           =   1455
   End
End
Attribute VB_Name = "frmMasterPropinsi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmdCancel_Click()
    If Trim(txtKode.Text) <> "" Then
        ShowRecord
    End If
    EnableButton Me, True, True, False, False, True, True, True, True
    EnableTextControl Me, False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Sub ShowRecord()
    If Not (rs.EOF And rs.BOF) Then
        txtKode.Text = rs!Kd_Propinsi
        txtNama.Text = rs!Nama_Propinsi
    Else
        ClearText Me
    End If
End Sub
Sub SaveRecord()
    rs.Fields("Kd_Propinsi") = txtKode.Text
    rs.Fields("Nama_Propinsi") = txtNama.Text
    rs.Update
End Sub

Private Sub cmdDelete_Click()
    Y = MsgBox("Hapus data propinsi " & txtNama.Text & "?", 36, "Konfirmasi")
    If Y = 6 Then g_objConn.Execute "Delete From Master_Propinsi Where Nama_Propinsi='" & txtNama.Text & "'"
    If Not rs.EOF Then
        rs.MoveNext
        ShowRecord
    Else
        If Not rs.BOF Then ShowRecord Else ClearText Me
    End If
End Sub

Private Sub cmdEdit_Click()
    EnableTextControl Me, True
    EnableButton Me, False, False, True, True, False, False, False, True
    txtNama.SetFocus
End Sub

Private Sub cmdFind_Click()
    Y = Search("Master_Propinsi", "Nama_Propinsi", , "Nama Propinsi", "Cari")
    If Y <> 6 Then
        rs.Find "Nama_Propinsi='" & Y & "'", , adSearchForward, 1
        If Not (rs.EOF) Then ShowRecord
    End If
End Sub

Private Sub cmdNew_Click()
    rs.AddNew
    ClearText Me
    EnableTextControl Me, True
    EnableButton Me, False, False, True, True, False, False, False, True
    txtKode.SetFocus
End Sub

Private Sub cmdSave_Click()
    SaveRecord
    EnableTextControl Me, False
    EnableButton Me, True, True, False, False, True, True, True, True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case vbKeyCode
    Case vbKeyUp: SendKeys "+{TAB}"
    Case vbKeyDown, vbKeyReturn: SendKeys "+{TAB}"
    End Select
End Sub

Private Sub Form_Load()
    EnableTextControl Me, False
    rs.Open "Select * from Master_Propinsi", g_objConn, adOpenDynamic, adLockOptimistic
    If Not (rs.EOF And rs.BOF) Then
        ShowRecord
        EnableButton Me, True, True, False, False, True, True, True, True
    Else
        EnableButton Me, True, False, False, False, False, False, False, True
    End If
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub
