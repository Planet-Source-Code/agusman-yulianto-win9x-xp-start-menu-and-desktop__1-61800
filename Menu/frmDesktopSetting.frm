VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDesktopSetting 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7665
   ControlBox      =   0   'False
   Icon            =   "frmDesktopSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDesktopSetting.frx":08CA
   ScaleHeight     =   7395
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6510
      Width           =   915
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Apply"
      Height          =   315
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6510
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   315
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6510
      Width           =   915
   End
   Begin VB.CommandButton cmdDefault 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Default"
      Height          =   315
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   2805
   End
   Begin VB.PictureBox PicBackGround 
      Height          =   315
      Left            =   2400
      ScaleHeight     =   255
      ScaleWidth      =   1155
      TabIndex        =   5
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowseColor 
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
      Height          =   315
      Left            =   3630
      TabIndex        =   4
      Top             =   6120
      Width           =   375
   End
   Begin VB.TextBox txtFileName 
      Height          =   345
      Left            =   150
      TabIndex        =   3
      Top             =   6870
      Width           =   2535
   End
   Begin VB.CommandButton cmdBrowseFile 
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
      Height          =   315
      Left            =   2730
      TabIndex        =   2
      Top             =   6930
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   6900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      Height          =   5025
      Left            =   90
      ScaleHeight     =   4965
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   630
      Width           =   7455
      Begin VB.Image Image2 
         Height          =   4965
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7395
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Background Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   150
      TabIndex        =   7
      Top             =   6120
      Width           =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Background Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   150
      TabIndex        =   6
      Top             =   6510
      Width           =   2370
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desktop Seeting"
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
      TabIndex        =   1
      Top             =   60
      Width           =   1740
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5325
      Left            =   60
      Picture         =   "frmDesktopSetting.frx":75E5
      Stretch         =   -1  'True
      Top             =   600
      Width           =   7515
   End
   Begin VB.Image Image11 
      Height          =   450
      Left            =   0
      Picture         =   "frmDesktopSetting.frx":247627
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10350
   End
End
Attribute VB_Name = "frmDesktopSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdApply_Click()
    Desktop.BackColor = PicBackGround.BackColor
    Desktop.Picture = Image2.Picture
    SaveSetting App.EXEName, "Desktop", "bgPicture", txtFileName.Text
    SaveSetting App.EXEName, "Desktop", "bgcolor", PicBackGround.BackColor
    Desktop.GambarDesktop = txtFileName.Text
End Sub

Private Sub cmdBrowseColor_Click()
On Error GoTo Err2
    CommonDialog1.CancelError = True
    CommonDialog1.ShowColor
    PicBackGround.BackColor = CommonDialog1.Color
    Picture1.BackColor = PicBackGround.BackColor
Err2:
End Sub

Private Sub cmdBrowseFile_Click()
On Error GoTo Err1
    With CommonDialog1
        .CancelError = True
        .Filter = "Picture|*.jpg;*.bmp;*.gif;*.wmf"
        .ShowOpen
        txtFileName.Text = .FileName
    End With
Err1:
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDefault_Click()
    PicBackGround.BackColor = &H80000001
    Picture1.BackColor = &H80000001
    txtFileName.Text = ""
    SaveSetting App.EXEName, "Desktop", "bgPicture", ""
    SaveSetting App.EXEName, "Desktop", "bgcolor", PicBackGround.BackColor
End Sub

Private Sub cmdOK_Click()
    Desktop.BackColor = PicBackGround.BackColor
    Desktop.Picture = Image2.Picture
    SaveSetting App.EXEName, "Desktop", "bgPicture", txtFileName.Text
    SaveSetting App.EXEName, "Desktop", "bgcolor", PicBackGround.BackColor
    Desktop.GambarDesktop = txtFileName.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Picture1.BackColor = GetSetting(App.EXEName, "Desktop", "bgcolor", &H80000001)
    PicBackGround.BackColor = Picture1.BackColor
    picName = GetSetting(App.EXEName, "Desktop", "bgPicture", "")
    Image2.Picture = LoadPicture(picName)
    txtFileName.Text = picName
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub isButton1_Click()
Unload Me
End Sub

Private Sub isButton2_Click()
Me.WindowState = 1
End Sub

Private Sub txtFileName_Change()
On Error GoTo Err2
    Image2.Picture = LoadPicture(txtFileName.Text)
    Exit Sub
Err2:
    Image2.Picture = LoadPicture()
End Sub
