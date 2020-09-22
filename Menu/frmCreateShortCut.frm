VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCreateShortCut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buat ShortCut"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCreateShortCut.frx":0000
   ScaleHeight     =   1665
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   675
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.exe|*.exe"
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CANCEL"
      Height          =   390
      Left            =   4050
      TabIndex        =   6
      Top             =   1125
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   390
      Left            =   3000
      TabIndex        =   5
      Top             =   1125
      Width           =   990
   End
   Begin VB.TextBox txtKeterangan 
      Height          =   390
      Left            =   1125
      TabIndex        =   4
      Top             =   600
      Width           =   3465
   End
   Begin VB.CommandButton cmdBrowse 
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
      Height          =   390
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   465
   End
   Begin VB.TextBox txtNamaFile 
      Height          =   390
      Left            =   1125
      TabIndex        =   1
      Top             =   150
      Width           =   3465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   675
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   225
      Width           =   855
   End
End
Attribute VB_Name = "frmCreateShortCut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FileName As String
Public Desc As String
Private Sub cmdBrowse_Click()
On Error GoTo Err1
    With CommonDialog1
        .CancelError = True
        .ShowOpen
        txtNamaFile.Text = .FileName
    End With
Err1:
End Sub

Private Sub cmdCancel_Click()
    FileName = ""
    Desc = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    FileName = txtNamaFile.Text
    Desc = txtKeterangan.Text
    Unload Me
End Sub

Private Sub Form_Load()
    FileName = ""
    Desc = ""
End Sub
