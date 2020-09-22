VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   180
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&CLOSE"
      Height          =   495
      Left            =   4770
      TabIndex        =   0
      Top             =   150
      Width           =   1275
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
