VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Desktop 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   ClientHeight    =   7905
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   12915
   ControlBox      =   0   'False
   Icon            =   "Desktop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   12915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   345
      Left            =   3810
      TabIndex        =   11
      Top             =   4740
      Visible         =   0   'False
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MCI.MMControl MMControl1 
      Height          =   345
      Left            =   8400
      TabIndex        =   8
      Top             =   5910
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   609
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7620
      Top             =   2460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Menu.MyMenu Transaction 
      Height          =   555
      Left            =   3750
      TabIndex        =   7
      Top             =   1875
      Visible         =   0   'False
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   979
      VerticalPicture =   "Desktop.frx":0E42
      Caption         =   "Transaction"
   End
   Begin Menu.MyMenu Start 
      Height          =   585
      Left            =   3780
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   1032
      VerticalPicture =   "Desktop.frx":3E3D
      Caption         =   "Start"
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5175
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4770
      Top             =   5610
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":DE82
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":E75C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":F036
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":F910
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":101EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":10AC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":1139E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":11C78
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":12552
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":12E2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":13706
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":13FE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":15782
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":1605C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":16936
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":16ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":177AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":18084
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":1895E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Desktop.frx":19238
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   6840
   End
   Begin Menu.MyMenu Report 
      Height          =   555
      Left            =   3750
      TabIndex        =   3
      Top             =   1260
      Visible         =   0   'False
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   979
      VerticalPicture =   "Desktop.frx":19B12
      Caption         =   "Report"
   End
   Begin Menu.MyMenu InputData 
      Height          =   555
      Left            =   3750
      TabIndex        =   4
      Top             =   645
      Visible         =   0   'False
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   979
      VerticalPicture =   "Desktop.frx":2083D
      Caption         =   "DateEntry"
   End
   Begin Menu.MyMenu Admin 
      Height          =   555
      Left            =   3750
      TabIndex        =   6
      Top             =   30
      Visible         =   0   'False
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   979
      VerticalPicture =   "Desktop.frx":27568
      Caption         =   "Admin"
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7725
      TabIndex        =   5
      Top             =   7350
      Width           =   525
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   675
      TabIndex        =   1
      ToolTipText     =   "Klik disini untuk memulai App. RASK Bandar Lampung"
      Top             =   7350
      Width           =   540
   End
   Begin VB.Image Help 
      Height          =   480
      Left            =   720
      MouseIcon       =   "Desktop.frx":2E293
      MousePointer    =   99  'Custom
      Picture         =   "Desktop.frx":2E6D5
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image ImgStart 
      Height          =   480
      Left            =   120
      Picture         =   "Desktop.frx":2E9DF
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image ShortCut 
      Height          =   600
      Index           =   0
      Left            =   480
      Picture         =   "Desktop.frx":2F2A9
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image ImgTutupStatus 
      Height          =   450
      Left            =   6960
      Picture         =   "Desktop.frx":2FB73
      Top             =   6600
      Width           =   75
   End
   Begin VB.Image ImgStatus 
      Height          =   450
      Left            =   6960
      Picture         =   "Desktop.frx":2FD95
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   5160
   End
   Begin VB.Image Bar 
      Height          =   675
      Left            =   1440
      Picture         =   "Desktop.frx":2FE4F
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   11325
   End
   Begin VB.Label lblProses 
      BackStyle       =   0  'Transparent
      Caption         =   "Proses...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   3810
      TabIndex        =   12
      Top             =   4380
      Visible         =   0   'False
      Width           =   5955
   End
   Begin VB.Label lblVersi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ver. 1.0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9180
      MouseIcon       =   "Desktop.frx":344E1
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   90
      Width           =   3540
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   210
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   630
      Width           =   945
   End
   Begin VB.Image Image1 
      DataField       =   "Picture"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   6210
      Top             =   5970
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblShortcut 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shortcut"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   6180
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Menu mnPopUp 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnCreateShortCut 
         Caption         =   "Create Shortcut"
      End
      Begin VB.Menu mnRataKanPosisi 
         Caption         =   "&Arrage Icon"
      End
      Begin VB.Menu mnDesktop 
         Caption         =   "Setting Desktop"
      End
   End
   Begin VB.Menu mnPopUp2 
      Caption         =   "PopUp2"
      Visible         =   0   'False
      Begin VB.Menu mnDeleteShortCut 
         Caption         =   "Delete ShortCut"
      End
   End
End
Attribute VB_Name = "Desktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kx As Integer
Dim ky As Integer
Dim ShortCutAktif As Integer
Public GambarDesktop  As String
Sub CreateShortCut(ByVal NM As String, ByVal Caption As String, Optional gbr As StdPicture)
On Error Resume Next
    If NomorBolong = 999 Then
        NShortcut = NShortcut + 1
    Else
        NShortcut = NomorBolong
    End If
    Load Desktop.ShortCut(NShortcut)
    Desktop.ShortCut(NShortcut).Tag = NM
    Set Desktop.ShortCut(NShortcut).Picture = gbr
    Desktop.ShortCut(NShortcut).Left = Desktop.ShortCut(NShortcut - 1).Left + Desktop.ShortCut(NShortcut - 1).Width + 350
    Desktop.ShortCut(NShortcut).Visible = True
    
    Load Desktop.lblShortcut(NShortcut)
    Desktop.lblShortcut(NShortcut).Caption = Caption
    Desktop.lblShortcut(NShortcut).Visible = True
    Desktop.lblShortcut(NShortcut).Left = Desktop.ShortCut(NShortcut).Left - (Desktop.lblShortcut(NShortcut).Width \ 2) + 250
    
End Sub

Private Sub File_MenuClick(Position As Integer)
On Error GoTo Errnya1
    Select Case Position
    Case 5
        HideMenus
        Start.Visible = False
        frmGabungFile.Show 1
    Case 4
        HideMenus
        Start.Visible = False
        frmPisahFile.Show 1
    Case 3
        HideMenus
        Start.Visible = False
        With CommonDialog1
            .CancelError = True
            .Filter = "*.mdb|*.mdb"
            .ShowSave
            If .FileName <> "" Then
                If Dir(.FileName) <> "" Then
                    MsgBox "File sudah ada. Tidak boleh menyimpan file dengan nama file yang sudah ada.!", vbInformation, "Info"
                    Exit Sub
                End If
                FileCopy App.Path + "\MasterCopy.dat", .FileName
                Y = MsgBox("File Baru " & .FileName & " telah terbentuk. Apakah akan diaktifkan?", vbYesNo + vbQuestion, "File Baru")
                If Y = 6 Then
                    CloseConnection
                    OpenConnection .FileName
                    MsgBox "Sukses. Anda sekarang bekerja dengan file '" & g_strDBName & "'. Data akan tersimpan pada file tersebut." & vbCrLf & _
                    "Selanjutnya Anda harus mengisi informasi tentang instansi. Klik OK.", vbInformation, "Info"
                    frmSettingUnitKerja.Show 1
                End If
            End If
        End With
    Case 2
        HideMenus
        Start.Visible = False
        frmGantiNamaFile.Show 1
    Case 1
        HideMenus
        Start.Visible = False
        frmImport.Show
    End Select
Errnya1:
End Sub

Private Sub Help_Click()
    If Dir(App.Path + "\Help.avi") <> "" Then
        MMControl1.FileName = App.Path + "\Help.avi"
        MMControl1.Command = "Open"
        MMControl1.Command = "Play"
    Else
        MsgBox "Couldn't find file " & App.Path + "\Help.avi", vbInformation, "Info"
    End If
End Sub

Private Sub Report_MenuClick(Position As Integer)
    Select Case Position
    Case 3
        HideMenus
        Start.Visible = False
        MsgBox "This is report 3", vbInformation, "Info"
    Case 2
        HideMenus
        Start.Visible = False
        MsgBox "This is report 2", vbInformation, "Info"
    Case 1
        HideMenus
        Start.Visible = False
        MsgBox "This is report 1", vbInformation, "Info"
    End Select
End Sub

Private Sub Admin_MenuClick(Position As Integer)
    Select Case Position
    Case 1
        HideMenus
        Start.Visible = False
        frmSetupUser.Show
    Case 2
        HideMenus
        Start.Visible = False
        frmDesktopSetting.Show
    Case 3
        HideMenus
        Start.Visible = False
        frmTimAnggaran.Show
    Case 4
        HideMenus
        Start.Visible = False
        frmSettingUnitKerja.Show
    Case 5
        HideMenus
        Start.Visible = False
        frmImport.Show
    End Select
End Sub

Private Sub Bar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStart.ForeColor = vbBlack
End Sub

Private Sub Form_Activate()
    If Mulai Then
        frmLogon.Show 1
        Call AmbilNama
        SetupMenu g_strUserName
    End If
    Mulai = False
End Sub

Private Sub Form_Click()
    Start.Visible = False
    Transaction.Visible = False
    InputData.Visible = False
    Admin.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 And KeyCode = Asc("M") Then lblStart_Click
End Sub

Private Sub Form_Load()
    lblVersi.Caption = ""
    Mulai = True
    NShortcut = 0
    Start.AddMenu "Exit", ImageList1.ListImages(8).Picture, True
    Start.AddMenu "Logoff", ImageList1.ListImages(9).Picture, True
    Start.AddMenu "Admin", ImageList1.ListImages(11).Picture, True
    Start.AddMenu "Data Entry", ImageList1.ListImages(2).Picture, True
    Start.AddMenu "Transaction", ImageList1.ListImages(10).Picture, True, True
    Start.AddMenu "Report", ImageList1.ListImages(14).Picture, True, True
    Start.Visible = False
    
    Admin.AddMenu "Setup Menu", ImageList1.ListImages(10).Picture, True, True
    Admin.AddMenu "Desktop Setting", ImageList1.ListImages(16).Picture, True, True
        
    Transaction.AddMenu "Transaction 1", ImageList1.ListImages(3).Picture, True
    Transaction.AddMenu "Transaction 2", ImageList1.ListImages(3).Picture, True
    Transaction.AddMenu "Transaction 3", ImageList1.ListImages(3).Picture, True
    Transaction.Visible = False
    
    Report.AddMenu "Report 1", ImageList1.ListImages(14).Picture, True
    Report.AddMenu "Report 2", ImageList1.ListImages(14).Picture, True
    Report.AddMenu "Report 3", ImageList1.ListImages(14).Picture, True
    Report.Visible = False
    
    InputData.AddMenu "Data Entry 1", ImageList1.ListImages(15).Picture, True
    InputData.AddMenu "Data Entry 2", ImageList1.ListImages(15).Picture, True
    InputData.AddMenu "Data Entry 3", ImageList1.ListImages(15).Picture, True
        
    
    Data1.DatabaseName = App.Path + "\Desktop.dat"
    Data1.RecordSource = "Select * from MyDesktop"
    Data1.Refresh
    
    Call OpenConfiguration
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnPopUp
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideMenus
    lblStart.ForeColor = vbBlack
End Sub

Private Sub Form_Resize()
    Bar.Left = 0
    Bar.Top = Me.Height - Bar.Height
    Bar.Width = Width
    ImgStart.Left = 150
    ImgStart.Top = Height - ImgStart.Height - 100
    lblStart.Left = ImgStart.Left + ImgStart.Width + 100
    lblStart.Top = ImgStart.Top + 50
    ImgTutupStatus.Top = Bar.Top + 100
    ImgTutupStatus.Left = Me.Width - 5300
    ImgStatus.Left = ImgTutupStatus.Left + ImgTutupStatus.Width
    ImgStatus.Top = ImgTutupStatus.Top
    lblStatus.Left = ImgTutupStatus.Left + 100
    lblStatus.Top = ImgStatus.Top + 100
    Help.Left = Width - Help.Width - 500
    Help.Top = Height - Help.Height - 1200
    lblHelp.Left = Help.Left - 250
    lblHelp.Top = Help.Top + Help.Height + 50
    lblVersi.Left = Me.Width - lblVersi.Width + 1300
End Sub

Private Sub ImgStart_Click()
    Start.Left = 0
    Start.Top = Bar.Top - Start.Height - 100
    Start.Visible = True
End Sub

Private Sub lblHelp_Click()
    Help_Click
End Sub

Private Sub Transaction_MenuClick(Position As Integer)
    Select Case Position
    Case 3
        HideMenus
        Start.Visible = False
        MsgBox "This is transaction 3", vbInformation, "Info"
    Case 2
        HideMenus
        Start.Visible = False
        MsgBox "This is transaction 2", vbInformation, "Info"
    Case 1
        HideMenus
        Start.Visible = False
        MsgBox "This is transaction 1", vbInformation, "Info"
    End Select
End Sub

Private Sub lblShortcut_DblClick(Index As Integer)
    ShortCut_DblClick Index
End Sub

Private Sub lblStart_Click()
    Start.Left = 0
    Start.Top = Bar.Top - Start.Height
    Start.Visible = True
End Sub

Private Sub lblStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStart.ForeColor = vbWhite
End Sub

Private Sub InputData_MenuClick(Position As Integer)
 Select Case Position
    Case 1
        HideMenus
        Start.Visible = False
        MsgBox "This is Data Entry 1", vbInformation, "Info"
    Case 2
        HideMenus
        Start.Visible = False
        MsgBox "This is Data Entry 2", vbInformation, "Info"
    Case 3
        HideMenus
        Start.Visible = False
        MsgBox "This is Data Entry 3", vbInformation, "Info"
   End Select
End Sub

Private Sub mnCreateShortCut_Click()
    frmCreateShortCut.Show 1
    If frmCreateShortCut.FileName <> "" Then
        CreateShortCut frmCreateShortCut.FileName, frmCreateShortCut.Desc, ShortCut(0).Picture
    End If
End Sub

Private Sub mnDeleteShortCut_Click()
    Y = MsgBox("Delete shortcut " & lblShortcut(ShortCutAktif).Caption & "?", 36, "Confirmation")
    If Y <> 6 Then Exit Sub
    Data1.Recordset.FindFirst "AppName='" & ShortCut(ShortCutAktif).Tag & "'"
    If Not Data1.Recordset.NoMatch Then Data1.Recordset.Delete
    Data1.Refresh
    Unload ShortCut(ShortCutAktif)
    Unload lblShortcut(ShortCutAktif)
End Sub

Private Sub mnDesktop_Click()
    frmDesktopSetting.Show
End Sub

Private Sub mnRataKanPosisi_Click()
    ShortCut(0).Top = 50
    lblShortcut(0).Top = ShortCut(0).Top + ShortCut(0).Height + 100
    For i = 1 To ShortCut.Count - 1
        ShortCut(i).Top = lblShortcut(i - 1).Top + 500
        lblShortcut(i).Top = ShortCut(i).Top + ShortCut(i).Height
        ShortCut(i).Left = ShortCut(0).Left
        lblShortcut(i).Left = ShortCut(0).Left
    Next i
End Sub

Private Sub ShortCut_DblClick(Index As Integer)
    If Dir(ShortCut(Index).Tag) <> "" Then
        If UCase(Left(FileTitle(ShortCut(Index).Tag), 2)) = "ST" Then
            X = Shell(ShortCut(Index).Tag & " " & g_strDBName & "~" & g_strDBServerName & "~" & g_strDBUserName & "~" & strDBPassword & "~" & g_strUserName, vbMaximizedFocus)
        Else
            X = Shell(ShortCut(Index).Tag, vbMaximizedFocus)
        End If
    Else
        Y = MsgBox("File " & ShortCut(Index).Tag & " tidak ditemukan. Apakah Anda ingin mencari?", 36, "File Tidak Ditemukan")
        If Y = 6 Then frmCreateShortCut.Show 1
    End If
End Sub

Private Sub ShortCut_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShortCutAktif = Index
    If Button = 1 Then
        kx = X
        ky = Y
    ElseIf Button = 2 Then
        PopupMenu mnPopUp2
    End If
End Sub

Private Sub ShortCut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ShortCut(Index).Left = ShortCut(Index).Left + (X - kx)
        ShortCut(Index).Top = ShortCut(Index).Top + (Y - ky)
        lblShortcut(Index).Top = ShortCut(Index).Top + ShortCut(Index).Height
        lblShortcut(Index).Left = ShortCut(Index).Left - (lblShortcut(Index).Width \ 2) + 150
    End If
End Sub
Function HideMenus()
    For Each X In Me.Controls
        If TypeOf X Is MyMenu Then
            If X.Name <> "Start" Then X.Visible = False
        End If
    Next
End Function
Private Sub Start_MenuClick(Position As Integer)
    Select Case Position
    Case 1
        SimpanKonfigurasi
        End
    Case 2
        HideMenus
        Start.Visible = False
        frmLogon.Show 1
    Case 4
        HideMenus
    Case 5
        HideMenus
    End Select
End Sub
Private Sub Start_MouseMove(ByVal Position As Integer)
    Select Case Position
    Case 6
        Transaction.Visible = False
        InputData.Visible = False
        Admin.Visible = False
        With Report
            .Left = Start.Left + Start.Width
            .Top = Start.Top + ((Start.Count - Position)) * 675 - 800
            .Visible = True
        End With
    Case 5
        Report.Visible = False
        InputData.Visible = False
        Admin.Visible = False
        With Transaction
            .Left = Start.Left + Start.Width
            .Top = Start.Top + ((Start.Count - Position)) * 675 - 800
            .Visible = True
        End With
    Case 4
        Transaction.Visible = False
        Report.Visible = False
        Admin.Visible = False
        With InputData
            .Left = Start.Left + Start.Width
            .Top = Start.Top + ((Start.Count - Position)) * 675 - 800
            .Visible = True
        End With
    Case 3
        Report.Visible = False
        InputData.Visible = False
        Transaction.Visible = False
        With Admin
            .Left = Start.Left + Start.Width
            .Top = Start.Top + ((Start.Count - Position)) * 675 - 500
            .Visible = True
        End With
     Case 0, 1, 2, 5
        HideMenus
    End Select
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Index = 1 Then
        Start.Left = 0
        Start.Top = StatusBar1.Top - Start.Height
        Start.Visible = True
    End If
End Sub
Sub SimpanKonfigurasi()
On Error Resume Next
    Data1.Database.Execute "Delete from MyDesktop"
    Data1.Refresh
    For i = 1 To ShortCut.Count - 1
        If ShortCut(i).Tag <> "" Then
            Data1.Recordset.AddNew
            Data1.Recordset.Fields("AppName") = ShortCut(i).Tag
            Data1.Recordset.Fields("AppDesc") = lblShortcut(i).Caption
            Data1.Recordset.Fields("Left") = ShortCut(i).Left
            Data1.Recordset.Fields("Top") = ShortCut(i).Top
            Image1.Picture = ShortCut(i).Picture
            Data1.Recordset.Update
        End If
    Next i
    Data1.Recordset.Close
    
    SaveSetting App.EXEName, "Desktop", "bgcolor", Me.BackColor
    SaveSetting App.EXEName, "Desktop", "bgPicture", GambarDesktop
End Sub
Sub OpenConfiguration()
    Data1.Refresh
    While Not Data1.Recordset.EOF
        NShortcut = NShortcut + 1
        Load ShortCut(NShortcut)
        If Data1.Recordset.Fields("AppName") & "" <> "" Then
            ShortCut(NShortcut).Tag = Data1.Recordset.Fields("AppName")
            ShortCut(NShortcut).Picture = Image1.Picture
            ShortCut(NShortcut).Left = Data1.Recordset.Fields("Left")
            ShortCut(NShortcut).Top = Data1.Recordset.Fields("Top")
            ShortCut(NShortcut).Visible = True
            Load lblShortcut(NShortcut)
            lblShortcut(NShortcut).Caption = Data1.Recordset.Fields("AppDesc")
            lblShortcut(NShortcut).Visible = True
            lblShortcut(NShortcut).Top = ShortCut(NShortcut).Top + ShortCut(NShortcut).Height
            lblShortcut(NShortcut).Left = ShortCut(NShortcut).Left - (lblShortcut(NShortcut).Width \ 2) + 150
        End If
        Data1.Recordset.MoveNext
    Wend
    GambarDesktop = GetSetting(App.EXEName, "Desktop", "bgPicture", "")
    Me.BackColor = GetSetting(App.EXEName, "Desktop", "bgcolor", &H80000001)
    Me.Picture = LoadPicture(GambarDesktop)
End Sub
Function NomorBolong() As Integer
On Error GoTo Errb
    NomorBolong = 999
    For i = 1 To NShortcut
        X = ShortCut(i).Tag
    Next i
    Exit Function
Errb:
    NomorBolong = X
End Function

Private Sub Timer1_Timer()
    lblStatus.Caption = g_strDBName & " - " & g_strUserName & " - " & Format(Now, "DD MMM YY HH:MM:SS")
End Sub
