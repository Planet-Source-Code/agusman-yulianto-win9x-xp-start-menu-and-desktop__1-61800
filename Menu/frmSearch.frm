VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&CANCEL"
      Height          =   375
      Left            =   5130
      TabIndex        =   3
      Top             =   540
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5130
      TabIndex        =   2
      Top             =   90
      Width           =   1065
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      ItemData        =   "frmSearch.frx":0000
      Left            =   1080
      List            =   "frmSearch.frx":0002
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   90
      Width           =   3945
   End
   Begin VB.Label lblField 
      Caption         =   "Field"
      Height          =   1245
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   885
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim m_table As String
Dim m_Field As String
Dim m_Criteria As String
Dim m_title As String
Dim m_Fieldtitle As String
Public Selected As String
Private Sub Fill(ByVal Table As String, ByVal Field As String, Optional Criteria As String)
    If Criteria <> "" Then
        rs.Open "Select " & Field & " as x From " & Table & " WHERE " & Criteria, g_objConn
    Else
        rs.Open "Select " & Field & " as x From " & Table, g_objConn
    End If
    If Not (rs.EOF And rs.BOF) Then
        Combo1.Clear
        While Not rs.EOF
            Combo1.AddItem rs("x") & ""
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set rs = Nothing
End Sub
Public Property Get Table() As String
    Table = m_table
End Property

Public Property Let Table(ByVal vNewValue As String)
    m_table = vNewValue
End Property
Public Property Get Field() As String
    Field = m_Field
End Property

Public Property Let Field(ByVal vNewValue As String)
    m_Field = vNewValue
End Property

Public Property Get Criteria() As String
    Criteria = m_Criteria
End Property

Public Property Let Criteria(ByVal vNewValue As String)
    m_Criteria = vNewValue
End Property

Private Sub Combo1_DblClick()
    Call Command1_Click
End Sub

Private Sub Command1_Click()
Dim ada As Boolean
    ada = False
    For I = 0 To Combo1.ListCount - 1
        If Combo1.Text = Left(Combo1.List(I), Len(Combo1.Text)) Then
            ada = True
            Exit For
        End If
    Next I
    If ada Then
        Selected = Combo1.Text
        Unload Me
    Else
        Selected = ""
        MsgBox "Data '" & Combo1.Text & "' tidak ada!", vbInformation, "Data tidak ada"
    End If
End Sub

Private Sub Command2_Click()
    Selected = ""
    Unload Me
End Sub

Private Sub Form_Load()
    Call Fill(m_table, m_Field, m_Criteria)
    lblField.caption = m_Field
End Sub

Public Property Get Title() As String
    Title = m_title
End Property

Public Property Let Title(ByVal vNewValue As String)
    m_title = vNewValue
    caption = m_title
End Property
Public Property Get FieldTitle() As String
    FieldTitle = m_Fieldtitle
End Property

Public Property Let FieldTitle(ByVal vNewValue As String)
    m_Fieldtitle = vNewValue
    lblField.caption = m_Fieldtitle
End Property

Public Sub Center()
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
End Sub
