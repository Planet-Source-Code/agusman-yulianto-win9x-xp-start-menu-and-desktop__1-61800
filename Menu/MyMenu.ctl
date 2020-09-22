VERSION 5.00
Begin VB.UserControl MyMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   KeyPreview      =   -1  'True
   ScaleHeight     =   465
   ScaleWidth      =   8850
   Begin Menu.MenuItem Menu 
      Height          =   645
      Index           =   1
      Left            =   510
      TabIndex        =   0
      Top             =   0
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   1138
   End
   Begin VB.Image Batas 
      Height          =   645
      Left            =   0
      Picture         =   "MyMenu.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "MyMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const m_def_Count = 0
Dim m_Count As Variant
Const m_def_Position = 1
Dim m_Position As Integer

    Public Event MenuClick(Position As Integer)
Public Event MouseMove(ByVal Position As Integer)
Private Sub ArrangeMenu()
    Menu(1).Top = UserControl.Height - Menu(1).Height
    Menu(1).Left = Batas.Left + Batas.Width
    For i = 2 To m_Count
        Menu(i).Top = Menu(i - 1).Top - Menu(i - 1).Height
    Next i
    For i = 1 To m_Count
        Menu(i).BackColor = &H8000000F
    Next i
    Menu(1).BackColor = vbBlue
    m_Position = 1
End Sub
    
Public Sub AddMenu(ByVal Caption As String, ByVal Pict As StdPicture, Optional Enabled As Boolean, Optional SubMenu As Boolean)
    If m_Count = 0 Then
        m_Count = m_Count + 1
        Menu(m_Count).Caption = Caption
        Set Menu(m_Count).Picture = Pict
        Menu(m_Count).SubMenu = SubMenu
        Menu(m_Count).Enabled = Enabled
        If Enabled Then Menu(m_Count).Tag = "Enabled" Else Menu(m_Count).Tag = "Disabled"
    Else
        m_Count = m_Count + 1
        UserControl.Height = UserControl.Height + Menu(1).Height
        Batas.Height = Batas.Height + Menu(1).Height
        Load Menu(m_Count)
        Menu(m_Count).Caption = Caption
        Set Menu(m_Count).Picture = Pict
        Menu(m_Count).Visible = True
        Menu(m_Count).SubMenu = SubMenu
        Menu(m_Count).Enabled = Enabled
        If Enabled Then Menu(m_Count).Tag = "Enabled" Else Menu(m_Count).Tag = "Disabled"
        ArrangeMenu
    End If
End Sub
Public Sub SetMenuItemByCaption(ByVal c As String, ByVal Enabled As Boolean)
    For i = 1 To m_Count
        If Menu(i).Caption = c Then
            Menu(i).Enabled = Enabled
            If Enabled Then
                Menu(i).Tag = "Enabled"
            Else
                Menu(i).Tag = "Disabled"
            End If
            Exit For
        End If
    Next i
End Sub
Public Sub SetMenuItemByIndex(ByVal i As Integer, ByVal Enabled As Boolean)
    Menu(i).Enabled = Enabled
    If Enabled Then Menu(i).Tag = "Enabled" Else Menu(i).Tag = "Disabled"
End Sub
Public Function GetMenuItem(ByVal i As Integer) As String
    GetMenuItem = Menu(i).Caption
End Function
Public Property Get Count() As Variant
    Count = m_Count
End Property

Public Property Let Count(ByVal New_Count As Variant)
    m_Count = New_Count
    PropertyChanged "Count"
End Property

Private Sub Menu_Click(Index As Integer)
    If Menu(Index).Enabled Then
        RaiseEvent MenuClick(m_Position)
    Else
        RaiseEvent MenuClick(0)
    End If
End Sub

Private Sub Menu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Menu(Index).Enabled Then
        Menu(Index).BackColor = vbBlue
        Menu(Index).ForeColor = vbWhite
        m_Position = Index
        RaiseEvent MouseMove(m_Position)
    Else
        RaiseEvent MouseMove(0)
    End If
    For i = 1 To m_Count
        If i <> Index Then
            Menu(i).BackColor = &H8000000F
            Menu(i).ForeColor = vbBlack
        End If
    Next i
End Sub

Private Sub UserControl_InitProperties()
    m_Count = m_def_Count
    Menu(1).Top = 0
    Batas.Top = 0
    Batas.Left = 0
    UserControl.Height = Menu(1).Height
    UserControl.Width = Batas.Width + Menu(1).Width
    m_Position = m_def_Position
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown
        If m_Position > 1 Then m_Position = m_Position - 1
    Case vbKeyU
        If m_Position < m_Count Then m_Position = m_Position + 1
    End Select
    For i = 1 To m_Count
        Menu(i).BackColor = &H8000000F
    Next i
    Menu(m_Position).BackColor = vbBlue
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(m_Position)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_Count = PropBag.ReadProperty("Count", m_def_Count)
    m_Position = PropBag.ReadProperty("Position", m_def_Position)
    Set VerticalPicture = PropBag.ReadProperty("VerticalPicture", Nothing)
    Menu(1).Caption = PropBag.ReadProperty("Caption", "Menu1")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Count", m_Count, m_def_Count)
    Call PropBag.WriteProperty("Position", m_Position, m_def_Position)
    Call PropBag.WriteProperty("VerticalPicture", VerticalPicture, Nothing)
    Call PropBag.WriteProperty("Caption", Menu(1).Caption, "Menu1")
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Position() As Integer
    Position = m_Position
End Property

Public Property Let Position(ByVal New_Position As Integer)
    m_Position = New_Position
    PropertyChanged "Position"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Batas,Batas,-1,Picture
Public Property Get VerticalPicture() As Picture
Attribute VerticalPicture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set VerticalPicture = Batas.Picture
End Property

Public Property Set VerticalPicture(ByVal New_VerticalPicture As Picture)
    Set Batas.Picture = New_VerticalPicture
    PropertyChanged "VerticalPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Menu(1),Menu,1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Menu(1).Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Menu(1).Caption() = New_Caption
    PropertyChanged "Caption"
End Property

