VERSION 5.00
Begin VB.UserControl MenuItem 
   BackColor       =   &H00FFFFC0&
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   ScaleHeight     =   420
   ScaleWidth      =   8655
   Begin VB.Image subIcon 
      Height          =   480
      Left            =   3075
      Picture         =   "MenuItem.ctx":0000
      Top             =   90
      Width           =   480
   End
   Begin VB.Label Judul 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu1"
      BeginProperty Font 
         Name            =   "Helvetica"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   765
   End
   Begin VB.Image PictureMenu 
      Height          =   405
      Left            =   120
      Picture         =   "MenuItem.ctx":0BC2
      Stretch         =   -1  'True
      Top             =   60
      Width           =   330
   End
End
Attribute VB_Name = "MenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const m_def_SubMenu = False
Dim m_SubMenu As Boolean
Public Event Click()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Judul,Judul,-1,MouseMove
'Default Property Values:
Const m_def_Enabled = -1
'Property Variables:
Dim m_Enabled As Boolean



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PictureMenu,PictureMenu,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = PictureMenu.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set PictureMenu.Picture = New_Picture
    PropertyChanged "Picture"
End Property

Private Sub Judul_Click()
    RaiseEvent Click
End Sub

Private Sub PictureMenu_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Judul.Caption = PropBag.ReadProperty("Caption", "Menu1")
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Judul.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    SubMenu = PropBag.ReadProperty("SubMenu", m_SubMenu)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
End Sub

Private Sub UserControl_Resize()
    subIcon.Left = UserControl.Width - subIcon.Width - 200
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Caption", Judul.Caption, "Menu1")
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", Judul.ForeColor, &H80000012)
    Call PropBag.WriteProperty("SubMenu", m_SubMenu, m_def_SubMenu)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Judul,Judul,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Judul.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Judul.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Judul,Judul,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Judul.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Judul.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Private Sub Judul_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Public Property Get SubMenu() As Boolean
    SubMenu = m_SubMenu
End Property

Public Property Let SubMenu(ByVal vNewValue As Boolean)
    m_SubMenu = vNewValue
    subIcon.Visible = m_SubMenu
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,-1
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    If Not m_Enabled Then
        Judul.ForeColor = &H80000010
    Else
        Judul.ForeColor = vbBlack
    End If
    PropertyChanged "Enabled"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
End Sub

