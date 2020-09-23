VERSION 5.00
Begin VB.UserControl LbDate 
   BackColor       =   &H00D8E9EC&
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1425
   LockControls    =   -1  'True
   ScaleHeight     =   1275
   ScaleWidth      =   1425
   Begin VB.Image Image1 
      Height          =   240
      Left            =   510
      Picture         =   "lbdate.ctx":0000
      Top             =   900
      Width           =   345
   End
   Begin VB.Label LbData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C56A28&
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   300
   End
End
Attribute VB_Name = "LbDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=LbData,LbData,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=LbData,LbData,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event Click() 'MappingInfo=LbData,LbData,-1,Click
Event DblClick() 'MappingInfo=LbData,LbData,-1,DblClick
'Default Property Values:
Const m_def_Tanggal = "0"
'Property Variables:
Dim m_Tanggal As String




Sub RePos()
    Width = 345
    LbData.Left = 0
    LbData.Width = Width
    Height = 255
End Sub


Private Sub UserControl_Resize()
    RePos
End Sub
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
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackHot", &HD8E9EC)
    LbData.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
    LbData.Caption = PropBag.ReadProperty("Value", "31")
    Set LbData.Font = PropBag.ReadProperty("Font", Ambient.Font)
    LbData.FontBold = PropBag.ReadProperty("FontBold", 0)
    m_Tanggal = PropBag.ReadProperty("Tanggal", m_def_Tanggal)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackHot", UserControl.BackColor, &HD8E9EC)
    Call PropBag.WriteProperty("ForeColor", LbData.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("Value", LbData.Caption, "31")
    Call PropBag.WriteProperty("Font", LbData.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", LbData.FontBold, 0)
    Call PropBag.WriteProperty("Tanggal", m_Tanggal, m_def_Tanggal)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LbData,LbData,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = LbData.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    LbData.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LbData,LbData,-1,Caption
Public Property Get Value() As String
Attribute Value.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Value = LbData.Caption
End Property

Public Property Let Value(ByVal New_Value As String)
    LbData.Caption() = New_Value
    PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LbData,LbData,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = LbData.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set LbData.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LbData,LbData,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = LbData.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    LbData.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property
'
Private Sub LbData_Click()
    RaiseEvent Click
End Sub

Private Sub LbData_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub LbData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub LbData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Tanggal() As String
    Tanggal = m_Tanggal
End Property

Public Property Let Tanggal(ByVal New_Tanggal As String)
    m_Tanggal = New_Tanggal
    PropertyChanged "Tanggal"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Tanggal = m_def_Tanggal
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Sub SetPicture(Optional IsNow As Boolean = True)
    If IsNow Then
        Picture = Image1.Picture
    Else
        Set Picture = Nothing
    End If
End Sub
