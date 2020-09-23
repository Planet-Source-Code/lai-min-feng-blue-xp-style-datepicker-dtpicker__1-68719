VERSION 5.00
Begin VB.UserControl OsenXPSpin 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1185
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   360
   ScaleWidth      =   1185
   ToolboxBitmap   =   "OsenXPSpin.ctx":0000
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1770
      Width           =   915
   End
   Begin VB.TextBox TxtSpin 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   75
      TabIndex        =   0
      Text            =   "0"
      Top             =   60
      Width           =   720
   End
   Begin VB.Image Img 
      Height          =   135
      Index           =   7
      Left            =   1410
      Picture         =   "OsenXPSpin.ctx":0312
      Top             =   1140
      Width           =   225
   End
   Begin VB.Image Img 
      Height          =   135
      Index           =   6
      Left            =   1680
      Picture         =   "OsenXPSpin.ctx":0504
      Top             =   1140
      Width           =   225
   End
   Begin VB.Image Img 
      Height          =   135
      Index           =   5
      Left            =   1710
      Picture         =   "OsenXPSpin.ctx":06F6
      Top             =   780
      Width           =   225
   End
   Begin VB.Image Img 
      Height          =   135
      Index           =   4
      Left            =   1710
      Picture         =   "OsenXPSpin.ctx":08E8
      Top             =   960
      Width           =   225
   End
   Begin VB.Image Img 
      Height          =   135
      Index           =   3
      Left            =   1440
      Picture         =   "OsenXPSpin.ctx":0ADA
      Top             =   960
      Width           =   225
   End
   Begin VB.Image Img 
      Height          =   135
      Index           =   2
      Left            =   1710
      Picture         =   "OsenXPSpin.ctx":0CCC
      Top             =   600
      Width           =   225
   End
   Begin VB.Image Img 
      Height          =   135
      Index           =   1
      Left            =   1440
      Picture         =   "OsenXPSpin.ctx":0EBE
      Top             =   600
      Width           =   225
   End
   Begin VB.Image Img 
      Height          =   135
      Index           =   0
      Left            =   1440
      Picture         =   "OsenXPSpin.ctx":10B0
      Top             =   780
      Width           =   225
   End
   Begin VB.Shape ShapeBorder 
      BorderColor     =   &H00B99D7F&
      Height          =   330
      Left            =   0
      Top             =   0
      Width           =   1125
   End
   Begin VB.Image ImgDown 
      Height          =   135
      Left            =   870
      Picture         =   "OsenXPSpin.ctx":12A2
      Top             =   165
      Width           =   225
   End
   Begin VB.Image ImgUp 
      Height          =   135
      Left            =   870
      Picture         =   "OsenXPSpin.ctx":1494
      Top             =   30
      Width           =   225
   End
End
Attribute VB_Name = "OsenXPSpin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_Max = 32000
Const m_def_Min = -32000
Const m_def_Decimal = 0
Const m_def_LargeChange = 1
Const m_def_Value = 0
'Property Variables:
Dim m_Font As Font
Dim m_Max As Long
Dim m_Min As Long
Dim m_Decimal As Integer
Dim m_LargeChange As Double
Dim m_Value As Double
'Event Declarations:
Event Change() 'MappingInfo=TxtSpin,TxtSpin,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."

Private Sub RePos()
Dim I As Integer
    If Width < 400 Then Width = 400
    ShapeBorder.Width = Width
    ImgUp.Left = Width - 255
    ImgDown.Left = ImgUp.Left
    TxtSpin.Width = Width - 345
    Height = 330
    TxtSpin.Top = 60
    If TxtSpin.FontSize > 8 Then
        I = TxtSpin.FontSize - 8
        I = I * 15
        TxtSpin.Top = TxtSpin.Top - I
    End If
    
    
End Sub

Sub ResetPic()
    If ImgUp.Picture <> Img(2).Picture Or _
        ImgDown.Picture <> Img(1).Picture Then
        ImgUp.Picture = Img(2).Picture
        ImgDown.Picture = Img(1).Picture
    End If
End Sub

Private Sub ImgDown_Click()
    Text1.SetFocus
    If Value > Min Then
        Value = Value - LargeChange
    End If
End Sub

Private Sub ImgDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgDown.Picture = Img(3).Picture
End Sub

Private Sub ImgDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ImgDown.Picture <> Img(0).Picture Then ImgDown.Picture = Img(0).Picture
    ImgUp.Picture = Img(2).Picture
End Sub

Private Sub ImgUp_Click()
    Text1.SetFocus
    If Value < Max Then
        Value = Value + LargeChange
    End If
End Sub

Private Sub ImgUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgUp.Picture = Img(4).Picture
End Sub

Private Sub ImgUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ImgUp.Picture <> Img(5).Picture Then ImgUp.Picture = Img(5).Picture
    ImgDown.Picture = Img(1).Picture
End Sub

Private Sub TxtSpin_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then KeyAscii = 0
End Sub

Private Sub TxtSpin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetPic
End Sub


Private Sub UserControl_InitProperties()
    BackColor = vbWhite
    m_Max = m_def_Max
    m_Min = m_def_Min
    m_Decimal = m_def_Decimal
    m_LargeChange = m_def_LargeChange
    m_Value = m_def_Value
    Set m_Font = Ambient.Font
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetPic
End Sub

Private Sub UserControl_Resize()
    RePos
    
End Sub

Private Sub TxtSpin_Change()
    RaiseEvent Change
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    TxtSpin.Enabled = New_Enabled
    ImgUp.Enabled = New_Enabled
    ImgDown.Enabled = New_Enabled
    If New_Enabled = False Then
        ImgUp.Picture = Img(6).Picture
        ImgDown.Picture = Img(7).Picture
        ShapeBorder.BorderColor = &HC0C0C0
    Else
        ResetPic
        ShapeBorder.BorderColor = &HB99D7F
    End If
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TxtSpin,TxtSpin,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = TxtSpin.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    TxtSpin.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TxtSpin,TxtSpin,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = TxtSpin.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    TxtSpin.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,32000
Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    PropertyChanged "Max"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,-32000
Public Property Get Min() As Long
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)
    m_Min = New_Min
    PropertyChanged "Min"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1
Public Property Get LargeChange() As Double
    LargeChange = m_LargeChange
End Property

Public Property Let LargeChange(ByVal New_LargeChange As Double)
    m_LargeChange = New_LargeChange
    PropertyChanged "LargeChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=4,0,0,0
Public Property Get Value() As Double
    Value = Val(TxtSpin)
End Property

Public Property Let Value(ByVal New_Value As Double)
    TxtSpin = New_Value
    PropertyChanged "Value"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    TxtSpin.Locked = PropBag.ReadProperty("Locked", False)
    TxtSpin.Text = PropBag.ReadProperty("Text", "Text1")
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Decimal = PropBag.ReadProperty("Decimal", m_def_Decimal)
    m_LargeChange = PropBag.ReadProperty("LargeChange", m_def_LargeChange)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    Set TxtSpin.Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    RePos

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    With PropBag
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("Locked", TxtSpin.Locked, False)
        Call .WriteProperty("Text", TxtSpin.Text, "Text1")
        Call .WriteProperty("Max", m_Max, m_def_Max)
        Call .WriteProperty("Min", m_Min, m_def_Min)
        Call .WriteProperty("Decimal", m_Decimal, m_def_Decimal)
        Call .WriteProperty("LargeChange", m_LargeChange, m_def_LargeChange)
        Call .WriteProperty("Value", m_Value, m_def_Value)
        Call .WriteProperty("Font", TxtSpin.Font, Ambient.Font)
    End With
    
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TxtSpin,TxtSpin,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = TxtSpin.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set TxtSpin.Font = New_Font
    PropertyChanged "Font"
    RePos
End Property

