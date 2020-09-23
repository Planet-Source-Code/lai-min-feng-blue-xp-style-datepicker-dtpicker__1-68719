VERSION 5.00
Begin VB.UserControl OsenXPComboBox 
   BackColor       =   &H00D8E9EC&
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   EditAtDesignTime=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   5175
   ToolboxBitmap   =   "ComboXP.ctx":0000
   Begin VB.PictureBox BackMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   2445
      TabIndex        =   0
      Top             =   0
      Width           =   2445
      Begin VB.PictureBox ImgUp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         Picture         =   "ComboXP.ctx":0312
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   2
         Top             =   30
         Width           =   255
      End
      Begin VB.TextBox TxtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   75
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   60
         Width           =   720
      End
      Begin VB.Shape ShapeBorder 
         BorderColor     =   &H00B99D7F&
         Height          =   315
         Left            =   0
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   510
      Width           =   2295
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   0
      Left            =   150
      Picture         =   "ComboXP.ctx":06C8
      Top             =   900
      Width           =   255
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   1
      Left            =   150
      Picture         =   "ComboXP.ctx":0A7E
      Top             =   1170
      Width           =   255
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   2
      Left            =   150
      Picture         =   "ComboXP.ctx":0E34
      Top             =   1470
      Width           =   255
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   3
      Left            =   150
      Picture         =   "ComboXP.ctx":11EA
      Top             =   1800
      Width           =   255
   End
End
Attribute VB_Name = "OsenXPComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const CB_SHOWDROPDOWN = &H14F
Const CB_GETDROPPEDSTATE = &H157
Event Click() 'MappingInfo=TxtData,TxtData,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event Change() 'MappingInfo=TxtData,TxtData,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=TxtData,TxtData,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=TxtData,TxtData,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
'Default Property Values:
Const m_def_Enabled = 0
'Property Variables:
Dim m_Enabled As Boolean

Public Sub OpenCombo(chwnd As Long)
    Dim rc As Long
    rc = SendMessage(chwnd, CB_GETDROPPEDSTATE, 0, 0)
    If rc = 0 Then
        SendMessage chwnd, CB_SHOWDROPDOWN, True, 0
    Else
        SendMessage chwnd, CB_SHOWDROPDOWN, False, 0
    End If
End Sub
'Event Declarations:

Private Sub RePos()
Dim I As Integer
    If Width < 400 Then Width = 400
    ShapeBorder.Width = Width
    ImgUp.Left = Width - 285
    BackMain.Width = Width
    
    With Combo1
        .Top = 30
        .Left = 0
        .Width = Width
    End With
    
    Height = 315
    
    With TxtData
        .Width = Width - 375
        .Top = 60
        If .FontSize > 8 Then
            I = .FontSize - 8
            I = I * 15
            .Top = .Top - I
        End If
    End With
    
End Sub

Private Sub Combo1_Click()
    TxtData = Combo1.Text
End Sub

Private Sub ImgUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ImgUp.Picture = Img(2).Picture
    OpenCombo Combo1.hWnd
    
End Sub

Private Sub ImgUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If ImgUp.Picture <> Img(1).Picture Then ImgUp.Picture = Img(1).Picture

End Sub

Private Sub TxtData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    ResetPic
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetPic
End Sub

Private Sub UserControl_Resize()
    RePos
End Sub
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = TxtData.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    TxtData.Text() = New_Text
    PropertyChanged "Text"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    TxtData.Text = PropBag.ReadProperty("Text", "0")
    TxtData.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    Set TxtData.Font = PropBag.ReadProperty("Font", Ambient.Font)
    TxtData.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Combo1.ListIndex = PropBag.ReadProperty("ListIndex", -1)
    TxtData.Locked = PropBag.ReadProperty("Locked", False)
    TxtData.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    TxtData.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    TxtData.DataField = PropBag.ReadProperty("FieldName", "")

    RePos
    
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Text", TxtData.Text, "0")
    Call PropBag.WriteProperty("BackColor", TxtData.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Font", TxtData.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", TxtData.ForeColor, &H80000008)
    Call PropBag.WriteProperty("ListIndex", Combo1.ListIndex, -1)
    Call PropBag.WriteProperty("Locked", TxtData.Locked, False)
    Call PropBag.WriteProperty("MaxLength", TxtData.MaxLength, 0)
    Call PropBag.WriteProperty("ToolTipText", TxtData.ToolTipText, "")
    Call PropBag.WriteProperty("FieldName", TxtData.DataField, "")
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
End Sub

Sub ResetPic()
    If ImgUp.Picture <> Img(0).Picture Then
        ImgUp.Picture = Img(0).Picture
    End If
End Sub
Private Sub TxtData_Click()
    RaiseEvent Click
End Sub
Private Sub TxtData_Change()
    RaiseEvent Change
End Sub
Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
    Combo1.AddItem Item, Index
End Sub
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = TxtData.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    TxtData.BackColor() = New_BackColor
    BackMain.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property
Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of a control or the system Clipboard."
    Combo1.Clear
End Sub
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = TxtData.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set TxtData.Font = New_Font
    RePos
    PropertyChanged "Font"
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = TxtData.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    TxtData.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = Combo1.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    Combo1.ListIndex() = New_ListIndex
    If Combo1.ListIndex > -1 Then TxtData.Text = Combo1.Text
    PropertyChanged "ListIndex"
End Property
Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = Combo1.ListCount
End Property
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = TxtData.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    TxtData.Locked() = New_Locked
    PropertyChanged "Locked"
End Property
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = TxtData.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    TxtData.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property
Private Sub TxtData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgUp_MouseDown Button, Shift, X, Y
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = TxtData.ToolTipText
End Property
Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    TxtData.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property
Public Property Get FieldName() As String
Attribute FieldName.VB_Description = "Returns/sets a value that describes the DataMember for a data connection."
    FieldName = TxtData.DataField
End Property

Public Property Let FieldName(ByVal New_FieldName As String)
    TxtData.DataField() = New_FieldName
    PropertyChanged "FieldName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    TxtData.Enabled = New_Enabled
    If New_Enabled = False Then
        ImgUp.Picture = Img(3).Picture
        ShapeBorder.BorderColor = &HC0C0C0
    Else
        ResetPic
        ShapeBorder.BorderColor = &HB99D7F
    End If
    
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
End Sub

