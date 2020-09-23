VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl MFDatePicker 
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   ScaleHeight     =   2160
   ScaleWidth      =   2190
   Begin VB.PictureBox ImgUp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1185
      Picture         =   "MFDatePicker.ctx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM dd, yyyy"
      Format          =   57868291
      CurrentDate     =   39227
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   1
      Left            =   1320
      Picture         =   "MFDatePicker.ctx":03B6
      Top             =   720
      Width           =   255
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   2
      Left            =   1320
      Picture         =   "MFDatePicker.ctx":076C
      Top             =   1020
      Width           =   255
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   3
      Left            =   1320
      Picture         =   "MFDatePicker.ctx":0B22
      Top             =   1350
      Width           =   255
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   0
      Left            =   1320
      Picture         =   "MFDatePicker.ctx":0ED8
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape ShapeBorder 
      BorderColor     =   &H00B99D7F&
      Height          =   315
      Left            =   -15
      Top             =   -15
      Width           =   1470
   End
End
Attribute VB_Name = "MFDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Event Declarations:
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=DTPicker1,DTPicker1,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=DTPicker1,DTPicker1,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=DTPicker1,DTPicker1,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event Change() 'MappingInfo=DTPicker1,DTPicker1,-1,Change
Attribute Change.VB_Description = "Occurs when the user selects a new date or changes a date in the edit portion of the control."
Event Click() 'MappingInfo=DTPicker1,DTPicker1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=DTPicker1,DTPicker1,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while the mousepointer is over an object."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=DTPicker1,DTPicker1,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mousepointer over an object."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=DTPicker1,DTPicker1,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while the mousepointer is over an object."
Event Resize()
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Private WithEvents TestForm As Form
Attribute TestForm.VB_VarHelpID = -1

Private Sub DTPicker1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    ResetPic
End Sub

Private Sub TestForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If ParentHwnd = hWnd Then
        DTPicker1.Value = ResultDate
    End If
End Sub

Private Sub ImgUp_Click()
On Error GoTo ErrXYZ
    Dim MyApi As POINTAPI
    Dim MyRect As RECT
    Dim MyLeft As Long, MyTop As Long
    
    If DTPicker1.Enabled Then
        GetCursorPos MyApi
        GetWindowRect ImgUp.hWnd, MyRect
        
        MyTop = (MyRect.Top * 15) + ImgUp.Height + 45
        If Screen.Height - MyTop < 3000 Then
            MyTop = MyTop - 2835 - 30 - Height
        End If
        
        MyLeft = (MyRect.Left * 15) - ImgUp.Left
        If Screen.Width - MyLeft < 2900 Then
            MyLeft = (MyRect.Left * 15) + ImgUp.Width + 45 - 2760
        End If
        
        If Value = 0 Then Value = Val(Format(Date, "#"))
        
        ResultDate = DTPicker1.Value
        DoEvents
        
        Set TestForm = MyForm
        ParentHwnd = hWnd
        With TestForm
            .InitValue = Value
            DoEvents
            .Left = MyLeft
            .Top = MyTop
            .Height = 2835
            .Width = 2745
            .Show
        End With
    End If
    Exit Sub
    
ErrXYZ:
    MsgBox err.Description
End Sub
'
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = DTPicker1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    DTPicker1.Enabled() = New_Enabled
    ImgUp.Enabled = New_Enabled
    If New_Enabled = False Then
        ImgUp.Picture = Img(3).Picture
        ShapeBorder.BorderColor = &HC0C0C0
    Else
        ResetPic
        ShapeBorder.BorderColor = &HB99D7F
    End If
    PropertyChanged "Enabled"
End Property

Private Sub ImgUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DTPicker1.Enabled = True Then
        ImgUp.Picture = Img(2).Picture
        ImgUp_Click
    End If
End Sub

Private Sub ImgUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DTPicker1.Enabled = True Then
        If ImgUp.Picture <> Img(1).Picture Then ImgUp.Picture = Img(1).Picture
    End If
End Sub

Private Sub TestForm_Unload(Cancel As Integer)
    ResetPic
End Sub

Private Sub UserControl_InitProperties()
    DTPicker1.Value = Now
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetPic
End Sub

Sub ResetPic()
    If ImgUp.Picture <> Img(0).Picture And DTPicker1.Enabled Then
        ImgUp.Picture = Img(0).Picture
    End If
End Sub

Private Sub UserControl_Resize()
    Height = 315
    ImgUp.Top = 30
    ImgUp.Left = Width - 285
    DTPicker1.Width = Width
    ShapeBorder.Width = Width
    ShapeBorder.Height = Height
End Sub

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=DTPicker1,DTPicker1,-1,Enabled
'Public Property Get Enabled() As Boolean
'    Enabled = DTPicker1.Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    DTPicker1.Enabled() = New_Enabled
'    If New_Enabled = False Then ImgUp.Picture = Img(2).Picture Else ImgUp.Picture = Img(2).Picture
'    PropertyChanged "Enabled"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DTPicker1,DTPicker1,-1,CustomFormat
Public Property Get CustomFormat() As String
Attribute CustomFormat.VB_Description = "Returns/sets the custom format string used to format the date and/or time displayed in the control."
    CustomFormat = DTPicker1.CustomFormat
End Property

Public Property Let CustomFormat(ByVal New_CustomFormat As String)
    DTPicker1.CustomFormat() = New_CustomFormat
    PropertyChanged "CustomFormat"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DTPicker1,DTPicker1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = DTPicker1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set DTPicker1.Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub DTPicker1_Change()
    RaiseEvent Change
End Sub

Private Sub DTPicker1_Click()
    RaiseEvent Click
End Sub

Private Sub DTPicker1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub DTPicker1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DTPicker1,DTPicker1,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = DTPicker1.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set DTPicker1.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DTPicker1,DTPicker1,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = DTPicker1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    DTPicker1.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12
Public Function ScaleY(ByVal Height As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
Attribute ScaleY.VB_Description = "Converts the value for the height of a Form, PictureBox, or Printer from one unit of measure to another."

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12
Public Function ScaleX(ByVal Width As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
Attribute ScaleX.VB_Description = "Converts the value for the width of a Form, PictureBox, or Printer from one unit of measure to another."

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DTPicker1,DTPicker1,-1,Value
Public Property Get Value() As Variant
Attribute Value.VB_Description = "Returns/sets the current date."
    Value = DTPicker1.Value
End Property

Public Property Let Value(ByVal New_Value As Variant)
    DTPicker1.Value() = New_Value
    PropertyChanged "Value"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    DTPicker1.Enabled = PropBag.ReadProperty("Enabled", True)
    DTPicker1.CustomFormat = PropBag.ReadProperty("CustomFormat", "MMM dd, yyyy")
    Set DTPicker1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    DTPicker1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    DTPicker1.Value = PropBag.ReadProperty("Value", 5 / 25 / 2007)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", DTPicker1.Enabled, True)
    Call PropBag.WriteProperty("CustomFormat", DTPicker1.CustomFormat, "MMM dd, yyyy")
    Call PropBag.WriteProperty("Font", DTPicker1.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", DTPicker1.MousePointer, 0)
    Call PropBag.WriteProperty("Value", DTPicker1.Value, 5 / 25 / 2007)
End Sub

