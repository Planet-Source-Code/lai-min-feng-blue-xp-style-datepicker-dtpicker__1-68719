VERSION 5.00
Object = "*\AMyControl.vbp"
Begin VB.Form Form1 
   Caption         =   "Demo"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2085
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   2085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdGetDate 
      Caption         =   "Get Date"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin MyControl.MFDatePicker MFDatePicker1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Form1.frx":0000
      Value           =   39234.5389583333
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGetDate_Click()
    MsgBox "Chosen Date : " & Format(MFDatePicker1.Value, "MMM dd,yyyy")
End Sub
