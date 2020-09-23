VERSION 5.00
Begin VB.Form MyForm 
   BackColor       =   &H00F5F9FA&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2880
   ClientLeft      =   2505
   ClientTop       =   2505
   ClientWidth     =   2775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin mycontrol.AButton CommandButton1 
      Height          =   255
      Left            =   2100
      TabIndex        =   55
      Top             =   2500
      Width           =   560
      _ExtentX        =   979
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "MyForm.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      PATHICON        =   ""
      PICPNG          =   "MyForm.frx":001C
      GCOLOR1         =   0
      GCOLOR2         =   0
      GRADDIR         =   0
   End
   Begin mycontrol.OsenXPComboBox CboMonth 
      Height          =   315
      Left            =   60
      TabIndex        =   54
      Top             =   60
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   60
      ScaleHeight     =   1725
      ScaleWidth      =   2625
      TabIndex        =   11
      Top             =   720
      Width           =   2625
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   2
         Left            =   780
         TabIndex        =   12
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   3
         Left            =   1140
         TabIndex        =   13
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   15
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   5
         Left            =   1860
         TabIndex        =   16
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   6
         Left            =   2220
         TabIndex        =   17
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   4
         Left            =   1500
         TabIndex        =   18
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   7
         Left            =   60
         TabIndex        =   19
         Top             =   330
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   8
         Left            =   420
         TabIndex        =   20
         Top             =   330
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "7"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   9
         Left            =   780
         TabIndex        =   21
         Top             =   330
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   10
         Left            =   1140
         TabIndex        =   22
         Top             =   330
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "9"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   11
         Left            =   1500
         TabIndex        =   23
         Top             =   330
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "10"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   12
         Left            =   1860
         TabIndex        =   24
         Top             =   330
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "11"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   13
         Left            =   2220
         TabIndex        =   25
         Top             =   330
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "12"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   14
         Left            =   60
         TabIndex        =   26
         Top             =   600
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "13"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   15
         Left            =   420
         TabIndex        =   27
         Top             =   600
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "14"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   16
         Left            =   780
         TabIndex        =   28
         Top             =   600
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "15"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   17
         Left            =   1140
         TabIndex        =   29
         Top             =   600
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   12937768
         Value           =   "16"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   18
         Left            =   1500
         TabIndex        =   30
         Top             =   600
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "17"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   19
         Left            =   1860
         TabIndex        =   31
         Top             =   600
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "18"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   20
         Left            =   2220
         TabIndex        =   32
         Top             =   600
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "19"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   21
         Left            =   60
         TabIndex        =   33
         Top             =   870
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "20"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   22
         Left            =   420
         TabIndex        =   34
         Top             =   870
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "21"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   23
         Left            =   780
         TabIndex        =   35
         Top             =   870
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "22"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   24
         Left            =   1140
         TabIndex        =   36
         Top             =   870
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "23"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   25
         Left            =   1500
         TabIndex        =   37
         Top             =   870
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "24"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   26
         Left            =   1860
         TabIndex        =   38
         Top             =   870
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "25"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   27
         Left            =   2220
         TabIndex        =   39
         Top             =   870
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "26"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   28
         Left            =   60
         TabIndex        =   40
         Top             =   1140
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "27"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   29
         Left            =   420
         TabIndex        =   41
         Top             =   1140
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "28"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   30
         Left            =   780
         TabIndex        =   42
         Top             =   1140
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "29"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   31
         Left            =   1140
         TabIndex        =   43
         Top             =   1140
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   "30"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   32
         Left            =   1500
         TabIndex        =   44
         Top             =   1140
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   33
         Left            =   1860
         TabIndex        =   45
         Top             =   1140
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   34
         Left            =   2220
         TabIndex        =   46
         Top             =   1140
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   35
         Left            =   60
         TabIndex        =   47
         Top             =   1410
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   36
         Left            =   420
         TabIndex        =   48
         Top             =   1410
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   37
         Left            =   780
         TabIndex        =   49
         Top             =   1410
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   38
         Left            =   1140
         TabIndex        =   50
         Top             =   1410
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   39
         Left            =   1500
         TabIndex        =   51
         Top             =   1410
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   40
         Left            =   1860
         TabIndex        =   52
         Top             =   1410
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin mycontrol.LbDate LbDate1 
         Height          =   255
         Index           =   41
         Left            =   2220
         TabIndex        =   53
         Top             =   1410
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         BackHot         =   16777215
         ForeColor       =   0
         Value           =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00B99D7F&
         Height          =   1725
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   2625
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DC9670&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   60
      ScaleHeight     =   285
      ScaleWidth      =   2625
      TabIndex        =   3
      Top             =   420
      Width           =   2625
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sun"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   30
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mon"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   450
         TabIndex        =   9
         Top             =   30
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tue"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   8
         Top             =   30
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wed"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   1170
         TabIndex        =   7
         Top             =   30
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thu"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   1560
         TabIndex        =   6
         Top             =   30
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fri"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   1950
         TabIndex        =   5
         Top             =   30
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sat"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   2280
         TabIndex        =   4
         Top             =   30
         Width           =   240
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00B99D7F&
         Height          =   285
         Left            =   0
         Top             =   0
         Width           =   2625
      End
   End
   Begin mycontrol.OsenXPSpin Spin 
      Height          =   330
      Left            =   1890
      TabIndex        =   2
      Top             =   60
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   582
      Text            =   "2003"
      Max             =   2100
      Min             =   1900
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "MyForm.frx":0034
      Top             =   2490
      Width           =   345
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00B99D7F&
      Height          =   2835
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   2745
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Today: "
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   420
      TabIndex        =   1
      Top             =   2520
      Width           =   555
   End
   Begin VB.Label LbNow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2003/07/13"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   2520
      Width           =   1020
   End
End
Attribute VB_Name = "MyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public InitValue As Long

Private Sub CboMonth_Change()
    SetDisplay CboMonth.ListIndex + 1, Spin.Value
End Sub

Private Sub cbomonth_Click()
    SetDisplay CboMonth.ListIndex + 1, Spin.Value
End Sub

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
Dim I As Integer, J As Integer

    I = Format(InitValue, "mm")
    J = Format(InitValue, "yyyy")
    DoEvents
    
    CboMonth.ListIndex = I - 1
    Spin.Value = J
    DoEvents
    
    SetDisplay I, J
    
    LbNow.Caption = GetDateSys
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    InitMonth
    Hide
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Image1_Click()
    ResultDate = DateSerial(Year(Now), Month(Now), Day(Now))
    Unload Me
End Sub

Private Sub LbDate1_Click(Index As Integer)
    ResultDate = LbDate1(Index).Tanggal
    Unload Me
End Sub

Private Sub LbDate1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Integer
    If LbDate1(Index).Value = "" Then Exit Sub
    For I = 0 To 41
        LbDate1(I).SetPicture False
        If I <> Index And LbDate1(I).ForeColor <> &HC0C0C0 Then
            LbDate1(I).BackColor = vbWhite
            LbDate1(I).FontBold = False
            LbDate1(I).ForeColor = vbBlack
        ElseIf LbDate1(I).ForeColor <> &HC0C0C0 Then
            LbDate1(I).BackColor = &HC56A28
            LbDate1(I).FontBold = True
            LbDate1(I).ForeColor = vbWhite
        End If
    Next
End Sub

Public Sub SetDisplay(ByVal Bulan As Integer, ByVal Tahun As Integer)
Dim StrTgl As String, StrBln As String, StrThn As String
Dim I As Integer, N As Integer, X As Long

On Error GoTo ErrX

StrTgl = Tahun & "/" & Format(Bulan, "00") & "/01"
X = Format(StrTgl, "#")
StrTgl = Format(StrTgl, "ddd")

    Select Case UCase(StrTgl)
        Case "MON": I = 1
        Case "TUE": I = 2
        Case "WED": I = 3
        Case "THU": I = 4
        Case "FRI": I = 5
        Case "SAT": I = 6
        Case "SUN": I = 7
    End Select
    
X = X - I

For N = 0 To 41

    LbDate1(N).Value = Format(Format(X + N, "DD"), "#")
    LbDate1(N).FontBold = False
    LbDate1(N).ForeColor = vbBlack
    LbDate1(N).BackColor = vbWhite
    LbDate1(N).Tanggal = X + N
    
    If Val(Format(X + N, "MM")) > Bulan Or Val(Format(N + X, "MM")) < Bulan Then
        LbDate1(N).ForeColor = &HC0C0C0
    End If
    
    StrTgl = DateSys(Str(X + N))
    
    If StrTgl = GetDateSys Then
        LbDate1(N).SetPicture
    Else
        LbDate1(N).SetPicture False
        
        If StrTgl = DateSys(Str(InitValue)) Then
            LbDate1(N).SetPicture False
            LbDate1(N).FontBold = True
            LbDate1(N).ForeColor = vbWhite
            LbDate1(N).BackColor = &HC56A28
        End If
        
    End If
    

Next
ErrX:
End Sub

Private Sub Spin_Change()
On Error Resume Next
    SetDisplay CboMonth.ListIndex + 1, Spin.Value
End Sub

Sub InitMonth()
Dim I As Integer, StrA As String
    With CboMonth
        .Clear
        For I = 1 To 12
            StrA = "2003/" & Format(I, "00") & "/01"
            .AddItem Format(StrA, "mmmm")
        Next
    End With
End Sub










