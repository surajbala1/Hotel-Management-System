VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Contact"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "<- Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Write Tech-Support@Support.com For Mailing Issue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   14520
      TabIndex        =   4
      Top             =   11400
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "Call 1234567890 For Strong Security"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   3
      Top             =   11400
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Call 1234567890 For Server Related Issue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   15000
      TabIndex        =   2
      Top             =   5640
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Call 1234567890 For Call Support"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Top             =   5640
      Width           =   4215
   End
   Begin VB.Image Image5 
      Height          =   4395
      Left            =   13920
      Picture         =   "Form4.frx":0000
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   7500
   End
   Begin VB.Image Image4 
      Height          =   4365
      Left            =   2760
      Picture         =   "Form4.frx":416A
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   7515
   End
   Begin VB.Image Image3 
      Height          =   4320
      Left            =   13920
      Picture         =   "Form4.frx":573C
      Top             =   1080
      Width           =   7500
   End
   Begin VB.Image Image2 
      Height          =   4320
      Left            =   2760
      Picture         =   "Form4.frx":F300
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   7500
   End
   Begin VB.Image Image1 
      Height          =   12375
      Left            =   0
      Picture         =   "Form4.frx":10A9F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form1.Show
End Sub
