VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "About Hotel Booking"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
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
      Height          =   540
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "You Can Book Your Room Before Visit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   6
      Top             =   9600
      Width           =   5415
   End
   Begin VB.Label Label3 
      Caption         =   "Cancellation Any Time With Zero Charge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   9480
      TabIndex        =   5
      Top             =   4920
      Width           =   6015
   End
   Begin VB.Label Label3 
      Caption         =   "Bulk Reservation Available"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   9480
      TabIndex        =   4
      Top             =   6480
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Booking Facality For Party And Bulk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   9360
      TabIndex        =   3
      Top             =   8040
      Width           =   5295
   End
   Begin VB.Label Label2 
      Caption         =   "24 x 7 Services Available"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Welcome To Hotel Booking System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   0
      Top             =   600
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   12375
      Left            =   0
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
End Sub
