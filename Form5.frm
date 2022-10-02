VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "More.."
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Various Food Available For Guest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   9480
      TabIndex        =   14
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000F&
      BorderWidth     =   3
      X1              =   7680
      X2              =   15120
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      Index           =   2
      X1              =   7680
      X2              =   15375
      Y1              =   3360
      Y2              =   3375
   End
   Begin VB.Image Image3 
      Height          =   3855
      Left            =   7920
      Picture         =   "Form5.frx":0000
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   6975
   End
   Begin VB.Image Image2 
      Height          =   3825
      Left            =   7920
      Picture         =   "Form5.frx":1661D
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   7005
   End
   Begin VB.Label Label3 
      Caption         =   "Special Security Force Available"
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
      Index           =   9
      Left            =   15960
      TabIndex        =   13
      Top             =   9960
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Guest May Allow For Direservation Folks "
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
      Index           =   8
      Left            =   15960
      TabIndex        =   12
      Top             =   8520
      Width           =   4935
   End
   Begin VB.Label Label3 
      Caption         =   "Tour Also Providing As Per Guest Request"
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
      Index           =   7
      Left            =   15960
      TabIndex        =   11
      Top             =   6960
      Width           =   5175
   End
   Begin VB.Label Label3 
      Caption         =   "Sepecial Request Guest Can make"
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
      Index           =   6
      Left            =   15960
      TabIndex        =   10
      Top             =   5280
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   " Guest Can Reserve Advance For More Days"
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
      Index           =   5
      Left            =   15960
      TabIndex        =   9
      Top             =   3840
      Width           =   5655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Some Benifits For Booking"
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
      Left            =   17640
      TabIndex        =   8
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      Index           =   1
      X1              =   15120
      X2              =   22815
      Y1              =   3360
      Y2              =   3375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      Index           =   1
      X1              =   15120
      X2              =   15135
      Y1              =   3360
      Y2              =   12375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      Index           =   0
      X1              =   7680
      X2              =   7695
      Y1              =   3360
      Y2              =   12375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   7695
      Y1              =   3360
      Y2              =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Need Indian Citizenship For Normal Booking"
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
      Index           =   4
      Left            =   600
      TabIndex        =   7
      Top             =   9960
      Width           =   5415
   End
   Begin VB.Label Label3 
      Caption         =   "NRI Guest May Required Some Special Document"
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
      Left            =   600
      TabIndex        =   6
      Top             =   8520
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "Required Gov Valid ID For Reservation"
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
      Left            =   600
      TabIndex        =   5
      Top             =   6960
      Width           =   4815
   End
   Begin VB.Label Label3 
      Caption         =   "You Can Cancelled Your Reservation Within 24Hrs"
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
      Left            =   600
      TabIndex        =   4
      Top             =   5280
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "Various Payment Method Online / Offline"
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
      Left            =   600
      TabIndex        =   3
      Top             =   3840
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Room Booking / Reservation"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Some Detail's Are Below About Booking !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   960
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   12375
      Left            =   -3360
      Picture         =   "Form5.frx":2196A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   26175
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
End Sub
