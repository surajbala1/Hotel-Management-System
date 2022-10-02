VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Home"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF80&
      Caption         =   "More.."
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
      Left            =   20520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Support"
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
      Left            =   18000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF80&
      Caption         =   "About"
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
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Employee Login"
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Back"
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Developed By: Omkar Khengare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   11520
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "STAY HAPPY ! Keep Smiling ! "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9000
      TabIndex        =   6
      Top             =   9840
      Width           =   3735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   7
      Index           =   1
      X1              =   7200
      X2              =   7215
      Y1              =   2400
      Y2              =   9495
   End
   Begin VB.Image Image4 
      Height          =   6615
      Left            =   14400
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   8295
   End
   Begin VB.Image Image3 
      Height          =   6690
      Left            =   7440
      Picture         =   "Form1.frx":975B
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   6540
   End
   Begin VB.Image Image2 
      Height          =   6675
      Left            =   120
      Picture         =   "Form1.frx":18DAB
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   6900
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   7
      Index           =   0
      X1              =   14160
      X2              =   14175
      Y1              =   2400
      Y2              =   9495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   6
      Index           =   1
      X1              =   0
      X2              =   22815
      Y1              =   9480
      Y2              =   9495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   6
      Index           =   0
      X1              =   0
      X2              =   22815
      Y1              =   2400
      Y2              =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Welcome To Hotel Booking Portal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   12375
      Left            =   0
      Picture         =   "Form1.frx":262E1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
confirm = MsgBox("Your Login Session Will Be Destroyed", vbYesNo + vbCritical, "Session Confirmation")
If confirm = vbYes Then
MsgBox ("Ok We Will Taking You Admin Portal"), vbInformation, "Good To Go"
Form15.Show
Else
MsgBox ("Ohhh !!"), vbInformation, "Now You On Your Portal"
End If
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
Form3.Show
End Sub

Private Sub Command4_Click()
Form4.Show
End Sub

Private Sub Command5_Click()
Form5.Show
End Sub
