VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "Password Recovery"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form11"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   14520
      TabIndex        =   15
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   10440
      TabIndex        =   14
      Top             =   3960
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   11400
      TabIndex        =   7
      Text            =   "Choose One"
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Submit"
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
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   14520
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   10440
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      Index           =   1
      X1              =   12360
      X2              =   12360
      Y1              =   2280
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      Index           =   0
      X1              =   12240
      X2              =   12240
      Y1              =   2280
      Y2              =   4680
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000002&
      Caption         =   "Enter Mobile No"
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
      Left            =   12480
      TabIndex        =   13
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000002&
      Caption         =   "Enter Name"
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
      Left            =   8880
      TabIndex        =   12
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000002&
      Caption         =   "Developed By:- Omkar Khengare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   11280
      Width           =   4095
   End
   Begin VB.Label Label5 
      Caption         =   "Please Choose One Bcoz(Single Server Multiple Login)"
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
      Left            =   8880
      TabIndex        =   10
      Top             =   9240
      Width           =   6615
   End
   Begin VB.Label Label5 
      Caption         =   "This Data Is Not Recover By Main Server"
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
      Left            =   9720
      TabIndex        =   9
      Top             =   8400
      Width           =   5055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":Please Note:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   11640
      TabIndex        =   8
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000002&
      Caption         =   "Single Portal For Multiple Recovery"
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
      Left            =   10200
      TabIndex        =   4
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000002&
      Caption         =   "Enter Password"
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
      Left            =   12480
      TabIndex        =   1
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      Caption         =   "User ID"
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
      Left            =   8880
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   9675
      Left            =   3960
      Picture         =   "Form11.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   16620
   End
   Begin VB.Image Image1 
      Height          =   12435
      Left            =   0
      Picture         =   "Form11.frx":183D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22770
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
rs.AddNew
rs.Fields("Enter_Name").Value = Text1.Text
rs.Fields("Enter_Mobile_No").Value = Text2.Text
rs.Fields("User_ID").Value = Text3.Text
rs.Fields("Enter_Password").Value = Text4.Text
rs.Fields("Type").Value = Combo1.Text
MsgBox "Your user ID & Password Has Been Recovered", vbInformation, "Recovered"
rs.Update
Form2.Show
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=D:\Hotel Management System\Hotel Booking System\Database\Login.mdb;persist security info=false"
rs.Open "select * from User", con, adOpenDynamic, adLockPessimistic

Combo1.AddItem "User"
Combo1.AddItem "Admin"
End Sub
