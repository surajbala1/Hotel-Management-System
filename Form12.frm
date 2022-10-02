VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00000000&
   Caption         =   "Data Storation For Local Server"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form12"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5880
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
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
      Height          =   420
      Left            =   10200
      TabIndex        =   10
      Text            =   "Choose One"
      Top             =   5160
      Width           =   2055
   End
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
      Height          =   420
      Left            =   13560
      TabIndex        =   9
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text3 
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
      Left            =   9120
      TabIndex        =   8
      Top             =   4200
      Width           =   1695
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
      Height          =   420
      Left            =   13560
      TabIndex        =   7
      Top             =   3240
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
      Height          =   420
      Left            =   9120
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000004&
      X1              =   4320
      X2              =   4800
      Y1              =   1200
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000004&
      X1              =   4800
      X2              =   18015
      Y1              =   1560
      Y2              =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      X1              =   4800
      X2              =   4815
      Y1              =   1560
      Y2              =   9975
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000002&
      Caption         =   "Developed By:- Omkar Khengare"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   11040
      Width           =   4575
   End
   Begin VB.Label Label8 
      Caption         =   "Please Dont Share Your Valid Credential With Any One"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   14
      Top             =   8880
      Width           =   7695
   End
   Begin VB.Label Label7 
      Caption         =   "Please Enter Your User ID And Password As Unique"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   13
      Top             =   8040
      Width           =   7335
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000002&
      Caption         =   ":Note:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10920
      TabIndex        =   12
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label Label5 
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
      Left            =   11520
      TabIndex        =   5
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label4 
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
      Left            =   7440
      TabIndex        =   4
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label3 
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
      Left            =   11520
      TabIndex        =   3
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label2 
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
      Left            =   7440
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      Caption         =   "Please Fill All Detail's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   0
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   8835
      Left            =   4320
      Picture         =   "Form12.frx":0000
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   13740
   End
   Begin VB.Image Image1 
      Height          =   12375
      Left            =   0
      Picture         =   "Form12.frx":1629
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22860
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command2_Click()
rs.AddNew
rs.Fields("Enter_Name").Value = Text1.Text
rs.Fields("Enter_Mobile_No").Value = Text2.Text
rs.Fields("User_ID").Value = Text3.Text
rs.Fields("Enter_Password").Value = Text4.Text
rs.Fields("Type").Value = Combo1.Text
MsgBox "Congratulation Your New User Id & Password Has Been Created", vbInformation, "Congratualtion"
rs.Update
Form2.Show

End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=D:\Hotel Management System\Hotel Booking System\Database\Login.mdb;persist security info=false"
rs.Open "select * from User", con, adOpenDynamic, adLockPessimistic

Combo1.AddItem "User"
Combo1.AddItem "Admin"
End Sub

