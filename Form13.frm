VERSION 5.00
Begin VB.Form Form13 
   Caption         =   "Complaint Raising Area"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form13"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "New Tickets"
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3000
      Width           =   1575
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
      TabIndex        =   11
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Raise Tickets"
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8040
      Width           =   2175
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
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   6
      Top             =   6360
      Width           =   5655
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
      Left            =   9480
      TabIndex        =   5
      Text            =   "Please Select"
      Top             =   5400
      Width           =   2055
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
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   9480
      TabIndex        =   4
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Click Here First Before Raise New Tickets"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   14
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000FF&
      Caption         =   "Click 'NEW Tickets' Before Raise New Tickets Otherwise Your Existing Data Will Be Replace"
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
      Left            =   6720
      TabIndex        =   12
      Top             =   2040
      Width           =   9615
   End
   Begin VB.Image Image3 
      Height          =   3495
      Left            =   12000
      Picture         =   "Form13.frx":0000
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   5175
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000002&
      Caption         =   "Expected Resolution Within 24 Hrs"
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
      Left            =   5040
      TabIndex        =   10
      Top             =   9600
      Width           =   4215
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000002&
      Caption         =   "Please Enter Actual Problem In Description Box"
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
      Left            =   5040
      TabIndex        =   9
      Top             =   8760
      Width           =   5775
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   8160
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000002&
      Caption         =   "Describe Your Problem"
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
      Left            =   6120
      TabIndex        =   3
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000002&
      Caption         =   "Which Section ?"
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
      Left            =   6120
      TabIndex        =   2
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000002&
      Caption         =   "Enter Your User ID"
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
      Left            =   6120
      TabIndex        =   1
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Please Raise Actual Problem"
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
      Left            =   7560
      TabIndex        =   0
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   8175
      Left            =   4920
      Picture         =   "Form13.frx":179F
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   12495
   End
   Begin VB.Image Image1 
      Height          =   12375
      Left            =   0
      Picture         =   "Form13.frx":4657
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
rs.Fields("User_ID").Value = Text1.Text
rs.Fields("Which_Section").Value = Combo1.Text
rs.Fields("Describe_Your_Problem").Value = Text2.Text
MsgBox "System Record Your Tickets....Expected Resolution Within 24 Hrs....Have A Good Day...!!!", vbInformation, "Tickets Raised"
rs.Update
Form6.Show

End Sub
Private Sub Command2_Click()
Form6.Show
End Sub
Sub Clear()
Text1.Text = ""
Combo1.Text = "Please Select"
Text2.Text = ""
End Sub

Private Sub Command3_Click()
rs.AddNew
Clear
MsgBox "Now You Good To Fill..!", vbInformation, "Please Fill All Field's"
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=D:\Hotel Management System\Hotel Booking System\Database\Login.mdb;persist security info=false"
rs.Open "select * from Call", con, adOpenDynamic, adLockPessimistic

Combo1.AddItem "Call Issue"
Combo1.AddItem "Server Issue"
End Sub
