VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "Login"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc loginado 
      Height          =   330
      Left            =   -1680
      Top             =   12000
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form2.frx":0000
      OLEDBString     =   $"Form2.frx":008F
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from User1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
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
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Clear"
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
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Login"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1335
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
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   9240
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3960
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
      Height          =   375
      Left            =   9240
      TabIndex        =   2
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000002&
      Caption         =   "Create New ID"
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
      Left            =   10080
      TabIndex        =   11
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000002&
      Caption         =   "Forget Password?"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   10
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   4560
      X2              =   16095
      Y1              =   6360
      Y2              =   6375
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Read Carefully"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   2
      Left            =   6000
      TabIndex        =   9
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "If You Forgot Your User Id And Password The Contact Administration"
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
      TabIndex        =   8
      Top             =   8640
      Width           =   8415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Please Enter Your Valid Credential"
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
      TabIndex        =   7
      Top             =   7680
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Password"
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
      Left            =   7200
      TabIndex        =   1
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "User Id"
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
      Left            =   7200
      TabIndex        =   0
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   6135
      Left            =   4560
      Picture         =   "Form2.frx":011E
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   11535
   End
   Begin VB.Image Image1 
      Height          =   12375
      Left            =   0
      Picture         =   "Form2.frx":871D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
loginado.RecordSource = "select * from User1 where User_ID='" + Text1.Text + "'and Enter_Password='" + Text2.Text + "'"
loginado.Refresh
If loginado.Recordset.EOF Then
MsgBox "Sorry Credential Not Matched..! Try With Valid Credential", vbCritical, "Incorrect User ID And Password"
Text1.Text = Clear
Text2.Text = Clear
Else
MsgBox "Congratulation Login Successfully..!!", vbInformation, "Credential Matched"
Text1.Text = Clear
Text2.Text = Clear
Form6.Show
End If
End Sub

Private Sub Command2_Click()
Text1.Text = Clear
Text2.Text = Clear
MsgBox "Data Cleared", vbInformation, "Ohhh!!"
End Sub
Private Sub Command3_Click()
Form1.Show
End Sub

Private Sub Label4_Click()
Form11.Show
End Sub

Private Sub Label5_Click()
Form12.Show
End Sub
