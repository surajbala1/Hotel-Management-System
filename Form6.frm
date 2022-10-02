VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form6 
   BackColor       =   &H8000000B&
   Caption         =   "Booking Main"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form6"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command13 
      BackColor       =   &H000000FF&
      Caption         =   "Raise Security and Mail Tickets"
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
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   11040
      Width           =   3975
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H000000FF&
      Caption         =   "Raise Call and Server Tickets"
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
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   10200
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF80&
      Caption         =   "View Full Report"
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
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8160
      Width           =   2535
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
      Height          =   450
      Left            =   3480
      TabIndex        =   34
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFFF80&
      Caption         =   "Exit"
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
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFF80&
      Caption         =   "Search"
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFF80&
      Caption         =   "Last"
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
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFF80&
      Caption         =   "First"
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFF80&
      Caption         =   "Next"
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
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFF80&
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF80&
      Caption         =   "Delete"
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
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Update "
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text6 
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
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   9120
      TabIndex        =   24
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox Text5 
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
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   6480
      TabIndex        =   22
      Top             =   4320
      Width           =   2055
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3480
      TabIndex        =   20
      Text            =   "Select Any"
      Top             =   4320
      Width           =   1815
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   600
      TabIndex        =   18
      Text            =   "Choose Day's"
      Top             =   4320
      Width           =   2535
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   19560
      TabIndex        =   16
      Text            =   "Select Guest Qty"
      Top             =   3000
      Width           =   2415
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
      Height          =   450
      Left            =   17040
      TabIndex        =   14
      Top             =   3000
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   14400
      TabIndex        =   12
      Text            =   "Select Valid ID"
      Top             =   3000
      Width           =   2175
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
      Height          =   450
      Left            =   9120
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
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
      Height          =   450
      Left            =   6480
      TabIndex        =   6
      Top             =   3000
      Width           =   2175
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
      Height          =   420
      Left            =   600
      TabIndex        =   4
      Text            =   "Choose Room No"
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Book New Room"
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Logout"
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
      Left            =   21000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   11880
      TabIndex        =   10
      Top             =   3000
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   50003969
      CurrentDate     =   44145
   End
   Begin VB.Line Line3 
      X1              =   6375
      X2              =   6360
      Y1              =   4920
      Y2              =   9015
   End
   Begin VB.Line Line2 
      X1              =   12600
      X2              =   12615
      Y1              =   4920
      Y2              =   9015
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   12615
      Y1              =   4920
      Y2              =   4935
   End
   Begin VB.Image Image5 
      Height          =   3645
      Left            =   6600
      Picture         =   "Form6.frx":0000
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   5820
   End
   Begin VB.Image Image4 
      Height          =   3615
      Left            =   480
      Picture         =   "Form6.frx":706B
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Guest_Name"
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
      Index           =   10
      Left            =   3480
      TabIndex        =   33
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   5055
      Left            =   16320
      Picture         =   "Form6.frx":E6AB
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   6015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Total_Amount"
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
      Left            =   9120
      TabIndex        =   23
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "See_Per_Day_Rate"
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
      Left            =   6480
      TabIndex        =   21
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Comfort"
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
      Left            =   3480
      TabIndex        =   19
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "How_Many_Day"
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
      Left            =   600
      TabIndex        =   17
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "How_Many_Guest"
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
      Left            =   19560
      TabIndex        =   15
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ID_Number"
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
      Left            =   17040
      TabIndex        =   13
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Gov_Valid_ID"
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
      Left            =   14400
      TabIndex        =   11
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Date_Of_Birth"
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
      Left            =   11880
      TabIndex        =   9
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Address_Full"
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
      Left            =   9120
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mobile_No"
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
      Left            =   6480
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Select_Room"
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
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   6495
      Left            =   360
      Picture         =   "Form6.frx":193EB
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   21975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Welcome To Main Booking Page"
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
      Index           =   1
      Left            =   7080
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   12375
      Left            =   0
      Picture         =   "Form6.frx":19C2B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
Form8.Show
End Sub

Private Sub Command10_Click()
rs.Close
rs.Open "select * from BookDB where Select_Room='" + Combo1.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
Display
Else
MsgBox ("Sorry No Data Found..!!"), vbCritical, "Data Not Found"
End If
End Sub

Private Sub Command12_Click()
Form13.Show
End Sub

Private Sub Command13_Click()
Form14.Show
End Sub

Private Sub Command14_Click()
Form15.Show
End Sub

Private Sub Command2_Click()
Form7.Show
End Sub
Private Sub Command3_Click()
Form10.Show
End Sub

Private Sub Command4_Click()
rs.Fields("Select_Room").Value = Combo1.Text
rs.Fields("Guest_Name").Value = Text1.Text
rs.Fields("Mobile_No").Value = Text2.Text
rs.Fields("Address_Full").Value = Text3.Text
rs.Fields("Date_Of_Birth").Value = DTPicker1.Value
rs.Fields("Gov_Valid_ID").Value = Combo2.Text
rs.Fields("ID_Number").Value = Text4.Text
rs.Fields("How_Many_Guest").Value = Combo3.Text
rs.Fields("How_Many_Day").Value = Combo4.Text
rs.Fields("Comfort").Value = Combo5.Text
rs.Fields("See_Per_Day_Rate").Value = Text5.Text
rs.Fields("Total_Amount").Value = Text6.Text
MsgBox "Data Updated Successfully.....!", vbInformation, "Data Updated"
rs.Update
End Sub

Private Sub Command5_Click()
confirm = MsgBox("Do You Want To Delete", vbYesNo + vbCritical, "Delete Confirmation")
If confirm = vbYes Then
rs.Delete adAffectCurrent
MsgBox ("Your Record Has Been Deleted...!"), vbInformation
rs.Update
Else
MsgBox ("Sorry No Record Deleted!!"), vbCritical, "Data Din't Deleted"
End If
End Sub

Private Sub Command6_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
Display
Else
Display
End If
End Sub
Private Sub Command7_Click()
rs.MoveNext
If Not rs.EOF Then
Display
Else
rs.MoveFirst
Display
End If
End Sub

Private Sub Command8_Click()
rs.MoveFirst
Display
End Sub

Private Sub Command9_Click()
rs.MoveLast
Display
End Sub

Private Sub Form_Load()
Form6.Refresh
con.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=D:\Hotel Management System\Hotel Booking System\Database\Login.mdb;persist security info=false"
rs.Open "select * from BookDB", con, adOpenDynamic, adLockPessimistic

Combo1.AddItem "1"
Combo1.AddItem "2"
Combo1.AddItem "3"
Combo1.AddItem "4"
Combo1.AddItem "5"
Combo1.AddItem "6"
Combo1.AddItem "7"
Combo1.AddItem "8"
Combo1.AddItem "9"
Combo1.AddItem "10"
Combo1.AddItem "11"
Combo1.AddItem "12"
Combo1.AddItem "13"
Combo1.AddItem "14"
Combo1.AddItem "15"
Combo1.AddItem "16"
Combo1.AddItem "17"
Combo1.AddItem "18"
Combo1.AddItem "19"
Combo1.AddItem "20"
Combo1.AddItem "21"
Combo1.AddItem "22"
Combo1.AddItem "23"
Combo1.AddItem "24"
Combo1.AddItem "25"
Combo1.AddItem "26"
Combo1.AddItem "27"
Combo1.AddItem "28"
Combo1.AddItem "29"
Combo1.AddItem "30"
Combo2.AddItem "Pan Card"
Combo2.AddItem "Aadhar Card"
Combo2.AddItem "Driving Licence"
Combo2.AddItem "Voting ID"
Combo3.AddItem "1"
Combo3.AddItem "2"
Combo3.AddItem "3"
Combo3.AddItem "4"
Combo3.AddItem "5"
Combo3.AddItem "6"
Combo3.AddItem "7"
Combo3.AddItem "8"
Combo3.AddItem "9"
Combo3.AddItem "10"
Combo3.AddItem "11"
Combo3.AddItem "11"
Combo3.AddItem "12"
Combo3.AddItem "14"
Combo3.AddItem "15"
Combo4.AddItem "1"
Combo4.AddItem "2"
Combo4.AddItem "3"
Combo4.AddItem "4"
Combo4.AddItem "5"
Combo4.AddItem "6"
Combo4.AddItem "7"
Combo4.AddItem "8"
Combo4.AddItem "9"
Combo4.AddItem "10"
Combo4.AddItem "11"
Combo4.AddItem "12"
Combo4.AddItem "13"
Combo4.AddItem "14"
Combo4.AddItem "15"
Combo4.AddItem "16"
Combo4.AddItem "17"
Combo4.AddItem "18"
Combo4.AddItem "19"
Combo4.AddItem "20"
Combo4.AddItem "21"
Combo4.AddItem "22"
Combo4.AddItem "23"
Combo4.AddItem "24"
Combo4.AddItem "25"
Combo4.AddItem "26"
Combo4.AddItem "27"
Combo4.AddItem "28"
Combo4.AddItem "29"
Combo4.AddItem "30"
Combo5.AddItem "Ac"
Combo5.AddItem "Non Ac"
Display
End Sub
Sub Display()
Combo1.Text = rs!Select_Room
Text1.Text = rs!Guest_Name
Text2.Text = rs!Mobile_No
Text3.Text = rs!Address_Full
DTPicker1.Value = rs!Date_Of_Birth
Combo2.Text = rs!Gov_Valid_ID
Text4.Text = rs!ID_Number
Combo3.Text = rs!How_Many_Guest
Combo4.Text = rs!How_Many_Day
Combo5.Text = rs!Comfort
Text5.Text = rs!See_Per_Day_Rate
Text6.Text = rs!Total_Amount
End Sub
