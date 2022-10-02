VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form7 
   Caption         =   "Please Add Information For Server"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Book New"
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
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Book Room"
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
      Left            =   15600
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Calculate Amount"
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
      Left            =   17400
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5160
      Width           =   2295
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
      Height          =   420
      Left            =   17520
      TabIndex        =   28
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Calulate P/D Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2640
      Width           =   2295
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
      Left            =   17520
      TabIndex        =   23
      Top             =   2040
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
      Left            =   11760
      TabIndex        =   21
      Text            =   "Select Any"
      Top             =   7080
      Width           =   2535
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
      Left            =   11760
      TabIndex        =   20
      Text            =   "Choose Day's"
      Top             =   5880
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
      Left            =   11760
      TabIndex        =   19
      Text            =   "Select Guest Qty"
      Top             =   4560
      Width           =   2535
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
      Height          =   375
      Left            =   11760
      TabIndex        =   18
      Top             =   3240
      Width           =   2535
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
      Left            =   11760
      TabIndex        =   17
      Text            =   "Select Valid ID"
      Top             =   2040
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   5520
      TabIndex        =   16
      Top             =   7080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
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
      Height          =   495
      Left            =   5520
      TabIndex        =   15
      Top             =   5880
      Width           =   2535
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
      Height          =   495
      Left            =   5520
      TabIndex        =   14
      Top             =   4560
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
      Height          =   495
      Left            =   5520
      MousePointer    =   4  'Icon
      TabIndex        =   13
      Top             =   3240
      Width           =   2535
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
      Left            =   5520
      TabIndex        =   12
      Text            =   "Choose Room No"
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
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
      Left            =   21120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
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
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Please Note:- Click ""Book New"" Button For New Booking Room"
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
      Left            =   13440
      TabIndex        =   32
      Top             =   8760
      Width           =   6615
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000002&
      Caption         =   "See_Per_Day_Rate"
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
      Left            =   15360
      TabIndex        =   27
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "800"
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
      Left            =   12120
      TabIndex        =   26
      Top             =   8880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Per Guest Rate"
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
      Left            =   9960
      TabIndex        =   25
      Top             =   8880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000002&
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
      Left            =   15480
      TabIndex        =   22
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   15120
      X2              =   15120
      Y1              =   1800
      Y2              =   8520
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   8520
      X2              =   8520
      Y1              =   1800
      Y2              =   8520
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
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
      Left            =   9240
      TabIndex        =   11
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
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
      Left            =   9240
      TabIndex        =   10
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
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
      Left            =   9240
      TabIndex        =   9
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
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
      Left            =   9240
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
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
      Left            =   9240
      TabIndex        =   7
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
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
      Left            =   3120
      TabIndex        =   6
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
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
      Index           =   3
      Left            =   3240
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
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
      Index           =   2
      Left            =   3120
      TabIndex        =   4
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
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
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
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
      Left            =   3240
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   6735
      Left            =   2880
      Picture         =   "Form7.frx":9B6F
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   17175
   End
   Begin VB.Image Image1 
      Height          =   12375
      Left            =   0
      Picture         =   "Form7.frx":A277
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Combo1_Click()
If Combo1.Text = "1" Then
MsgBox "This Is 1st Compart Of Hotel. Fully Loaded With Facalities. Comfortable Bed Room With Sponze Sheet, Lighting, Air Conditioner, Delicious Foods, Room Controler In Hand & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "2" Then
MsgBox "This Is 2nd Compart Of Hotel. Comfortable Bed Room With Sponze Sheet, Air Conditioner, Fully Loaded With Facalities. Lighting, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "3" Then
MsgBox "This Is 3rd Compart Of Hotel. Air Conditioner, Lighting, Room Controler In Hand, Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "4" Then
MsgBox "This Is 4th Compart Of Hotel. Fully Loaded With Facalities, Air Conditioner, Comfortable Bed Room With Sponze Sheet, Lighting, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "5" Then
MsgBox "This Is 5th Compart Of Hotel. Lighting, Room Controler In Hand, Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Air Conditioner, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "6" Then
MsgBox "This Is 6th Compart Of Hotel. Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Air Conditioner, Lighting, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "7" Then
MsgBox "This Is 7th Compart Of Hotel. Air Conditioner, Lighting, Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "8" Then
MsgBox "This Is 8th Compart Of Hotel. Fully Loaded With Facalities, Air Conditioner, Comfortable Bed Room With Sponze Sheet, Lighting, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "9" Then
MsgBox "This Is 9th Compart Of Hotel. Lighting, Room Controler In Hand, Delicious Foods, Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Air Conditioner & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "10" Then
MsgBox "This Is 10th Compart Of Hotel. Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Air Conditioner, Lighting, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "11" Then
MsgBox "This Is 11th Compart Of Hotel. Air Conditioner, Lighting, Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "12" Then
MsgBox "This Is 12th Compart Of Hotel. Fully Loaded With Facalities, Comfortable Bed Room With Sponze Sheet, Air Conditioner, Lighting, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "13" Then
MsgBox "This Is 13th Compart Of Hotel. Lighting, Room Controler In Hand, Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Air Conditioner, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "14" Then
MsgBox "This Is 14th Compart Of Hotel. Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Air Conditioner, Lighting, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "15" Then
MsgBox "This Is 15th Compart Of Hotel. Fully Loaded With Facalities, Air Conditioner, Comfortable Bed Room With Sponze Sheet, Lighting, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "16" Then
MsgBox "This Is 16th Compart Of Hotel. Air Conditioner, Lighting, Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "17" Then
MsgBox "This Is 17th Compart Of Hotel. Lighting, Room Controler In Hand, Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Air Conditioner, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "18" Then
MsgBox "This Is 18th Compart Of Hotel. Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Air Conditioner, Lighting, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "19" Then
MsgBox "This Is 19th Compart Of Hotel. Fully Loaded With Facalities, Air Conditioner, Comfortable Bed Room With Sponze Sheet, Lighting, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "20" Then
MsgBox "This Is 20th Compart Of Hotel. Air Conditioner, Lighting, Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "21" Then
MsgBox "This Is 21th Compart Of Hotel. Air Conditioner, Lighting, Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "22" Then
MsgBox "This Is 22th Compart Of Hotel. Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Air Conditioner, Lighting, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "23" Then
MsgBox "This Is 23th Compart Of Hotel. Fully Loaded With Facalities, Comfortable Bed Room With Sponze Sheet, Air Conditioner, Lighting, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "24" Then
MsgBox "This Is 24th Compart Of Hotel. Air Conditioner, Lighting, Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "25" Then
MsgBox "This Is 25th Compart Of Hotel. Room Controler In Hand, Delicious Foods, Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Air Conditioner, Lighting & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "26" Then
MsgBox "This Is 26th Compart Of Hotel. Fully Loaded With Facalities, Comfortable Bed Room With Sponze Sheet, Air Conditioner, Lighting, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "27" Then
MsgBox "This Is 27th Compart Of Hotel. Air Conditioner, Lighting, Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "28" Then
MsgBox "This Is 28th Compart Of Hotel. Lighting, Room Controler In Hand, Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Air Conditioner, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "29" Then
MsgBox "This Is 29th Compart Of Hotel. Comfortable Bed Room With Sponze Sheet, Fully Loaded With Facalities, Air Conditioner, Lighting, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If
If Combo1.Text = "30" Then
MsgBox "This Is 30th Compart Of Hotel. Fully Loaded With Facalities, Comfortable Bed Room With Sponze Sheet, Air Conditioner, Lighting, Room Controler In Hand, Delicious Foods, & Many More...!", vbInformation, "Available Facalities"
End If

End Sub
Private Sub Command1_Click()
Form6.Show
End Sub
Private Sub Command2_Click()
Form9.Show
End Sub
Private Sub Command3_Click()
Text5.Text = Val(Label4.Caption) * Val(Combo3.Text)
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
MsgBox "Room book Successfully.....!", vbInformation, "Room Booked"
rs.Update
Form6.Show
End Sub
Sub Clear()
Combo1.Text = "Select_Room"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo2.Text = "Gov_Valid_ID"
Text4.Text = ""
Combo3.Text = "How_Many_Guest"
Combo4.Text = "How_Many_Day"
Combo5.Text = "Comfort"
Text5.Text = ""
Text6.Text = ""
End Sub

Private Sub Command5_Click()
rs.AddNew
Clear
End Sub

Private Sub Command6_Click()
Text6.Text = Val(Text5.Text) * Val(Combo4.Text)
End Sub

Private Sub Form_Load()
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
End Sub
