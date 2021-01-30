VERSION 5.00
Begin VB.Form login_form 
   Caption         =   "Login"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19305
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   19305
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5520
      Top             =   2760
   End
   Begin VB.Timer timer1 
      Interval        =   3000
      Left            =   600
      Top             =   7920
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00372406&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EAC191&
      Height          =   570
      IMEMode         =   3  'DISABLE
      Left            =   15360
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00372406&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EAC191&
      Height          =   570
      Left            =   11160
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00654114&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   13080
      Top             =   8160
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00654114&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   10680
      Top             =   8160
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00654114&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   8280
      Top             =   8160
      Width           =   2055
   End
   Begin VB.Image Image7 
      Height          =   855
      Left            =   6600
      Picture         =   "Login.frx":0000
      Stretch         =   -1  'True
      Top             =   9360
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Image Image6 
      Height          =   615
      Left            =   4560
      Picture         =   "Login.frx":A4F8
      Stretch         =   -1  'True
      Top             =   9480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   2520
      Picture         =   "Login.frx":10527
      Stretch         =   -1  'True
      Top             =   9480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "*TOP BRANDS"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   9600
      Width           =   2055
   End
   Begin VB.Image Image4 
      Height          =   1815
      Left            =   13560
      Picture         =   "Login.frx":15EEE
      Stretch         =   -1  'True
      Top             =   8880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   1815
      Left            =   15840
      Picture         =   "Login.frx":171B2
      Stretch         =   -1  'True
      Top             =   8880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "*OUR SPONSORS"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   11280
      TabIndex        =   7
      Top             =   9600
      Width           =   2055
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00281A04&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   11280
      Top             =   8760
      Width           =   2055
   End
   Begin VB.Label passmxerror 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5BF4&
      Height          =   255
      Left            =   15360
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label news 
      BackStyle       =   0  'Transparent
      Caption         =   "Sorry Username and Password Doesn't Match"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5BF4&
      Height          =   375
      Left            =   20520
      TabIndex        =   5
      Top             =   2760
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00EBA438&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   5880
      Top             =   8160
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   4455
      Left            =   600
      Picture         =   "Login.frx":1C94D
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   19095
   End
   Begin VB.Label login_btn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   17640
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005EE199&
      Height          =   375
      Left            =   13560
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005EE199&
      Height          =   375
      Left            =   9360
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00281A04&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   480
      Top             =   8760
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   11175
      Left            =   0
      Picture         =   "Login.frx":3A1FB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20535
   End
End
Attribute VB_Name = "login_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim imgtimer As Integer
Dim imgon As String
Dim imgoff As String

Private Sub Command1_Click()
Timer2.Enabled = True
End Sub

Private Sub Form_Load()
Call connect
imgtimer = 1
imgon = &HEBA438
imgoff = &H654114
Shape5.FillColor = RGB(9, 19, 28)
Shape6.FillColor = RGB(9, 19, 28)
End Sub

Private Sub Label3_click()
If Shape5.Width <> 6615 Then
Shape5.Width = 6615
Image3.Visible = True
Image4.Visible = True
Else
Shape5.Width = 2055
Image3.Visible = False
Image4.Visible = False
End If
End Sub

Private Sub Label4_Click()
If Shape6.Width <> 10335 Then
Shape6.Width = 10335
Image5.Visible = True
Image6.Visible = True
Image7.Visible = True
Else
Shape6.Width = 2175
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False
End If
End Sub


Private Sub login_btn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
login_btn.ForeColor = &H288EF4
End Sub

Private Sub login_btn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
login_btn.ForeColor = &HFFFFFF
End Sub

Private Sub login_btn_Click()
If Text1.Text = "" Or Text2.Text = "" Then
news.Caption = "Please fill all fields"
Timer2.Enabled = True
Else
rs.Open "select * from login_table", con, adOpenKeyset, adLockOptimistic
If rs.Fields(0) = Text1.Text And rs.Fields(1) = Text2.Text Then
  MDI_Form.Show
  login_form.Hide
Else
  news.Caption = "Sorry Username and password doesn't match"
  Timer2.Enabled = True
End If
rs.Close
 End If
End Sub

Private Sub Text2_Change()
If Len(Text2.Text) = Text2.MaxLength Then
 passmxerror.Caption = "Max Length Reached"
Else
 passmxerror.Caption = ""
End If
End Sub

Private Sub timer1_Timer()
 Image2.Picture = LoadPicture("F:\Project46\MiniProject\Images\ban" & imgtimer & ".jpg")
 For i = 1 To 4
  If i = imgtimer Then
   Shape1(i).FillColor = &HEBA438
 ElseIf i <> imgtimer Then
   Shape1(i).FillColor = &H654114
 End If
 Next i
 If imgtimer = 4 Then
  imgtimer = 1
 Else
imgtimer = imgtimer + 1
 End If
End Sub

Private Sub Timer2_Timer()
If news.Left <= -4000 Then
 news.Left = 20520
 Timer2.Enabled = False
Else
news.Left = news.Left - 50
End If
End Sub

