VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Employee_form 
   BackColor       =   &H001F0E03&
   Caption         =   "Form1"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   18000
   WindowState     =   2  'Maximized
   Begin VB.ComboBox empcombo 
      Appearance      =   0  'Flat
      BackColor       =   &H00391E0F&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005757F4&
      Height          =   480
      Left            =   3960
      Sorted          =   -1  'True
      TabIndex        =   36
      Top             =   1560
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11040
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton Female 
      BackColor       =   &H001F0E03&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00539BEC&
      Height          =   450
      Left            =   3960
      TabIndex        =   32
      Top             =   8160
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0056300E&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005757F4&
      Height          =   495
      Index           =   6
      Left            =   3120
      MaxLength       =   6
      TabIndex        =   24
      Top             =   8160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox EmpDOJ 
      Height          =   495
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   2115
      TabIndex        =   23
      Top             =   7440
      Width           =   2175
   End
   Begin VB.PictureBox empDOB 
      Height          =   495
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   2115
      TabIndex        =   22
      Top             =   6000
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00391E0F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005757F4&
      Height          =   495
      Index           =   5
      Left            =   3960
      TabIndex        =   21
      Top             =   8880
      Width           =   3855
   End
   Begin VB.OptionButton Male 
      BackColor       =   &H001F0E03&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00539BEC&
      Height          =   450
      Left            =   5880
      TabIndex        =   19
      Top             =   8160
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00391E0F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005757F4&
      Height          =   495
      Index           =   4
      Left            =   3960
      MaxLength       =   6
      TabIndex        =   18
      Top             =   5040
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00391E0F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005757F4&
      Height          =   495
      Index           =   3
      Left            =   3960
      TabIndex        =   17
      Top             =   4320
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00391E0F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005757F4&
      Height          =   495
      Index           =   2
      Left            =   3960
      TabIndex        =   16
      Top             =   3600
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00391E0F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005757F4&
      Height          =   495
      Index           =   1
      Left            =   3960
      TabIndex        =   15
      Top             =   2280
      Width           =   3855
   End
   Begin VB.ComboBox departmentCombo 
      BackColor       =   &H00391E0F&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005757F4&
      Height          =   480
      ItemData        =   "empform.frx":0000
      Left            =   3960
      List            =   "empform.frx":000A
      TabIndex        =   20
      Top             =   6720
      Width           =   3855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FDD1AE&
      X1              =   240
      X2              =   6015
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00442C0D&
      X1              =   240
      X2              =   6015
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Details"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0052A7F3&
      Height          =   615
      Left            =   240
      TabIndex        =   38
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "In this form Admin can add Employee Details"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   600
      Width           =   5415
   End
   Begin VB.Label pincodeerror 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D48FF&
      Height          =   255
      Left            =   3960
      TabIndex        =   35
      Top             =   5520
      Width           =   3855
   End
   Begin VB.Label Newbtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   615
      Left            =   10680
      TabIndex        =   34
      ToolTipText     =   "Add New Datas of Employee"
      Top             =   9600
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   6
      Left            =   7920
      TabIndex        =   33
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   5
      Left            =   7920
      TabIndex        =   31
      Top             =   8880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   4
      Left            =   7920
      TabIndex        =   30
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   3
      Left            =   7920
      TabIndex        =   29
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   2
      Left            =   7920
      TabIndex        =   28
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   1
      Left            =   7920
      TabIndex        =   27
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label importimgbtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Import"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   8280
      TabIndex        =   26
      ToolTipText     =   "Add New Datas of Employee"
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label clearbtn 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear All"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0056300E&
      Height          =   495
      Left            =   6240
      TabIndex        =   25
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   6000
      Top             =   9480
      Width           =   1815
   End
   Begin VB.Label update_btn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   615
      Left            =   16440
      TabIndex        =   14
      ToolTipText     =   "Update a Record"
      Top             =   9600
      Width           =   1695
   End
   Begin VB.Label delete_btn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   615
      Left            =   14520
      TabIndex        =   13
      ToolTipText     =   "Delete a Record"
      Top             =   9600
      Width           =   1695
   End
   Begin VB.Label exit_btn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   615
      Left            =   18360
      TabIndex        =   12
      ToolTipText     =   "Exit Window"
      Top             =   9600
      Width           =   1695
   End
   Begin VB.Label save_btn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   615
      Left            =   12600
      TabIndex        =   11
      ToolTipText     =   "Add New Datas of Employee"
      Top             =   9600
      Width           =   1695
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00978313&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   18360
      Top             =   9480
      Width           =   1695
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00978313&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   12600
      Top             =   9480
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   3135
      Left            =   8280
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Image empphoto 
      Height          =   2895
      Left            =   8400
      Picture         =   "empform.frx":001E
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00828282&
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00539BEC&
      Height          =   495
      Left            =   960
      TabIndex        =   9
      Top             =   8880
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00539BEC&
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   8160
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Join"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00539BEC&
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00539BEC&
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00539BEC&
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   6720
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pincode"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00539BEC&
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00539BEC&
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "House Name"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00539BEC&
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00539BEC&
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00539BEC&
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H002C1612&
      FillStyle       =   0  'Solid
      Height          =   2655
      Left            =   480
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00978313&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14520
      Top             =   9480
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00978313&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   16440
      Top             =   9480
      Width           =   1695
   End
   Begin VB.Shape Shape8 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00978313&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   8280
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00978313&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   10680
      Top             =   9480
      Width           =   1695
   End
End
Attribute VB_Name = "Employee_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim empincre
Dim status As Integer
Dim imgpath As String

Private Sub clearbtn_Click()
Dim warn As Integer
warn = MsgBox("Clear All Fields", vbOKCancel, "Clear Data")
If warn = 1 Then
For i = 1 To 6
Text1(i) = ""
Next i
Female.Value = False
Male.Value = False
departmentCombo.Text = ""
'empDOB.Value = 1 / 1 / 1111
'EmpDOJ.Value = 1 / 1 / 1111

End If
End Sub

Private Sub delete_btn_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from Employee_Table where Emp_id='" & empcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
rs.Delete
rs.Update
rs.Close
empcombo.Clear
Call empcomboload
End Sub




Private Sub empcombo_click()
If rs.State = 1 Then rs.Close
rs.Open "select * from Employee_Table where Emp_id='" & empcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
If rs.EOF <> True Then
 empcombo.Text = rs.Fields(0)
 Text1(1).Text = rs.Fields(1)
 empphoto.Picture = LoadPicture(rs.Fields(10))
End If
End Sub

Private Sub Female_Click()
Call gendertest
End Sub

Private Sub empcomboload()
If rs.State = 1 Then rs.Close
rs.Open "select * from Employee_Table", con, adOpenKeyset, adLockOptimistic
rs.MoveFirst
While rs.EOF = False
empcombo.AddItem (rs.Fields(0))
rs.MoveNext
Wend
rs.Close
End Sub

Private Sub Form_Load()
connect
empincre = 1
emp_auto
status = 0
Call empcomboload
End Sub

Private Sub Label14_Click()
Employee_form.Hide
End Sub

Private Sub importimgbtn_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "JPG|*jpg"
imgpath = CommonDialog1.FileName
empphoto.Picture = LoadPicture(imgpath)
Print imgpath
End Sub

Private Sub Male_Click()
Call gendertest
End Sub

Private Sub mvnxtbtn_Click()
rs.MoveNext
End Sub

Private Sub mvprebtn_Click()
rs.MovePrevious
End Sub


Private Sub newbtn_Click()
clearbtn_Click
empcombo.Clear
emp_auto
Call empcomboload
End Sub

Private Sub save_btn_Click()
Call checkfields
If status = 0 Then
If rs.State = 1 Then
    rs.Close
End If

rs.Open "select * from Employee_Table", con, adOpenKeyset, adLockOptimistic
Call gendertest
With rs
.AddNew
.Fields(0) = empcombo.Text
.Fields(1) = Text1(1)
.Fields(2) = Text1(2)
.Fields(3) = Text1(3)
.Fields(4) = Text1(4)
'.Fields(5) = empDOB.Value
'.Fields(6) = departmentCombo.
'.Fields(7) = EmpDOJ.Value
.Fields(8) = Text1(6).Text
.Fields(9) = Val(Text1(5))
.Fields(10) = imgpath
.Update
End With
MsgBox ("Employee Details Stored")
rs.Close
End If 'Status if
empcombo.Clear
Call empcomboload
End Sub

Private Sub emp_auto()
Dim empidval

If rs.State = 1 Then
rs.Close
End If

rs.Open "select * from Employee_Table", con, adOpenKeyset, adLockOptimistic
If rs.EOF <> True Then rs.MoveLast


empidval = rs.Fields(0)
If empincre < 10 Then
empincre = Mid(empidval, 4, 1)
Else
empincre = Mid(empidval, 4, 2)
End If
empcombo.Text = "Emp" & empincre + 1
rs.Close
End Sub

Private Sub gendertest()
If Female.Value = True Then
    Text1(6).Text = Female.Caption
Else
    Text1(6).Text = Male.Caption
End If
End Sub

Private Sub checkfields()
status = 0
For i = 1 To 6
 If Text1(i).Text = "" Then
  Label11(i).Visible = True
  status = 1
Else
  Label11(i).Visible = False
 End If
Next i
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
 If IsNumeric(Text1(4).Text) = False Then
  pincodeerror.Caption = "Only Numeric values"
Else
  pincodeerror.Caption = ""
 End If
End Sub

Private Sub update_btn_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from Employee_Table where Emp_id = '" & empcombo.Text & "'", con, adOpenKeyset, adLockOptimistic

Call gendertest
With rs
.Fields(0) = empcombo.Text
.Fields(1) = Text1(1)
.Fields(2) = Text1(2)
.Fields(3) = Text1(3)
.Fields(4) = Text1(4)
'.Fields(5) = empDOB.Value
'.Fields(6) = departmentCombo.
'.Fields(7) = EmpDOJ.Value
.Fields(8) = Text1(6).Text
.Fields(9) = Val(Text1(5))
.Fields(10) = imgpath
.Update
End With
MsgBox ("Employee Details Stored")
rs.Close
End Sub
