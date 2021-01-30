VERSION 5.00
Begin VB.Form Game 
   BackColor       =   &H001D181B&
   Caption         =   "Fun"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15225
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.Timer gametimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10680
      Top             =   1800
   End
   Begin VB.Timer thinktimer 
      Interval        =   100
      Left            =   15120
      Top             =   4680
   End
   Begin VB.Timer timeout 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   18360
      Top             =   8640
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H001F1F1F&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   10320
      Width           =   20655
      Begin VB.Timer loadtimer 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   8160
         Top             =   120
      End
      Begin VB.Shape loadingbar 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H005B5BF4&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   0
         Top             =   0
         Width           =   15
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copyrighted at Track && Trail"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0052A7F3&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   20175
      End
   End
   Begin VB.Label timerdot 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   8520
      TabIndex        =   13
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label hourlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1095
      Left            =   7320
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label seclabel 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   10080
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label minlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1095
      Left            =   8880
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Image bigcycleimage 
      Height          =   3975
      Left            =   11760
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   6735
   End
   Begin VB.Label keepthinklabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Keep  Thinking"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   13440
      TabIndex        =   9
      Top             =   3720
      Width           =   3735
   End
   Begin VB.Shape waitingshape 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   13560
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape waitingshape 
      FillColor       =   &H00F3F3F3&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   13680
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape waitingshape 
      FillColor       =   &H00E5E5E5&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   13800
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape waitingshape 
      FillColor       =   &H00D4D4D4&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   13920
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape waitingshape 
      FillColor       =   &H00C1C1C1&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   14040
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape waitingshape 
      FillColor       =   &H00B0B0B0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   14160
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape waitingshape 
      FillColor       =   &H00A5A5A5&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   14280
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape waitingshape 
      FillColor       =   &H00919191&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   14400
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape waitingshape 
      FillColor       =   &H00818181&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   8
      Left            =   14520
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape waitingshape 
      FillColor       =   &H00737373&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   9
      Left            =   14640
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape waitingshape 
      FillColor       =   &H00676767&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   10
      Left            =   14760
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape waitingshape 
      FillColor       =   &H005C5C5C&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   11
      Left            =   14880
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D7D7D7&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   16680
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C6C6C6&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   16560
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00B5B5B5&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   16440
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00ABABAB&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   16320
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H009C9C9C&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   16200
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00939393&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   16080
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00888888&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   15960
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H007C7C7C&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   8
      Left            =   15840
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00747474&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   9
      Left            =   15720
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00676767&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   10
      Left            =   15600
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H005C5C5C&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   11
      Left            =   15480
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label selectlabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0052A7F3&
      Height          =   375
      Left            =   18600
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Shape timershape 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H005B5BF4&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   15240
      Top             =   8640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Image cy3img 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   18600
      Picture         =   "Form2.frx":249C8
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label checkbtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   15240
      TabIndex        =   1
      Top             =   6840
      Width           =   3015
   End
   Begin VB.Image cy2img 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   18600
      Picture         =   "Form2.frx":4AC65
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Image cy1img 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   18600
      Picture         =   "Form2.frx":70172
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Shape Shape8 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H002A2A2A&
      FillStyle       =   0  'Solid
      Height          =   4815
      Left            =   18600
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Helpbtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   15240
      TabIndex        =   7
      Top             =   8040
      Width           =   3015
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   15240
      Top             =   7920
      Width           =   3015
   End
   Begin VB.Label newimgbtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   11880
      TabIndex        =   6
      Top             =   6840
      Width           =   3015
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   15240
      Top             =   6720
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "In this from Admin can play a mini game for refreshness"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   4575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00422C0D&
      X1              =   120
      X2              =   5520
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H006B4614&
      X1              =   120
      X2              =   5520
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fun Section"
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
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label quitbtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I Quit"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   12000
      TabIndex        =   0
      Top             =   8040
      Width           =   3015
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   12000
      Top             =   7920
      Width           =   3015
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   12000
      Top             =   6720
      Width           =   3015
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H001F1F1F&
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   11760
      Top             =   6600
      Width           =   6735
   End
   Begin VB.Shape secselectbox 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   975
      Left            =   1560
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape firstselectbox 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   975
      Left            =   8760
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   42
      Left            =   7560
      Picture         =   "Form2.frx":756CD
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   43
      Left            =   8760
      Picture         =   "Form2.frx":7737C
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   44
      Left            =   9960
      Picture         =   "Form2.frx":79077
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   45
      Left            =   360
      Picture         =   "Form2.frx":7ABE1
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   46
      Left            =   1560
      Picture         =   "Form2.frx":7C953
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   47
      Left            =   2760
      Picture         =   "Form2.frx":7DEFB
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   48
      Left            =   3960
      Picture         =   "Form2.frx":7FAFE
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   49
      Left            =   5160
      Picture         =   "Form2.frx":81739
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   50
      Left            =   6360
      Picture         =   "Form2.frx":8347C
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   51
      Left            =   7560
      Picture         =   "Form2.frx":85076
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   52
      Left            =   8760
      Picture         =   "Form2.frx":86CE0
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   53
      Left            =   9960
      Picture         =   "Form2.frx":883B9
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   41
      Left            =   6360
      Picture         =   "Form2.frx":89F5F
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   40
      Left            =   5160
      Picture         =   "Form2.frx":8BD57
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   39
      Left            =   3960
      Picture         =   "Form2.frx":8D9BB
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   38
      Left            =   2760
      Picture         =   "Form2.frx":8F710
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   37
      Left            =   1560
      Picture         =   "Form2.frx":9100F
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   36
      Left            =   360
      Picture         =   "Form2.frx":91ED3
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   35
      Left            =   9960
      Picture         =   "Form2.frx":92B44
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   34
      Left            =   8760
      Picture         =   "Form2.frx":944E0
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   33
      Left            =   7560
      Picture         =   "Form2.frx":95EC1
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   32
      Left            =   6360
      Picture         =   "Form2.frx":97A64
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   31
      Left            =   5160
      Picture         =   "Form2.frx":997E0
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   30
      Left            =   3960
      Picture         =   "Form2.frx":9B062
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   29
      Left            =   2760
      Picture         =   "Form2.frx":9C5BD
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   28
      Left            =   1560
      Picture         =   "Form2.frx":9E133
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   27
      Left            =   360
      Picture         =   "Form2.frx":9FE54
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   26
      Left            =   9960
      Picture         =   "Form2.frx":A1951
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   25
      Left            =   8760
      Picture         =   "Form2.frx":A35AE
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   24
      Left            =   7560
      Picture         =   "Form2.frx":A5124
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   23
      Left            =   6360
      Picture         =   "Form2.frx":A6ADC
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   22
      Left            =   5160
      Picture         =   "Form2.frx":A8552
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   21
      Left            =   3960
      Picture         =   "Form2.frx":A9833
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   20
      Left            =   2760
      Picture         =   "Form2.frx":AA3F6
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   19
      Left            =   1560
      Picture         =   "Form2.frx":AACEC
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   18
      Left            =   360
      Picture         =   "Form2.frx":AC8B3
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   17
      Left            =   9960
      Picture         =   "Form2.frx":AE48B
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   16
      Left            =   8760
      Picture         =   "Form2.frx":AFA6C
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   15
      Left            =   7560
      Picture         =   "Form2.frx":B1289
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   14
      Left            =   6360
      Picture         =   "Form2.frx":B2E05
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   13
      Left            =   5160
      Picture         =   "Form2.frx":B4A8A
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   12
      Left            =   3960
      Picture         =   "Form2.frx":B6813
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   11
      Left            =   2760
      Picture         =   "Form2.frx":B8404
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   10
      Left            =   1560
      Picture         =   "Form2.frx":BA0A4
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   9
      Left            =   360
      Picture         =   "Form2.frx":BBCE0
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   8
      Left            =   9960
      Picture         =   "Form2.frx":BCDB8
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   7
      Left            =   8760
      Picture         =   "Form2.frx":BDA34
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   6
      Left            =   7560
      Picture         =   "Form2.frx":BE7EE
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   5
      Left            =   6360
      Picture         =   "Form2.frx":BFC7B
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   4
      Left            =   5160
      Picture         =   "Form2.frx":C1395
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   3
      Left            =   3960
      Picture         =   "Form2.frx":C2235
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   2
      Left            =   2760
      Picture         =   "Form2.frx":C3649
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   1
      Left            =   1560
      Picture         =   "Form2.frx":C47DF
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   0
      Left            =   360
      Picture         =   "Form2.frx":C5283
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H002A2A2A&
      FillStyle       =   0  'Solid
      Height          =   6615
      Left            =   240
      Top             =   2400
      Width           =   10935
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H002A2A2A&
      FillStyle       =   0  'Solid
      Height          =   6615
      Left            =   11640
      Top             =   2400
      Width           =   6975
   End
End
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selectedimg As Integer
Dim sec As Integer
Dim imagesize As Integer
Dim imgar(54) As Integer
Dim imgchance As Integer
Dim imgpos(54, 2) As Long
Dim first As Integer
Dim poor As Integer
Dim thinkvari As Integer
Dim flag As Integer
Dim onlyonce1 As Integer
Dim gh As Integer
Dim gm As Integer
Dim gs As Integer

Private Sub getimgposition()
For i = 0 To 53
 imgpos(i, 0) = Image1(i).Index
 imgpos(i, 1) = Image1(i).Left
 imgpos(i, 2) = Image1(i).Top
Next i
End Sub

Private Sub putimgposition()
For i = 0 To 53
 Image1(i).Left = imgpos(i, 0)
 Image1(i).Top = imgpos(i, 1)
Next i
End Sub




Private Sub cy1img_Click()
bigcycleimage.Picture = LoadPicture("F:\Project46\MiniProject\Images\p5pb13058137.jpg")
imgchance = 1
Call arrangeimage
Call Helpbtn_Click
loadtimer.Enabled = True
End Sub

Private Sub cy2img_Click()
bigcycleimage.Picture = LoadPicture("F:\Project46\MiniProject\Images\Untitled-1-Recovered-Recovered.jpg")
imgchance = 2
Call arrangeimage

Call Helpbtn_Click
loadtimer.Enabled = True
End Sub

Private Sub cy3img_Click()
bigcycleimage.Picture = LoadPicture("F:\Project46\MiniProject\Images\ath.jpg")
imgchance = 3
Call arrangeimage
Call Helpbtn_Click
loadtimer.Enabled = True
End Sub

Private Sub arrangeimage()
If imgchance = 1 Then
For i = 1 To 54
 Image1(i - 1).Picture = LoadPicture("F:\Project46\MiniProject\Images\Game\One\images\85b3314cd32bc6f3bd59f21e9dbec0d0_" & i & ".gif")
Next i
ElseIf imgchance = 2 Then
For i = 1 To 54
Image1(i - 1).Picture = LoadPicture("F:\Project46\MiniProject\Images\Game\Two\images\Untitled-1-Recovered_" & i & ".jpg")
Next i
ElseIf imgchance = 3 Then
For i = 1 To 54
 Image1(i - 1).Picture = LoadPicture("F:\Project46\MiniProject\Images\Game\Three\images\ath_" & i & ".jpg")
Next i
End If
If onlyonce1 = 0 Then
Call getimgposition
onlyonce1 = onlyonce1 + 1
End If
End Sub



Private Sub gametimer_Timer()
 If gs < 10 Then
 seclabel.Caption = "0" & gs
 Else
 seclabel.Caption = gs
 End If
 gs = gs + 1
   If gm < 10 Then
  minlabel.Caption = "0" & gm
  Else
  minlabel.Caption = gm
  End If
 If gs = 60 Then
  gs = 0
  gm = gm + 1
  If gh < 10 Then
  hourlabel.Caption = "0" & gh
  Else
  hourlabel.Caption = gh
  End If
 If gm = 60 Then
  gm = 0
  gh = gh + 1
 End If
 End If
 If gs Mod 2 = 0 Then
  timerdot.Visible = True
 Else
  timerdot.Visible = False
 End If
End Sub

Private Sub loadtimer_Timer()
If loadingbar.Width <= 20535 Then
 loadingbar.Width = loadingbar.Width + 2000
Else
 loadingbar.Width = 0
 loadtimer.Enabled = False
End If
End Sub

Private Sub newimgbtn_Click()
If imgchance = 1 Or imgchance = 2 Or imgchance = 3 Then
Dim rndindex As Integer
Dim imgchecker As Integer
Dim k As Integer

Randomize
For i = 0 To 54
 imgar(i) = 0
Next i

k = 0

For i = 0 To 1500
If k <= 53 Then
rndindex = CInt(Int((53 * Rnd()) + 1))
imgchecker = isindextaken(rndindex)
If imgchecker <> 1 Then
imgar(k) = rndindex
k = k + 1
Else
imgchecker = 0
End If
Else
Exit For
End If
Next i
Call allocaterandomimages
gametimer.Enabled = True
End If
End Sub

Private Sub allocaterandomimages()
If imgchance = 1 Then
For i = 0 To 53
 Image1(imgar(i)).Left = imgpos(i, 1)
 Image1(imgar(i)).Top = imgpos(i, 2)
Next i
ElseIf imgchance = 2 Then
For i = 0 To 53
 Image1(imgar(i)).Left = imgpos(i, 1)
 Image1(imgar(i)).Top = imgpos(i, 2)
Next i
ElseIf imgchance = 3 Then
For i = 0 To 53
 Image1(imgar(i)).Left = imgpos(i, 1)
 Image1(imgar(i)).Top = imgpos(i, 2)
Next i
End If
End Sub

Function isindextaken(rndindexs) As Integer
 For i = 0 To 54
  If imgar(i) = rndindexs Then
  isindextaken = 1
  Exit For
  End If
 Next i
End Function

Private Sub swapimage(k, j)
Dim templeft As Long
Dim temptop As Long

templeft = Image1(k).Left
Image1(k).Left = Image1(j).Left
Image1(j).Left = templeft

temptop = Image1(k).Top
Image1(k).Top = Image1(j).Top
Image1(j).Top = temptop

End Sub

Private Sub checkbtn_Click()
If imgchance = 1 Or imgchance = 2 Or imgchance = 3 Then
If imagesize = 0 Then
 Call imageshort
 imagesize = 1
Else
 Call imagelarge
 imagesize = 0
End If
If checkbtn.Caption = "Check" Then
 checkbtn.Caption = "Uncheck"
Else
 checkbtn.Caption = "Check"
End If
End If
End Sub


Private Sub imageshort()
Dim k As Integer
k = 1
 For i = 0 To 53
  If i <> 0 And i <> 9 And i <> 18 And i <> 27 And i <> 36 And i <> 45 Then
   Image1(i).Left = Image1(i).Left - (k * 110)
   k = k + 1
  Else
   k = 1
  End If
 Next i
 
For i = 0 To 53
 If i >= 9 And i <= 17 Then
  Image1(i).Top = Image1(i).Top - 110
 ElseIf i >= 18 And i <= 26 Then
  Image1(i).Top = Image1(i).Top - 220
 ElseIf i >= 27 And i <= 35 Then
  Image1(i).Top = Image1(i).Top - 330
 ElseIf i >= 34 And i <= 44 Then
  Image1(i).Top = Image1(i).Top - 440
 ElseIf i >= 45 And i <= 53 Then
  Image1(i).Top = Image1(i).Top - 550
 End If
Next i
End Sub

Private Sub imagelarge()
Dim k As Integer
k = 1
 For i = 0 To 53
  If i <> 0 And i <> 9 And i <> 18 And i <> 27 And i <> 36 And i <> 45 Then
   Image1(i).Left = Image1(i).Left + (k * 110)
   k = k + 1
  Else
   k = 1
  End If
 Next i
 
For i = 0 To 53
 If i >= 9 And i <= 17 Then
  Image1(i).Top = Image1(i).Top + 110
 ElseIf i >= 18 And i <= 26 Then
  Image1(i).Top = Image1(i).Top + 220
 ElseIf i >= 27 And i <= 35 Then
  Image1(i).Top = Image1(i).Top + 330
 ElseIf i >= 34 And i <= 44 Then
  Image1(i).Top = Image1(i).Top + 440
 ElseIf i >= 45 And i <= 53 Then
  Image1(i).Top = Image1(i).Top + 550
 End If
Next i
End Sub

Private Sub Form_Load()
thinkvari = 0
start = 0
first = 0
imagesize = 0
loaddefaultimg
timerstart = 0
onlyonce1 = 0
End Sub

Private Sub loaddefaultimg()
For i = 1 To 54
 Image1(i - 1).Picture = LoadPicture("F:\Project46\MiniProject\Images\Game\Default\images\Untitled-1_" & i & ".jpg")
Next i
End Sub

Private Sub Image1_Click(Index As Integer)
If imgchance = 1 Or imgchance = 2 Or imgchance = 3 Then
If selectedimg = 0 Then
 selectedimg = selectedimg + 1
 first = Index
 firstselectbox.Visible = True
 secselectbox.Visible = False
 firstselectbox.Left = Image1(Index).Left
 firstselectbox.Top = Image1(Index).Top
Else
 sec = Index
 secselectbox.Visible = True
 secselectbox.Left = Image1(Index).Left
 secselectbox.Top = Image1(Index).Top
 Call swapimage(first, sec)
 selectedimg = 0
End If
Call checkans
End If
End Sub

Private Sub checkans()
flag = 0
For i = 0 To 54
 If Image1(imgar(i)).Left = imgpos(imgar(i), 1) And Image1(imgar(i)).Top = imgpos(imgar(i), 2) Then
 Else
  flag = 1
  Exit For
 End If
Next i

If flag = 0 Then
 MsgBox "Congratzzz...You Beat Mee (Time: " & hourlabel.Caption & ":" & minlabel.Caption & ":" & seclabel.Caption & ")"""
 gametimer.Enabled = False
 minlabel.Caption = "00"
 seclabel.Caption = "00"
 hourlabel.Caption = "00"
End If
End Sub

Private Sub quitbtn_Click()
If imgchance = 1 Or imgchance = 2 Or imgchance = 3 Then
For i = 0 To 53
 Image1(i).Left = imgpos(i, 1)
 Image1(i).Top = imgpos(i, 2)
Next i
End If
 gametimer.Enabled = False
 MsgBox "Hehehe You Can't Beat Me...(Time: " & hourlabel.Caption & ":" & minlabel.Caption & ":" & seclabel.Caption & ")"
 minlabel.Caption = "00"
 seclabel.Caption = "00"
 hourlabel.Caption = "00"
End Sub

Private Sub Helpbtn_Click()
If imgchance = 1 Or imgchance = 2 Or imgchance = 3 Then
timeout.Enabled = True
timershape.Visible = True
timershape.Width = 3015
Helpbtn.Enabled = False
bigcycleimage.Visible = True
End If
End Sub

Private Sub thinktimer_Timer()

If thinkvari <= 11 Then
 waitingshape(thinkvari).Visible = True
 Shape9(thinkvari).Visible = True
 thinkvari = thinkvari + 1
Else
 For i = 0 To 11
  waitingshape(i).Visible = False
  Shape9(i).Visible = False
  thinkvari = 0
 Next i
End If
End Sub

Private Sub timeout_Timer()
If timershape.Width >= 20 Then
 timershape.Width = timershape.Width - 5
Else
 timeout.Enabled = False
 timershape.Visible = False
 Helpbtn.Enabled = True
 bigcycleimage.Visible = False
End If
End Sub

