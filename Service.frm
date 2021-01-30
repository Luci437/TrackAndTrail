VERSION 5.00
Begin VB.Form Sales 
   BackColor       =   &H000E0901&
   Caption         =   "Sales details"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19020
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   9630
   ScaleWidth      =   19020
   WindowState     =   2  'Maximized
   Begin VB.ListBox prdidlist 
      Appearance      =   0  'Flat
      BackColor       =   &H00080808&
      Columns         =   1
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A5BBFE&
      Height          =   4305
      ItemData        =   "Service.frx":0000
      Left            =   12240
      List            =   "Service.frx":0002
      TabIndex        =   49
      Top             =   3600
      Width           =   870
   End
   Begin VB.ListBox qtylist 
      Appearance      =   0  'Flat
      BackColor       =   &H00080808&
      Columns         =   1
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A5BBFE&
      Height          =   4305
      ItemData        =   "Service.frx":0004
      Left            =   17880
      List            =   "Service.frx":0006
      TabIndex        =   43
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox cstname 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   375
      Left            =   7320
      TabIndex        =   41
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Timer tyreanimation2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   9240
   End
   Begin VB.Timer tyreanimation 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   9240
   End
   Begin VB.ComboBox prdidcombo 
      Appearance      =   0  'Flat
      BackColor       =   &H00080808&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   420
      Left            =   13920
      TabIndex        =   4
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox prdqntbox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   330
      Left            =   16080
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton addtolistbtn 
      Caption         =   "ADD"
      Height          =   375
      Left            =   19200
      TabIndex        =   8
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox prdamtbox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   375
      Left            =   17400
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox prdnamebox 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   375
      Left            =   12480
      TabIndex        =   5
      Top             =   2760
      Width           =   3255
   End
   Begin VB.ComboBox mopcombo 
      Appearance      =   0  'Flat
      BackColor       =   &H00080808&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   420
      ItemData        =   "Service.frx":0008
      Left            =   7320
      List            =   "Service.frx":0015
      TabIndex        =   3
      Top             =   4080
      Width           =   2295
   End
   Begin VB.ComboBox cstridcombo 
      Appearance      =   0  'Flat
      BackColor       =   &H00080808&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   420
      Left            =   7320
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
   Begin VB.ComboBox salesidcombo 
      Appearance      =   0  'Flat
      BackColor       =   &H00080808&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   420
      ItemData        =   "Service.frx":0039
      Left            =   7320
      List            =   "Service.frx":003B
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox discountamt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   375
      Left            =   14760
      TabIndex        =   9
      Top             =   8760
      Width           =   1095
   End
   Begin VB.ListBox prdamtlist 
      Appearance      =   0  'Flat
      BackColor       =   &H00080808&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A5BBFE&
      Height          =   4305
      ItemData        =   "Service.frx":003D
      Left            =   18360
      List            =   "Service.frx":003F
      TabIndex        =   25
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ListBox prdlist 
      Appearance      =   0  'Flat
      BackColor       =   &H00080808&
      Columns         =   1
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A5BBFE&
      Height          =   4305
      ItemData        =   "Service.frx":0041
      Left            =   13200
      List            =   "Service.frx":0043
      TabIndex        =   24
      Top             =   3600
      Width           =   4575
   End
   Begin VB.TextBox totalamttopay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   375
      Left            =   16800
      TabIndex        =   10
      Top             =   8760
      Width           =   2775
   End
   Begin VB.Label savelabel 
      Alignment       =   2  'Center
      BackColor       =   &H00110609&
      BackStyle       =   0  'Transparent
      Caption         =   "Data Saved"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5BF4&
      Height          =   375
      Left            =   8160
      TabIndex        =   48
      Top             =   10250
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label notfillederror 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   9720
      TabIndex        =   47
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label notfillederror 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   9720
      TabIndex        =   46
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label notfillederror 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   9720
      TabIndex        =   45
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00110609&
      BackStyle       =   0  'Transparent
      Caption         =   "QTY"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF3E8&
      Height          =   375
      Left            =   17880
      TabIndex        =   44
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H001F130A&
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER NAME"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   4560
      TabIndex        =   42
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   41
      Left            =   19800
      Picture         =   "Service.frx":0045
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   40
      Left            =   19320
      Picture         =   "Service.frx":3A24
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   39
      Left            =   18840
      Picture         =   "Service.frx":7403
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   38
      Left            =   18360
      Picture         =   "Service.frx":ADE2
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   37
      Left            =   17880
      Picture         =   "Service.frx":E7C1
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   36
      Left            =   17400
      Picture         =   "Service.frx":121A0
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   35
      Left            =   16920
      Picture         =   "Service.frx":15B7F
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   34
      Left            =   16440
      Picture         =   "Service.frx":1955E
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   33
      Left            =   15960
      Picture         =   "Service.frx":1CF3D
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   32
      Left            =   15480
      Picture         =   "Service.frx":2091C
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   31
      Left            =   15000
      Picture         =   "Service.frx":242FB
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   30
      Left            =   14520
      Picture         =   "Service.frx":27CDA
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   29
      Left            =   14040
      Picture         =   "Service.frx":2B6B9
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   28
      Left            =   13560
      Picture         =   "Service.frx":2F098
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   27
      Left            =   13080
      Picture         =   "Service.frx":32A77
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   26
      Left            =   12600
      Picture         =   "Service.frx":36456
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   25
      Left            =   12120
      Picture         =   "Service.frx":39E35
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   24
      Left            =   11640
      Picture         =   "Service.frx":3D814
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   23
      Left            =   11160
      Picture         =   "Service.frx":411F3
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   22
      Left            =   10680
      Picture         =   "Service.frx":44BD2
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   21
      Left            =   10200
      Picture         =   "Service.frx":485B1
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   20
      Left            =   9720
      Picture         =   "Service.frx":4BF90
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   19
      Left            =   9240
      Picture         =   "Service.frx":4F96F
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   18
      Left            =   8760
      Picture         =   "Service.frx":5334E
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   17
      Left            =   8280
      Picture         =   "Service.frx":56D2D
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   16
      Left            =   7800
      Picture         =   "Service.frx":5A70C
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   15
      Left            =   7320
      Picture         =   "Service.frx":5E0EB
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   14
      Left            =   6840
      Picture         =   "Service.frx":61ACA
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   13
      Left            =   6360
      Picture         =   "Service.frx":654A9
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   12
      Left            =   5880
      Picture         =   "Service.frx":68E88
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   11
      Left            =   5400
      Picture         =   "Service.frx":6C867
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   10
      Left            =   4920
      Picture         =   "Service.frx":70246
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   9
      Left            =   4440
      Picture         =   "Service.frx":73C25
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   8
      Left            =   3960
      Picture         =   "Service.frx":77604
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   7
      Left            =   3480
      Picture         =   "Service.frx":7AFE3
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   6
      Left            =   3000
      Picture         =   "Service.frx":7E9C2
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   5
      Left            =   2520
      Picture         =   "Service.frx":823A1
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   4
      Left            =   2040
      Picture         =   "Service.frx":85D80
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   3
      Left            =   1560
      Picture         =   "Service.frx":8975F
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   2
      Left            =   1080
      Picture         =   "Service.frx":8D13E
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   1
      Left            =   600
      Picture         =   "Service.frx":90B1D
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "Service.frx":944FC
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   41
      Left            =   19800
      Picture         =   "Service.frx":97EDB
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   40
      Left            =   19320
      Picture         =   "Service.frx":9B8B4
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   39
      Left            =   18840
      Picture         =   "Service.frx":9F28D
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   38
      Left            =   18360
      Picture         =   "Service.frx":A2C66
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   37
      Left            =   17880
      Picture         =   "Service.frx":A663F
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   36
      Left            =   17400
      Picture         =   "Service.frx":AA018
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   35
      Left            =   16920
      Picture         =   "Service.frx":AD9F1
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   34
      Left            =   16440
      Picture         =   "Service.frx":B13CA
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   33
      Left            =   15960
      Picture         =   "Service.frx":B4DA3
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   32
      Left            =   15480
      Picture         =   "Service.frx":B877C
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   31
      Left            =   15000
      Picture         =   "Service.frx":BC155
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   30
      Left            =   14520
      Picture         =   "Service.frx":BFB2E
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   29
      Left            =   14040
      Picture         =   "Service.frx":C3507
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   28
      Left            =   13560
      Picture         =   "Service.frx":C6EE0
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   27
      Left            =   13080
      Picture         =   "Service.frx":CA8B9
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   26
      Left            =   12600
      Picture         =   "Service.frx":CE292
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   25
      Left            =   12120
      Picture         =   "Service.frx":D1C6B
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   24
      Left            =   11640
      Picture         =   "Service.frx":D5644
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   23
      Left            =   11160
      Picture         =   "Service.frx":D901D
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   22
      Left            =   10680
      Picture         =   "Service.frx":DC9F6
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   21
      Left            =   10200
      Picture         =   "Service.frx":E03CF
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   20
      Left            =   9720
      Picture         =   "Service.frx":E3DA8
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   19
      Left            =   9240
      Picture         =   "Service.frx":E7781
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   18
      Left            =   8760
      Picture         =   "Service.frx":EB15A
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   17
      Left            =   8280
      Picture         =   "Service.frx":EEB33
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   16
      Left            =   7800
      Picture         =   "Service.frx":F250C
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   15
      Left            =   7320
      Picture         =   "Service.frx":F5EE5
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   14
      Left            =   6840
      Picture         =   "Service.frx":F98BE
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   13
      Left            =   6360
      Picture         =   "Service.frx":FD297
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   12
      Left            =   5880
      Picture         =   "Service.frx":100C70
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   11
      Left            =   5400
      Picture         =   "Service.frx":104649
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   10
      Left            =   4920
      Picture         =   "Service.frx":108022
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   9
      Left            =   4440
      Picture         =   "Service.frx":10B9FB
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   8
      Left            =   3960
      Picture         =   "Service.frx":10F3D4
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   7
      Left            =   3480
      Picture         =   "Service.frx":112DAD
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   6
      Left            =   3000
      Picture         =   "Service.frx":116786
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   5
      Left            =   2520
      Picture         =   "Service.frx":11A15F
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   4
      Left            =   2040
      Picture         =   "Service.frx":11DB38
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   3
      Left            =   1560
      Picture         =   "Service.frx":121511
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   2
      Left            =   1080
      Picture         =   "Service.frx":124EEA
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   1
      Left            =   600
      Picture         =   "Service.frx":1288C3
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "Service.frx":12C29C
      Stretch         =   -1  'True
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label21 
      BackColor       =   &H001F130A&
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT  IMAGES"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF3E8&
      Height          =   345
      Left            =   4320
      TabIndex        =   40
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Image prdimage 
      Height          =   3375
      Left            =   4440
      Picture         =   "Service.frx":12FC75
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   7455
   End
   Begin VB.Label Label20 
      BackColor       =   &H001F130A&
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT    ID"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF3E8&
      Height          =   345
      Left            =   12360
      TabIndex        =   39
      Top             =   1845
      Width           =   1215
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00757575&
      X1              =   12120
      X2              =   12120
      Y1              =   1440
      Y2              =   9360
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00757575&
      X1              =   12120
      X2              =   19800
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label19 
      BackColor       =   &H00110609&
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF3E8&
      Height          =   255
      Left            =   15960
      TabIndex        =   38
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Shape Shape27 
      BorderColor     =   &H80000005&
      Height          =   255
      Left            =   15960
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackColor       =   &H00110609&
      BackStyle       =   0  'Transparent
      Caption         =   "AMOUNT"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF3E8&
      Height          =   255
      Left            =   17280
      TabIndex        =   37
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackColor       =   &H00110609&
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT NAME "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF3E8&
      Height          =   255
      Left            =   12360
      TabIndex        =   36
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Shape Shape26 
      BorderColor     =   &H80000005&
      Height          =   255
      Left            =   17280
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Shape Shape25 
      BorderColor     =   &H80000005&
      Height          =   255
      Left            =   12360
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label todaydate 
      BackColor       =   &H001F130A&
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/000"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF3E8&
      Height          =   465
      Left            =   5280
      TabIndex        =   35
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackColor       =   &H001F130A&
      BackStyle       =   0  'Transparent
      Caption         =   "DATE :"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   465
      Left            =   4200
      TabIndex        =   34
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackColor       =   &H001F130A&
      BackStyle       =   0  'Transparent
      Caption         =   "MODE OF PAY"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   4560
      TabIndex        =   33
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H001F130A&
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER ID"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   4560
      TabIndex        =   32
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H001F130A&
      BackStyle       =   0  'Transparent
      Caption         =   "SALES ID"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   4440
      TabIndex        =   31
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00110609&
      BackStyle       =   0  'Transparent
      Caption         =   "DISCOUNT %"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5BF4&
      Height          =   375
      Left            =   14400
      TabIndex        =   30
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Label disincre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
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
      Left            =   16200
      TabIndex        =   29
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label disdecre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   14160
      TabIndex        =   28
      Top             =   8610
      Width           =   255
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   16320
      X2              =   16320
      Y1              =   9000
      Y2              =   8880
   End
   Begin VB.Shape Shape24 
      BorderColor     =   &H80000005&
      Height          =   465
      Left            =   14715
      Top             =   8715
      Width           =   1185
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   14280
      X2              =   14280
      Y1              =   9000
      Y2              =   8880
   End
   Begin VB.Shape Shape23 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFF3E8&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   16200
      Top             =   8640
      Width           =   255
   End
   Begin VB.Shape Shape22 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFF3E8&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   14160
      Top             =   8640
      Width           =   255
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   14720
      X2              =   14280
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000004&
      X1              =   15885
      X2              =   16325
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Label Label12 
      BackColor       =   &H00110609&
      BackStyle       =   0  'Transparent
      Caption         =   "PRICE"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF3E8&
      Height          =   375
      Left            =   18360
      TabIndex        =   27
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00110609&
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT NAME AND ITS DETAILS"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF3E8&
      Height          =   375
      Left            =   12360
      TabIndex        =   26
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label clrallsalesbordbtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLEAR ALL"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15240
      TabIndex        =   22
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFF3E8&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   15240
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00110609&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL AMOUNT"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5BF4&
      Height          =   375
      Left            =   16680
      TabIndex        =   21
      Top             =   8280
      Width           =   3015
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000005&
      Height          =   375
      Left            =   16680
      Top             =   8640
      Width           =   3015
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00757575&
      Height          =   7935
      Left            =   4200
      Top             =   1440
      Width           =   15615
   End
   Begin VB.Shape Shape19 
      BorderColor     =   &H80000005&
      Height          =   8175
      Left            =   4080
      Top             =   1320
      Width           =   15855
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00110609&
      Caption         =   "SYSTEM FUNCTIONS"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5BF4&
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H001F130A&
      BackStyle       =   0  'Transparent
      Caption         =   "CLICK TO"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   720
      TabIndex        =   19
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H001F130A&
      BackStyle       =   0  'Transparent
      Caption         =   "CLICK TO"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   720
      TabIndex        =   18
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H001F130A&
      BackStyle       =   0  'Transparent
      Caption         =   "CLICK TO"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   720
      TabIndex        =   17
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H001F130A&
      BackStyle       =   0  'Transparent
      Caption         =   "CLICK TO"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   720
      TabIndex        =   16
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label delbtn 
      BackColor       =   &H00100A0A&
      BackStyle       =   0  'Transparent
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDF7F4&
      Height          =   495
      Left            =   2160
      TabIndex        =   15
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label updatebtn 
      BackColor       =   &H00100A0A&
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDF7F4&
      Height          =   495
      Left            =   2160
      TabIndex        =   14
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label savebtn 
      BackColor       =   &H00100A0A&
      BackStyle       =   0  'Transparent
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDF7F4&
      Height          =   495
      Left            =   2160
      TabIndex        =   13
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label newbtn 
      BackColor       =   &H00100A0A&
      BackStyle       =   0  'Transparent
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDF7F4&
      Height          =   495
      Left            =   2160
      TabIndex        =   12
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Shape Shape18 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   3795
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   135
   End
   Begin VB.Shape Shape17 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   105
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   135
   End
   Begin VB.Shape Shape16 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   3795
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   135
   End
   Begin VB.Shape Shape15 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   105
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   135
   End
   Begin VB.Shape Shape14 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   3795
      Shape           =   3  'Circle
      Top             =   3600
      Width           =   135
   End
   Begin VB.Shape Shape13 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   105
      Shape           =   3  'Circle
      Top             =   3600
      Width           =   135
   End
   Begin VB.Shape Shape12 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   3795
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   135
   End
   Begin VB.Shape Shape11 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   105
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   135
   End
   Begin VB.Shape Shape10 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   3795
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   135
   End
   Begin VB.Shape Shape9 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   105
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   135
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000B&
      X1              =   3855
      X2              =   3855
      Y1              =   1680
      Y2              =   5500
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000B&
      X1              =   165
      X2              =   165
      Y1              =   1680
      Y2              =   5500
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00F4D0C6&
      Caption         =   "FORM"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   340
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SALES DETAILS"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEEEE0&
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   180
      X2              =   3830
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   120
      Shape           =   3  'Circle
      Top             =   600
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   180
      X2              =   3830
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   100
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   100
      Left            =   120
      Shape           =   3  'Circle
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape delshape 
      FillColor       =   &H000B0405&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   240
      Top             =   4680
      Width           =   3555
   End
   Begin VB.Shape updateshape 
      FillColor       =   &H000B0405&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   240
      Top             =   3720
      Width           =   3555
   End
   Begin VB.Shape saveshape 
      FillColor       =   &H000B0405&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   240
      Top             =   2760
      Width           =   3555
   End
   Begin VB.Shape newshape 
      FillColor       =   &H000B0405&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   240
      Top             =   1800
      Width           =   3555
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H80000005&
      Height          =   255
      Left            =   15120
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label clrpartioption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17640
      TabIndex        =   23
      ToolTipText     =   "This clear the selected option from sales board"
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape Shape20 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFF3E8&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   17640
      Top             =   600
      Width           =   1815
   End
   Begin VB.Shape Shape21 
      BorderColor     =   &H80000005&
      Height          =   255
      Left            =   17520
      Top             =   960
      Width           =   2055
   End
   Begin VB.Shape Shape28 
      BorderColor     =   &H00404040&
      Height          =   3615
      Left            =   4320
      Top             =   5640
      Width           =   7695
   End
End
Attribute VB_Name = "Sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Dim totalamt As Long
Dim cnter As Integer
Dim status As Integer

Private Sub addtolistbtn_Click()
Dim qnt As Long
Dim qamt As Long

If prdnamebox.Text <> "" Then
    If prdqntbox.Text = "" Then
     qnt = 1
    Else
     qnt = Val(prdqntbox.Text)
    End If
    prdlist.AddItem (prdnamebox.Text)
    qtylist.AddItem (qnt)
    prdidlist.AddItem (prdidcombo.Text)
    qamt = qnt * Val(prdamtbox.Text)
    prdamtlist.AddItem (qamt)
    prdnamebox.Text = ""
    prdamtbox.Text = ""
    prdqntbox.Text = ""
    calculatetotalamt
End If
End Sub

Private Sub calculatetotalamt()
totalamt = 0
For i = 0 To prdlist.ListCount - 1
 totalamt = totalamt + Val(prdamtlist.List(i))
Next i
caldiscound
End Sub

Private Sub caldiscound()
Dim discountedamt As Long
 discountedamt = (Val(discountamt.Text) / 100) * totalamt
 totalamttopay = totalamt - discountedamt
totalamttopay.Text = totalamttopay
End Sub

Private Sub clrallsalesbordbtn_Click()
 prdlist.Clear
 prdamtlist.Clear
 qtylist.Clear
 prdidlist.Clear
 totalamttopay.Text = ""
End Sub

Private Sub clrpartioption_Click()
 Dim one As Integer
 If prdlist.ListIndex >= 0 Then
 one = prdlist.ListIndex
 prdlist.RemoveItem (one)
 prdamtlist.RemoveItem (one)
 qtylist.RemoveItem (one)
 Else
  MsgBox "Please Select Any Option", vbCritical, "Selection Error"
 End If
End Sub

Private Sub cstridcombo_Click()
If rs.State = 1 Then rs.Close
 rs.Open "select Cr_name from Customer_details where Cr_id = '" & cstridcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
  cstname.Text = rs.Fields(0)
 rs.Close
End Sub

Private Sub disdecre_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Line4.Y1 = 9050
 Line4.Y2 = 9050
 
 Line7.Y1 = 9050
 Line7.Y2 = 8930
 
 Shape22.Top = 8690
 
 disdecre.Top = 8660
 
 If IsNumeric(discountamt.Text) = False Then
  discountamt.Text = 0
 End If
 If Val(discountamt.Text) > 0 Then
  discountamt.Text = Val(discountamt.Text) - 1
 End If
 caldiscound
End Sub

Private Sub disdecre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Line4.Y1 = 9000
 Line4.Y2 = 9000
 
 Line7.Y1 = 9000
 Line7.Y2 = 8880
 
 Shape22.Top = 8640
 
  disdecre.Top = 8610
  
End Sub

Private Sub disincre_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Line3.Y1 = 9050
 Line3.Y2 = 9050
 
 Line8.Y1 = 9050
 Line8.Y2 = 8930
 
 Shape23.Top = 8690
 
 disincre.Top = 8660

 If IsNumeric(discountamt.Text) = False Then
  discountamt.Text = 0
 End If
 If Val(discountamt.Text) < 100 Then
  discountamt.Text = Val(discountamt.Text) + 1
 End If
 caldiscound
End Sub

Private Sub disincre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Line3.Y1 = 9000
 Line3.Y2 = 9000
 
 Line8.Y1 = 9000
 Line8.Y2 = 8880
 
 Shape23.Top = 8640
 
  disincre.Top = 8610
  
End Sub



Private Sub Form_Load()
connect
autosaleid
autosaleidload
datefinder
totalamt = 0
cnter = 0
status = 0
autocstidcomboload
autoprdidcomboload
autoimageload
End Sub

Private Sub autosaleidload()
salesidcombo.Clear
If rs.State = 1 Then rs.Close
 rs.Open "select * from Billing_table", con, adOpenKeyset, adLockOptimistic
 rs.MoveFirst
 While rs.EOF = False
  salesidcombo.AddItem (rs.Fields(0))
  rs.MoveNext
 Wend
rs.Close
End Sub

Private Sub autoprdidcomboload()
If rs.State = 1 Then rs.Close
 rs.Open "select Pr_id from Stock_table", con, adOpenKeyset, adLockOptimistic
 rs.MoveFirst
 While rs.EOF = False
  prdidcombo.AddItem (rs.Fields(0))
  rs.MoveNext
 Wend
rs.Close
End Sub

Private Sub autocstidcomboload()
If rs.State = 1 Then rs.Close
 rs.Open "select * from Customer_details", con, adOpenKeyset, adLockOptimistic
 rs.MoveFirst
 While rs.EOF = False
  cstridcombo.AddItem (rs.Fields(0))
  rs.MoveNext
 Wend
 rs.Close
End Sub

Private Sub autosaleid()
 Dim l As Integer, id As Integer
If rs.State = 1 Then rs.Close
 rs.Open "select * from Billing_table", con, adOpenKeyset, adLockOptimistic
  rs.MoveFirst
  l = 0
  While rs.EOF = False
   id = Val(Mid(rs.Fields(0), 5))
   If id > l Then
    l = id
   End If
   rs.MoveNext
  Wend
  l = l + 1
  salesidcombo.Text = "Sal-" & l
 rs.Close
End Sub

Private Sub autoimageload()
For i = 0 To 41
 Image2(i).Picture = LoadPicture("F:\Project46\MiniProject\Images\slashback.jpg")
 Image3(i).Picture = LoadPicture("F:\Project46\MiniProject\Images\slidefront.jpg")
Next i
End Sub

Private Sub datefinder()
todaydate.Caption = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Sub

Private Sub newbtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 newshape.FillColor = &H262626
 autosaleid
 clearallfields
cnter = 0
End Sub

Private Sub clearallfields()
cstridcombo.Text = ""
cstname.Text = ""
mopcombo.Text = ""
prdidcombo.Text = ""
clrallsalesbordbtn_Click
discountamt.Text = ""
prdimage.Picture = LoadPicture()
End Sub

Private Sub newbtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 newshape.FillColor = &HB0405
End Sub

Private Sub prdidcombo_Click()
If rs.State = 1 Then rs.Close
 rs.Open "select * from Stock_table where Pr_id = '" & prdidcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
  prdnamebox.Text = rs.Fields(1)
  prdamtbox.Text = rs.Fields(3) + rs.Fields(4)
  prdimage.Picture = LoadPicture("F:\Project46\MiniProject\Images\Views\" & rs.Fields(8))
rs.Close
End Sub

Private Sub prdlist_change()
    calculatetotalamt
    caldiscound
End Sub
Private Sub prdlist_Click()
Dim prdname As String
Dim prdindex As Integer

prdindex = prdlist.ListIndex
prdname = prdlist.List(prdindex)

If rs.State = 1 Then rs.Close
 rs.Open "select Pr_image from Stock_table where Pr_name = '" & prdname & "'", con, adOpenKeyset, adLockOptimistic
 prdimage.Picture = LoadPicture("F:\Project46\MiniProject\Images\Views\" & rs.Fields(0))
rs.Close
End Sub

Private Sub salesidcombo_Click()
clrallsalesbordbtn_Click
 If rs.State = 1 Then rs.Close
  rs.Open "select * from Billing_table where Sales_id = '" & salesidcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
   cstridcombo.Text = rs.Fields(1)
   mopcombo.Text = rs.Fields(2)
   todaydate.Caption = rs.Fields(3)
   discountamt.Text = rs.Fields(4)
   totalamttopay.Text = rs.Fields(5)
  rs.Close
  rs.Open "select * from Sales_table where Sales_id = '" & salesidcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
   While rs.EOF = False
    prdidlist.AddItem (rs.Fields(1))
    qtylist.AddItem (rs.Fields(2))
    rs.MoveNext
   Wend
  rs.Close
  For i = 0 To prdidlist.ListCount - 1
   rs.Open "select Pr_name,Pr_price,pr_tax from Stock_table where Pr_id = '" & prdidlist.List(i) & "'", con, adOpenKeyset, adLockOptimistic
    prdlist.List(i) = rs.Fields(0)
    prdamtlist.List(i) = rs.Fields(1) + rs.Fields(2)
   rs.Close
  Next i
  rs.Open "select Cr_name from Customer_details where Cr_id = '" & cstridcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
   cstname.Text = rs.Fields(0)
  rs.Close
  calculatetotalamt
End Sub

Private Sub Savebtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 saveshape.FillColor = &H262626
 checkallfields
 If status = 0 Then
  If rs.State = 1 Then rs.Close
   rs.Open "select * from Billing_table", con, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs.Fields(0) = salesidcombo.Text
    rs.Fields(1) = cstridcombo.Text
    rs.Fields(2) = mopcombo.Text
    rs.Fields(3) = todaydate.Caption
    rs.Fields(4) = Val(discountamt.Text)
    rs.Fields(5) = Val(totalamttopay.Text)
    rs.Update
   rs.Close
  For i = 0 To prdlist.ListCount - 1
  rs.Open "select * from Sales_table", con, adOpenKeyset, adLockOptimistic
   rs.AddNew
    rs.Fields(0) = salesidcombo.Text
    rs.Fields(1) = prdidcombo.List(i)
    rs.Fields(2) = qtylist.List(i)
   rs.Update
  rs.Close
  Next i
 '----------------------------------
 tyreanimation.Enabled = True
 savelabel.Visible = True
 autosaleidload
 End If
End Sub

Private Sub checkallfields()
 status = 0
 If cstridcombo.Text = "" Then
  notfillederror(0).Visible = True
  status = 1
 Else
  notfillederror(0).Visible = False
 End If
 If cstname.Text = "" Then
  notfillederror(1).Visible = True
  status = 1
 Else
  notfillederror(1).Visible = False
 End If
 If mopcombo.Text = "" Then
  notfillederror(2).Visible = True
 Else
  notfillederror(2).Visible = False
 End If
End Sub

Private Sub Savebtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 saveshape.FillColor = &HB0405
End Sub

Private Sub tyreanimation_Timer()
If cnter <= 41 Then
 Image2(cnter).Visible = True
 Image3(cnter).Visible = True
 cnter = cnter + 1
Else
 cnter = 0
 tyreanimation.Enabled = False
 tyreanimation2.Enabled = True
End If
End Sub

Private Sub tyreanimation2_Timer()
If cnter <= 41 Then
 Image2(cnter).Visible = False
 Image3(cnter).Visible = False
 cnter = cnter + 1
Else
 cnter = 0
 tyreanimation2.Enabled = False
 savelabel.Visible = False
End If
End Sub

Private Sub updatebtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 updateshape.FillColor = &H262626
  If rs.State = 1 Then rs.Close
   rs.Open "select * from Billing_table", con, adOpenKeyset, adLockOptimistic
    rs.Fields(1) = cstridcombo.Text
    rs.Fields(2) = mopcombo.Text
    rs.Fields(3) = todaydate.Caption
    rs.Fields(4) = Val(discountamt.Text)
    rs.Fields(5) = Val(totalamttopay.Text)
    rs.Update
   rs.Close
  For i = 0 To prdlist.ListCount - 1
  rs.Open "select * from Sales_table", con, adOpenKeyset, adLockOptimistic
    rs.Fields(1) = prdidcombo.List(i)
    rs.Fields(2) = qtylist.List(i)
   rs.Update
  rs.Close
  Next i
 '----------------------------------
 tyreanimation.Enabled = True
 savelabel.Visible = True
 autosaleidload
End Sub

Private Sub updatebtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 updateshape.FillColor = &HB0405
End Sub

Private Sub delbtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 delshape.FillColor = &H262626
 If rs.State = 1 Then rs.Close
  rs.Open "select * from Billing_table where Sales_id = '" & salesidcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
  rs.Delete
 rs.Close
  rs.Open "select * from Sales_table where Sales_id = '" & salesidcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
   rs.Delete
  rs.Close
  MsgBox "Datas Deleted"
  autosaleidload
End Sub

Private Sub delbtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 delshape.FillColor = &HB0405
End Sub
