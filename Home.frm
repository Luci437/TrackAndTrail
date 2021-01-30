VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form StockEntry 
   BackColor       =   &H00000000&
   Caption         =   "Home"
   ClientHeight    =   9690
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19170
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9690
   ScaleWidth      =   19170
   WindowState     =   2  'Maximized
   Begin VB.ComboBox prdspecificcodecombo 
      Appearance      =   0  'Flat
      BackColor       =   &H00262626&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0096EF9E&
      Height          =   450
      ItemData        =   "Home.frx":0000
      Left            =   4080
      List            =   "Home.frx":0002
      TabIndex        =   45
      Top             =   7680
      Width           =   3495
   End
   Begin VB.PictureBox dop 
      Height          =   495
      Left            =   4080
      ScaleHeight     =   435
      ScaleWidth      =   3435
      TabIndex        =   44
      Top             =   9120
      Width           =   3495
   End
   Begin VB.ComboBox prdcatidcombo 
      Appearance      =   0  'Flat
      BackColor       =   &H00262626&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0096EF9E&
      Height          =   450
      Left            =   11520
      TabIndex        =   32
      Top             =   3720
      Width           =   3495
   End
   Begin VB.ComboBox prdstatuscombo 
      Appearance      =   0  'Flat
      BackColor       =   &H00262626&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0096EF9E&
      Height          =   450
      ItemData        =   "Home.frx":0004
      Left            =   4080
      List            =   "Home.frx":000E
      TabIndex        =   31
      Top             =   8400
      Width           =   3495
   End
   Begin VB.ComboBox prdcodecombo 
      Appearance      =   0  'Flat
      BackColor       =   &H00262626&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0096EF9E&
      Height          =   450
      Left            =   11520
      TabIndex        =   30
      Top             =   1680
      Width           =   3495
   End
   Begin VB.ComboBox prdcomboid 
      Appearance      =   0  'Flat
      BackColor       =   &H00262626&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0096EF9E&
      Height          =   450
      Left            =   4080
      TabIndex        =   29
      Top             =   1680
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   18600
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00262626&
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
      ForeColor       =   &H0096EF9E&
      Height          =   495
      Index           =   8
      Left            =   11520
      TabIndex        =   23
      Top             =   4440
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00262626&
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
      ForeColor       =   &H0096EF9E&
      Height          =   495
      Index           =   7
      Left            =   11520
      TabIndex        =   22
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00262626&
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
      ForeColor       =   &H0096EF9E&
      Height          =   495
      Index           =   6
      Left            =   4560
      TabIndex        =   21
      Top             =   6000
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00262626&
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
      ForeColor       =   &H0096EF9E&
      Height          =   495
      Index           =   5
      Left            =   4560
      TabIndex        =   20
      Top             =   5280
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00262626&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0096EF9E&
      Height          =   495
      Index           =   4
      Left            =   4080
      TabIndex        =   19
      Text            =   "0"
      Top             =   4560
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00262626&
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
      ForeColor       =   &H0096EF9E&
      Height          =   495
      Index           =   3
      Left            =   4080
      TabIndex        =   18
      Top             =   3840
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00262626&
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
      ForeColor       =   &H0096EF9E&
      Height          =   495
      Index           =   2
      Left            =   4080
      TabIndex        =   17
      Top             =   3120
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00262626&
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
      ForeColor       =   &H0096EF9E&
      Height          =   495
      Index           =   1
      Left            =   4080
      TabIndex        =   16
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label errorchecker 
      Alignment       =   2  'Center
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
      Index           =   12
      Left            =   15120
      TabIndex        =   50
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label dec2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   495
      Left            =   7080
      TabIndex        =   49
      Top             =   5880
      Width           =   495
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   7080
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label decval2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   495
      Left            =   4080
      TabIndex        =   48
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label incval1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   495
      Left            =   7080
      TabIndex        =   47
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label decval1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   495
      Left            =   4080
      TabIndex        =   46
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   7080
      Top             =   5280
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4080
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label errorchecker 
      Alignment       =   2  'Center
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
      Index           =   11
      Left            =   15120
      TabIndex        =   43
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label errorchecker 
      Alignment       =   2  'Center
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
      Index           =   10
      Left            =   7680
      TabIndex        =   42
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label errorchecker 
      Alignment       =   2  'Center
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
      Index           =   9
      Left            =   7680
      TabIndex        =   41
      Top             =   7680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label errorchecker 
      Alignment       =   2  'Center
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
      Index           =   8
      Left            =   15120
      TabIndex        =   40
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label errorchecker 
      Alignment       =   2  'Center
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
      Index           =   7
      Left            =   15120
      TabIndex        =   39
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label errorchecker 
      Alignment       =   2  'Center
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
      Left            =   7680
      TabIndex        =   38
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label errorchecker 
      Alignment       =   2  'Center
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
      Left            =   7680
      TabIndex        =   37
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label errorchecker 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Enabled         =   0   'False
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
      Left            =   7680
      TabIndex        =   36
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label errorchecker 
      Alignment       =   2  'Center
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
      Left            =   7680
      TabIndex        =   35
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label errorchecker 
      Alignment       =   2  'Center
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
      Left            =   7680
      TabIndex        =   34
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label errorchecker 
      Alignment       =   2  'Center
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
      Left            =   7680
      TabIndex        =   33
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label newbtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17760
      TabIndex        =   28
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   17760
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Savebtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17760
      TabIndex        =   26
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label updatebtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17760
      TabIndex        =   27
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   17760
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   17760
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label delbtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17760
      TabIndex        =   25
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   17760
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label addimgbtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add Image"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E3A5A2&
      Height          =   375
      Left            =   17760
      TabIndex        =   24
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00442C0D&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   17760
      Shape           =   4  'Rounded Rectangle
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00442C0D&
      X1              =   11400
      X2              =   11400
      Y1              =   3720
      Y2              =   4920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00442C0D&
      X1              =   11400
      X2              =   11400
      Y1              =   1680
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00442C0D&
      X1              =   3960
      X2              =   3960
      Y1              =   7680
      Y2              =   9600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00442C0D&
      X1              =   3960
      X2              =   3960
      Y1              =   1680
      Y2              =   6480
   End
   Begin VB.Label prdcatlable 
      BackStyle       =   0  'Transparent
      Caption         =   "Category Name"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   8640
      TabIndex        =   15
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label prdcatidlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Category ID"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   8640
      TabIndex        =   14
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label prdcodedislabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Code Description"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   8640
      TabIndex        =   13
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label prddoplable 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Purchase"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   9120
      Width           =   3015
   End
   Begin VB.Label prdstatuslabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Status"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   8400
      Width           =   3015
   End
   Begin VB.Label prdpartidlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Specific ID"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   7680
      Width           =   3015
   End
   Begin VB.Label prdinstklable 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Initial Stock"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   6000
      Width           =   3015
   End
   Begin VB.Label prdrollable 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Reload Limit"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label prdstocklabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Stock"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Label prdtaxlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Product GST"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label prdpricelabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Price"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label prdcodelabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label prdnamelabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label prdidlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID    "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "In this form Admin can add details of new products"
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
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Products Details"
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
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00442C0D&
      X1              =   120
      X2              =   5895
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FDD1AE&
      X1              =   120
      X2              =   5895
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Shape Shape8 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4080
      Top             =   6000
      Width           =   495
   End
   Begin VB.Image prdimage 
      Height          =   6615
      Left            =   6720
      Picture         =   "Home.frx":0021
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   13815
   End
End
Attribute VB_Name = "StockEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim status As Integer
Dim imgpath As String


'PRODUCT IMAGE ADDING BUTTON
Private Sub addimgbtn_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "JPG|*jpg"
prdimage.Picture = LoadPicture(CommonDialog1.FileName)
imgpath = CommonDialog1.FileTitle
MsgBox (imgpath)
End Sub

Private Sub dec2_Click()
If Text1(6).Text = "" Then
 Text1(6).Text = 0
Else
Text1(6).Text = Val(Text1(6).Text) + 1
End If
End Sub

Private Sub decval1_Click()
If Text1(5).Text = "0" Or Text1(5).Text = "" Then
 Text1(5).Text = 0
Else
Text1(5).Text = Val(Text1(5).Text) - 1
End If
End Sub

Private Sub decval2_Click()
If Text1(6).Text = "0" Or Text1(6).Text = "" Then
 Text1(6).Text = 0
Else
Text1(6).Text = Val(Text1(6).Text) - 1
End If
End Sub

Private Sub delbtn_Click()
If rs.State = 1 Then rs.Close
 rs.Open "select * from Product_table where Pr_id = '" & prdcomboid.Text & "'", con, adOpenKeyset, adLockOptimistic
 rs.Delete
 rs.Close
 '---------------
 rs.Open "select * from Stock_table where Pr_id = '" & prdcomboid.Text & "'", con, adOpenKeyset, adLockOptimistic
  rs.Delete
 rs.Close
 rs.Open "select * from Product_code p,Stock_table s where p.Pr_code = s.Pr_code", con, adOpenKeyset, adLockOptimistic
  rs.Delete
 rs.Close
 MsgBox "Record Deleted"
 newbtn_Click
 Call autopr
 Call autoparticularid
End Sub

'FORM LOAD
Private Sub Form_Load()
connect
status = 0
Call autopr
Call autoparticularid
Call prdcomboload
Call prdcodecomboautoload
Call prdspecificcodecomboload
Call prdcatidcomboload
End Sub
'PRODUCT CATEGORY ID AUTO LOAD
Private Sub prdcatidcomboload()
prdcatidcombo.Clear
If rs.State = 1 Then rs.Close
 rs.Open "select * from Product_category", con, adOpenKeyset, adLockOptimistic
  rs.MoveFirst
   While rs.EOF = False
   prdcatidcombo.AddItem (rs.Fields(0))
   rs.MoveNext
   Wend
 rs.Close
End Sub
'PRODUCT CODE COMBO LOAD
Private Sub prdcodecomboautoload()
If rs.State = 1 Then rs.Close
rs.Open "select Pr_code from Product_code", con, adOpenKeyset, adLockOptimistic
rs.MoveFirst
While rs.EOF = False
 prdcodecombo.AddItem (rs.Fields(0))
 rs.MoveNext
Wend
rs.Close
End Sub

'PRODUCT ID COMBO LOAD
Private Sub prdcomboload()
If rs.State = 1 Then rs.Close
rs.Open "select Pr_id from Stock_table", con, adOpenKeyset, adLockOptimistic
If rs.RecordCount = 0 Then

Else
rs.MoveFirst
While rs.EOF = False
 prdcomboid.AddItem (rs.Fields(0))
 rs.MoveNext
Wend
rs.Close
End If
End Sub

Private Sub Label3_click()
If Text1(5).Text = "0" Or Text1(5).Text = "" Then
 Text1(5).Text = 0
Else
Text1(5).Text = Val(Text1(5).Text) - 1
End If
End Sub

Private Sub incval1_Click()
If Text1(5).Text = "" Then
 Text1(5).Text = 0
Else
Text1(5).Text = Val(Text1(5).Text) + 1
End If
End Sub

'NEW BUTTON
Private Sub newbtn_Click()
prdcomboid.Clear
prdspecificcodecombo.Clear
Call clearfields
autopr
Call prdcomboload
Call autoparticularid
prdimage.Picture = LoadPicture("F:\Project46\MiniProject\Images\Products\Cyclebgi2.jpg")
End Sub

'AUTO INCREMENT CODE
Public Sub autopr()
Dim stname As String
Dim stno As Integer
Dim lrg As Integer
If rs.State = 1 Then rs.Close
rs.Open "select * from Stock_table", con, adOpenKeyset, adLockOptimistic
If rs.RecordCount = 0 Then
 prdcomboid.Text = "Pr1"
Else
lrg = 0
rs.MoveFirst
While rs.EOF = False
stno = Val(Mid(rs.Fields(0), 3))
If stno > lrg Then
 lrg = stno
End If
rs.MoveNext
Wend

lrg = lrg + 1
prdcomboid.Text = "Pr" & lrg
rs.Close
End If
End Sub

'TO CLEAR ALL FIELDS
Private Sub clearfields()
For i = 1 To 8
Text1(i).Text = ""
Next i
prdstatuscombo.Text = ""
prdcodecombo.Text = ""
prdcatidcombo.Text = ""
End Sub

'PRODUCT CATEGORY ID LOAD
Private Sub prdcatidcombo_Click()
If rs.State = 1 Then rs.Close
 rs.Open "select * from Product_category where Cat_id = '" & prdcatidcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
   Text1(8).Text = rs.Fields(1)
 rs.Close
End Sub

'PRODUCT CODE COMBO LOAD
Private Sub prdcodecombo_click()
If rs.State = 1 Then rs.Close
 rs.Open "select * from Product_code where Pr_code = '" & prdcodecombo.Text & "'", con, adOpenKeyset, adLockOptimistic

  Text1(7).Text = rs.Fields(1)
 rs.Close
 stockcounter
End Sub

Private Sub prdcomboid_Click()
If rs.State = 1 Then rs.Close
 rs.Open "select * from Stock_table where Pr_id = '" & prdcomboid.Text & "'", con, adOpenKeyset, adLockOptimistic
 With rs
   Text1(1).Text = rs.Fields(1)
   Text1(2).Text = rs.Fields(3)
   Text1(3).Text = rs.Fields(4)
   Text1(4).Text = rs.Fields(5)
   Text1(5).Text = rs.Fields(6)
   Text1(6).Text = rs.Fields(7)
   prdcodecombo.Text = rs.Fields(2)
   prdimage.Picture = LoadPicture("F:\Project46\MiniProject\Images\Products\" & rs.Fields(8) & "")
 End With
 rs.Close
 '------------------------------
 rs.Open "select P_id from Product_table where Pr_id = '" & prdcomboid.Text & "'", con, adOpenKeyset, adLockOptimistic
 If rs.RecordCount <> 0 Then
 prdspecificcodecombo.Text = rs.Fields(0)
 Else
 prdspecificcodecombo.Text = ""
 End If
 rs.Close
 '-------------------------------
 rs.Open "select * from Product_code p, Product_category c where p.Cat_id = c.Cat_id", con, adOpenKeyset, adLockOptimistic
  prdcatidcombo.Text = rs.Fields(2)
  Text1(8).Text = rs.Fields(4)
 rs.Close
 '-------------------------------
 rs.Open "select Pr_status from Product_table where P_id = '" & prdspecificcodecombo.Text & "' ", con, adOpenKeyset, adLockOptimistic
  prdstatuscombo.Text = rs.Fields(0)
 rs.Close
 '-------------------------------
 rs.Open "select * from Product_code where Pr_code = '" & prdcodecombo.Text & "' ", con, adOpenKeyset, adLockOptimistic
  Text1(7).Text = rs.Fields(1)
 rs.Close
 '-------------------------------
 Call stockcounter
End Sub

Private Sub stockcounter()
 Dim totalstock As Integer
 If rs.State = 1 Then rs.Close
  rs.Open "select Pr_code from Product_code", con, adOpenKeyset, adLockOptimistic
  If rs.RecordCount = 0 Then
  Text1(4).Text = 0
  Else
   rs.MoveFirst
   totalstock = 0
    While rs.EOF = False
     If prdcodecombo.Text = rs.Fields(0) Then
      totalstock = totalstock + 1
     End If
     rs.MoveNext
    Wend
  Text1(4).Text = totalstock
  rs.Close
  End If
End Sub

Private Sub prdspecificcodecomboload()
 If rs.State = 1 Then rs.Close
  rs.Open "select P_id from Product_table", con, adOpenKeyset, adLockOptimistic
  If rs.RecordCount = 0 Then
  
  Else
   rs.MoveFirst
   While rs.EOF = False
    prdspecificcodecombo.AddItem (rs.Fields(0))
    rs.MoveNext
   Wend
  rs.Close
  End If
End Sub

'PRODUCT SPECIFIC CODE LOAD
Private Sub prdspecificcodecombo_click()
 If rs.State = 1 Then rs.Close
 rs.Open "select * from Product_table where P_id='" & prdspecificcodecombo.Text & "'", con, adOpenKeyset, adLockOptimistic
 prdstatuscombo.Text = rs.Fields(2)
 rs.Close
End Sub

'SAVE BUTTON
Private Sub savebtn_Click()
Call checkallfields
    If status = 0 Then
     If rs.State = 1 Then rs.Close
      rs.Open "select * from Stock_table", con, adOpenKeyse, adLockOptimistic
       With rs
        .AddNew
        .Fields(0) = prdcomboid.Text
        .Fields(1) = Text1(1).Text
        .Fields(2) = prdcodecombo.Text
        .Fields(3) = Text1(2).Text
        .Fields(4) = Text1(3).Text
        .Fields(5) = Text1(4).Text
        .Fields(6) = Text1(5).Text
        .Fields(7) = Text1(6).Text
        .Fields(8) = imgpath
        .Update
       End With
      rs.Close
      '---------------PRODUCT CODE RECORD
     rs.Open "select * from Product_code", con, adOpenKeyset, adLockOptimistic
      With rs
       .AddNew
       .Fields(0) = prdcodecombo.Text
       .Fields(1) = Text1(7).Text
       .Fields(2) = prdcatidcombo.Text
       .Update
      End With
      rs.Close
      '---------------PRODUCT CATEGORY RECORD
     'rs.Open "select * from Product_category", con, adOpenKeyset, adLockOptimistic
     '  With rs
     '   .AddNew
     '   .Fields(0) = prdcatidcombo.Text
     '   .Fields(1) = Text1(8).Text
     '   .Update
     '  End With
     ' rs.Close
      '--------------PRODUCT TABLE
      rs.Open "select * from Product_table", con, adOpenKeyset, adLockOptimistic
       With rs
        .AddNew
        .Fields(0) = prdcomboid.Text
        .Fields(1) = prdspecificcodecombo.Text
        .Fields(2) = prdstatuscombo.Text
        '.Fields(3) = dop.Value
        .Update
       End With
      rs.Close
      MsgBox "Record Added"
      Call prdcomboid_Click
    End If 'status if
End Sub

'FIELD CHECKING()
Private Sub checkallfields()
status = 0
    For i = 1 To 8
     If i <> 4 Then
      If Text1(i).Text = "" Then
      errorchecker(i).Visible = True
      status = 1
      Else
      errorchecker(i).Visible = False
      End If
    End If
    Next i
    
    If prdspecificcodecombo.Text = "" Then
     errorchecker(9).Visible = True
     status = 1
    Else
     errorchecker(9).Visible = False
    End If

    If prdstatuscombo.Text = "" Then
     errorchecker(10).Visible = True
     status = 1
    Else
     errorchecker(10).Visible = False
    End If
    
    If prdcodecombo.Text = "" Then
     errorchecker(11).Visible = True
     status = 1
    Else
     errorchecker(11).Visible = False
    End If

    If prdcatidcombo.Text = "" Then
     errorchecker(12).Visible = True
     status = 1
    Else
     errorchecker(12).Visible = False
    End If
    
End Sub
'AUTO PARTICULAR PRODUCT ID
Private Sub autoparticularid()
 Dim l As Integer, id As Integer
 
 If rs.State = 1 Then rs.Close
 rs.Open "select * from Product_table", con, adOpenKeyset, adLockOptimistic
 If rs.RecordCount = 0 Then
  prdspecificcodecombo.Text = "P1"
 Else
 rs.MoveFirst
 l = 0
 
 While rs.EOF = False
 id = Val(Mid(rs.Fields(1), 2))
    If l < id Then
     l = id
    End If
    rs.MoveNext
 Wend
 l = l + 1
 prdspecificcodecombo.Text = "P" & l
 rs.Close
 End If
End Sub

Private Sub updatebtn_Click()
 If rs.State = 1 Then rs.Close
  rs.Open "select * from Stock_table where Pr_id = '" & prdcomboid.Text & "'", con, adOpenKeyset, adLockOptimistic
   rs.Fields(1) = Text1(1).Text
   rs.Fields(2) = prdcodecombo.Text
   rs.Fields(3) = Text1(2).Text
   rs.Fields(4) = Text1(3).Text
   rs.Fields(5) = Text1(4).Text
   rs.Fields(6) = Text1(5).Text
   rs.Fields(7) = Text1(6).Text
   If imgpath <> "" Then
   rs.Fields(8) = imgpath
   End If
 rs.Update
 rs.Close
 rs.Open "select * from Product_table where Pr_id = '" & prdcomboid.Text & "'", con, adOpenKeyset, adLockOptimistic
  rs.Fields(2) = prdstatuscombo.Text
  'rs.Fields(3) = DATE OF PURCHASE
  rs.Update
 rs.Close
 rs.Open "select * from Product_code p,Stock_table s where p.Pr_code = s.Pr_code", con, adOpenKeyset, adLockOptimistic
  rs.Fields(0) = prdcodecombo.Text
  rs.Fields(1) = Text1(7).Text
  rs.Fields(2) = prdcatidcombo.Text
  rs.Update
 rs.Close
 MsgBox "Record Update"
End Sub
