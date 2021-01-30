VERSION 5.00
Begin VB.Form prddisp 
   Caption         =   "Form4"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16005
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   16005
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00100A05&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3255
      Left            =   240
      TabIndex        =   35
      Top             =   7560
      Width           =   6015
      Begin VB.Label removbtn 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   480
         TabIndex        =   37
         Top             =   1440
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Shape removebutton 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H005B3A11&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label buytitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Buy Now"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   480
         TabIndex        =   36
         Top             =   240
         Width           =   5175
      End
      Begin VB.Shape buybutton 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H005B3A11&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Timer timetimer 
      Interval        =   1000
      Left            =   5880
      Top             =   2040
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00100A05&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2895
      Left            =   240
      TabIndex        =   25
      Top             =   2760
      Width           =   6015
      Begin VB.Timer Timer8 
         Interval        =   1
         Left            =   5640
         Top             =   2520
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accessories"
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
         Height          =   495
         Left            =   3120
         TabIndex        =   28
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label catcylabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cycles"
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
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Shape Accesshp 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H005B3A11&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   3120
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Shape cycatshp 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00442C0D&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00180C89&
         Height          =   495
         Left            =   0
         TabIndex        =   26
         Top             =   120
         Width           =   6015
      End
      Begin VB.Shape pt4 
         BackColor       =   &H00231001&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00231001&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   3840
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Shape pt3 
         BackColor       =   &H00231001&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00231001&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1680
         Top             =   2280
         Width           =   375
      End
      Begin VB.Shape pt2 
         BackColor       =   &H00231001&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00231001&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   2160
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Shape pt1 
         BackColor       =   &H00231001&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00231001&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   12480
      Top             =   7320
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   12840
      Top             =   7320
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   13200
      Top             =   7320
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   13560
      Top             =   7320
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   13920
      Top             =   7320
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   14280
      Top             =   7320
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   14640
      Top             =   7320
   End
   Begin VB.Timer Timer9 
      Interval        =   10
      Left            =   8520
      Top             =   6840
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDD1AE&
      Height          =   735
      Left            =   4560
      TabIndex        =   38
      Top             =   1920
      Width           =   255
   End
   Begin VB.Line Line4 
      BorderColor     =   &H005B3A11&
      X1              =   240
      X2              =   6240
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00491C0E&
      X1              =   240
      X2              =   6240
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Label timdot 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDD1AE&
      Height          =   735
      Left            =   5280
      TabIndex        =   34
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label timesec 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FDD1AE&
      Height          =   735
      Left            =   5280
      TabIndex        =   33
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label timemin 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FDD1AE&
      Height          =   735
      Left            =   4560
      TabIndex        =   32
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label timehr 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FDD1AE&
      Height          =   1575
      Left            =   3480
      TabIndex        =   31
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00491C0E&
      X1              =   240
      X2              =   6240
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H005B3A11&
      X1              =   240
      X2              =   6240
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Formdetaillabel 
      BackStyle       =   0  'Transparent
      Caption         =   "This form show all details of Product in the shop"
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
      Left            =   240
      TabIndex        =   30
      Top             =   960
      Width           =   5895
   End
   Begin VB.Label formname 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Items"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0052A7F3&
      Height          =   615
      Left            =   240
      TabIndex        =   29
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label prdcatlabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Product Category:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00180C89&
      Height          =   495
      Left            =   13920
      TabIndex        =   24
      Top             =   10080
      Width           =   3375
   End
   Begin VB.Label prdcddesclabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Code Description:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00180C89&
      Height          =   495
      Left            =   6480
      TabIndex        =   23
      Top             =   10080
      Width           =   3255
   End
   Begin VB.Label prdstatuslabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Product Status:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00180C89&
      Height          =   495
      Left            =   14520
      TabIndex        =   22
      Top             =   8400
      Width           =   2775
   End
   Begin VB.Label prdCat 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEEFE4&
      Height          =   615
      Left            =   17520
      TabIndex        =   21
      Top             =   10080
      Width           =   3495
   End
   Begin VB.Label prdcddesc 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEEFE4&
      Height          =   615
      Left            =   9960
      TabIndex        =   20
      Top             =   10080
      Width           =   3975
   End
   Begin VB.Label Prdstatus 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEEFE4&
      Height          =   615
      Left            =   17520
      TabIndex        =   19
      Top             =   8400
      Width           =   3495
   End
   Begin VB.Label prdserial 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEEFE4&
      Height          =   615
      Left            =   17520
      TabIndex        =   18
      Top             =   9360
      Width           =   3495
   End
   Begin VB.Label prdcode 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEEFE4&
      Height          =   615
      Left            =   9960
      TabIndex        =   17
      Top             =   9240
      Width           =   3495
   End
   Begin VB.Label prdid 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEEFE4&
      Height          =   615
      Left            =   9960
      TabIndex        =   16
      Top             =   8400
      Width           =   3495
   End
   Begin VB.Label prdseriallabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Code:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00180C89&
      Height          =   615
      Left            =   14160
      TabIndex        =   15
      Top             =   9360
      Width           =   3135
   End
   Begin VB.Label prdcodelabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00180C89&
      Height          =   615
      Left            =   6600
      TabIndex        =   14
      Top             =   9240
      Width           =   3135
   End
   Begin VB.Label prdidlabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID :"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00180C89&
      Height          =   615
      Left            =   6600
      TabIndex        =   13
      Top             =   8400
      Width           =   3135
   End
   Begin VB.Label stocknumber 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B3A11&
      Height          =   1575
      Left            =   6840
      TabIndex        =   12
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label stocklabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Left"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00180C89&
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label prdname 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D76222&
      Height          =   495
      Left            =   6480
      TabIndex        =   10
      Top             =   7680
      Width           =   13575
   End
   Begin VB.Label pricetag 
      BackStyle       =   0  'Transparent
      Caption         =   "Price: "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   30.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00180C89&
      Height          =   735
      Left            =   10920
      TabIndex        =   9
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Index           =   6
      Left            =   16080
      TabIndex        =   8
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Index           =   5
      Left            =   15600
      TabIndex        =   7
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Index           =   4
      Left            =   15120
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Index           =   3
      Left            =   14640
      TabIndex        =   5
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Index           =   2
      Left            =   14160
      TabIndex        =   4
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Index           =   1
      Left            =   13680
      TabIndex        =   3
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Index           =   0
      Left            =   13200
      TabIndex        =   2
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000D72CC&
      Height          =   1335
      Left            =   18360
      TabIndex        =   1
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Image Image8 
      Height          =   5415
      Left            =   6600
      Stretch         =   -1  'True
      Top             =   720
      Width           =   13455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000D72CC&
      Height          =   975
      Left            =   6480
      TabIndex        =   0
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Image Image7 
      Height          =   330
      Left            =   7680
      Picture         =   "Product_view.frx":0000
      Stretch         =   -1  'True
      Top             =   6720
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image Image6 
      Height          =   570
      Left            =   7200
      Picture         =   "Product_view.frx":2E27
      Stretch         =   -1  'True
      Top             =   6600
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image Image5 
      Height          =   855
      Left            =   6480
      Picture         =   "Product_view.frx":5E1B
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image4 
      Height          =   330
      Left            =   18480
      Picture         =   "Product_view.frx":90EA
      Stretch         =   -1  'True
      Top             =   6720
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   18720
      Picture         =   "Product_view.frx":BF10
      Stretch         =   -1  'True
      Top             =   6600
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   19200
      Picture         =   "Product_view.frx":EEFF
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   11535
      Left            =   0
      Picture         =   "Product_view.frx":121B9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20535
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   15120
      Top             =   840
      Width           =   4935
   End
End
Attribute VB_Name = "prddisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nextbtncntr As Integer
Dim prdimgcntr As Integer
Dim ar(10) As Integer
Dim length
Dim z As Long
Dim totalimg As Integer
Dim front1 As Integer
Dim front2 As Integer
Dim front3 As Integer
Dim front4 As Integer


Private Sub buytitle_Click()
If removebutton.Visible = False Then
 removebutton.Visible = True
 removbtn.Visible = True
 buybutton.FillColor = &H442C0D
End If
End Sub

Private Sub catcylabel_Click()
cycatshp.FillColor = &H442C0D
Accesshp.FillColor = &H5B3A11
If rs.State = 1 Then rs.Close
 rs.Open "select * from Stock_table s,Product_table p where s.Pr_id = p.Pr_id", con, adOpenKeyset, adLockOptimistic '<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>
  rs2.Open "select * from ", con, adOpenKeyset, adLockOptimistic '<<<<<<<<<<<<<<<<<<<<<<<<check here>>>>>>>>>>>>>>>>>>>>>>.
 displayallfields
End Sub

Private Sub Form_Load()
connect
prdimgcntr = 0
nextbtncntr = 0
totalimg = 6
front1 = 1
front2 = 1
front3 = 1
front4 = 1
If rs.State = 1 Then rs.Close
 rs.Open "select * from Stock_table s,Product_table p,Product_code c,Product_category e where s.Pr_id = p.Pr_id and s.Pr_code = c.Pr_code and c.Cat_id=e.Cat_id", con, adOpenKeyset, adLockOptimistic
End Sub




Private Sub Label4_Click()
'------------------------------
If rs.EOF = False Then
rs.MoveNext
displayallfields
Else
MsgBox "No More Records"
End If
End Sub

Private Sub Label2_Click()
'------------------------------------
If rs.BOF = False Then
rs.MovePrevious
displayallfields
Else
MsgBox "No More Records"
End If
End Sub

Private Sub displayallfields()
 If rs.EOF = False Then
 Image8.Picture = LoadPicture("F:\Project46\MiniProject\Images\Views\" & rs.Fields(8))
 z = rs.Fields(3) + rs.Fields(4)        '<--------------Change here to put price
 Call changetozero
 findprice (z)
 timer1.Enabled = True
 stocknumber.Caption = rs.Fields(5)
 prdid.Caption = rs.Fields(0)
 prdname.Caption = rs.Fields(1)
 prdcode.Caption = rs.Fields(2)
 'Prdstatus.Caption = rs.Fields(11)
 'prdserial.Caption = rs.Fields(10)
 'prdcddesc.Caption = rs.Fields(14)
 'prdCat.Caption = rs.Fields(17)
 End If
End Sub

Private Sub Label5_Click()
cycatshp.FillColor = &H5B3A11
Accesshp.FillColor = &H442C0D
If rs.State = 1 Then rs.Close
 rs.Open "select * from Stock_table s,Product_table p,Product_code c,Product_category e where s.Pr_id = p.Pr_id and s.Pr_code = c.Pr_code and c.Cat_id = 'C2'", con, adOpenKeyset, adLockOptimistic
 displayallfields
End Sub

Private Sub removbtn_Click()
removebutton.Visible = False
buybutton.FillColor = &H5B3A11
removbtn.Visible = False
End Sub

Private Sub Timer8_Timer()

If front1 = 1 Then
 If pt1.Left <= Frame1.Width Then
  pt1.Left = pt1.Left + Rnd(10) * 100
 Else
  front1 = 0
 End If
Else
 If pt1.Left >= Frame1.Left Then
  pt1.Left = pt1.Left - Rnd(10) * 100
 Else
  front1 = 1
 End If
End If  'first one

If front2 = 1 Then
 If pt2.Left <= Frame1.Width Then
  pt2.Left = pt2.Left + Rnd(10) * 100
 Else
  front2 = 0
 End If
Else
 If pt2.Left >= Frame1.Left - 1000 Then
  pt2.Left = pt2.Left - Rnd(10) * 100
 Else
  front2 = 1
 End If
End If  'second one

If front3 = 1 Then
 If pt3.Left <= Frame1.Width Then
  pt3.Left = pt3.Left + Rnd(10) * 100
 Else
  front3 = 0
 End If
Else
 If pt3.Left >= Frame1.Left - 2000 Then
  pt3.Left = pt3.Left - Rnd(10) * 100
 Else
  front3 = 1
 End If
End If  'third one

If front4 = 1 Then
 If pt4.Left <= Frame1.Width Then
  pt4.Left = pt4.Left + Rnd(10) * 200
 Else
  front4 = 0
 End If
Else
 If pt4.Left >= Frame1.Left - 2000 Then
  pt4.Left = pt4.Left - Rnd(10) * 200
 Else
  front4 = 1
 End If
End If  'fourth one

End Sub

Private Sub Timer9_Timer()
nextbtncntr = nextbtncntr + 1
If nextbtncntr > 10 And nextbtncntr < 20 Then
 Image4.Visible = True
 Image7.Visible = True
End If
If nextbtncntr > 20 And nextbtncntr < 30 Then
 Image3.Visible = True
 Image6.Visible = True
End If
If nextbtncntr > 30 And nextbtncntr < 40 Then
 Image2.Visible = True
 Image5.Visible = True
End If
If nextbtncntr > 50 Then
 nextbtncntr = 0
 Image4.Visible = False
 Image3.Visible = False
 Image2.Visible = False
 Image5.Visible = False
 Image6.Visible = False
 Image7.Visible = False
End If
End Sub

Private Sub timer1_Timer()
Dim p As Integer
p = Rnd(10) * 10
If p = ar(0) Then
   If length = 1 Then
 Shape1.FillColor = &H0&
 End If
 timer1.Enabled = False
 Timer2.Enabled = True
End If
Label1(0).Caption = p
End Sub

Private Sub Timer2_Timer()
Dim p As Integer
p = Rnd(10) * 10
If p = ar(1) Then
 Timer2.Enabled = False
 Timer3.Enabled = True
  If length = 2 Then
 Shape1.FillColor = &H0&
 End If
End If
Label1(1).Caption = p
End Sub

Private Sub Timer3_Timer()
Dim p As Integer
p = Rnd(10) * 10
If p = ar(2) Then
 Timer3.Enabled = False
 Timer4.Enabled = True
  If length = 3 Then
 Shape1.FillColor = &H0&
 End If
End If
Label1(2).Caption = p
End Sub

Private Sub Timer4_Timer()
Dim p As Integer
p = Rnd(10) * 10
If p = ar(3) Then
 Timer4.Enabled = False
 Timer5.Enabled = True
  If length = 4 Then
 Shape1.FillColor = &H0&
 End If
End If
Label1(3).Caption = p
End Sub

Private Sub Timer5_Timer()
Dim p As Integer
p = Rnd(10) * 10
If p = ar(4) Then
 Timer5.Enabled = False
 Timer6.Enabled = True
 If length = 5 Then
 Shape1.FillColor = &H0&
 End If
End If
Label1(4).Caption = p
End Sub

Private Sub Timer6_Timer()
Dim p As Integer
p = Rnd(10) * 10
If p = ar(5) Then
 Timer6.Enabled = False
 Timer7.Enabled = True
 If length = 6 Then
 Shape1.FillColor = &H0&
 End If
End If
Label1(5).Caption = p
End Sub

Private Sub Timer7_Timer()
Dim p As Integer
p = Rnd(10) * 10
If p = ar(6) Then
 Timer7.Enabled = False
 If length = 7 Then
 Shape1.FillColor = &H0&
 End If
End If
Label1(6).Caption = p
End Sub

Function changetozero()
For i = 0 To 6
 Label1(i) = "0"
Next i
Shape1.FillColor = &H404040
End Function

Function findprice(k)
length = Len(k)
Labelalloc (length)
Call valalloc(length, k)
End Function

Function Labelalloc(X)
For i = 0 To X - 1
 Label1(i).Visible = True
Next i
For i = X To 7 - 1
 Label1(i).Visible = False
Next i
End Function

Function valalloc(X, Y)
For i = 0 To X - 1
 ar(i) = Val(Mid(Y, i + 1, 1))
Next i
End Function

Private Sub timetimer_Timer()
timehr = Hour(Time)
timemin = Minute(Time)
timesec = Second(Time)
End Sub
