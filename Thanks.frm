VERSION 5.00
Begin VB.Form Thanks 
   BackColor       =   &H000A0000&
   Caption         =   "Vote of Thanks"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer lightstyle5 
      Enabled         =   0   'False
      Interval        =   24
      Left            =   5880
      Top             =   1080
   End
   Begin VB.Timer lightstyle4 
      Enabled         =   0   'False
      Interval        =   24
      Left            =   5520
      Top             =   1080
   End
   Begin VB.Timer lightstyle3 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   5160
      Top             =   1080
   End
   Begin VB.Timer selectstyle 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4800
      Top             =   600
   End
   Begin VB.Timer lightstyle2 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   4800
      Top             =   1080
   End
   Begin VB.Frame tempdisp 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6615
      Left            =   4440
      TabIndex        =   4
      Top             =   1920
      Width           =   12735
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Creators of this software are Abijith M A, Athul Krishna ES,Kamal E J and MidhunLal P H"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2640
         Width           =   12375
      End
   End
   Begin VB.Timer lightstyle1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4440
      Top             =   1080
   End
   Begin VB.Timer textani 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   8760
      Top             =   8640
   End
   Begin VB.Timer disptext 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8280
      Top             =   8640
   End
   Begin VB.Timer dispcolor 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7800
      Top             =   8640
   End
   Begin VB.Frame display 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5655
      Left            =   4920
      TabIndex        =   2
      Top             =   2400
      Width           =   11775
      Begin VB.Label disptextlabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   51.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00271203&
         Height          =   1455
         Left            =   0
         TabIndex        =   3
         Top             =   1800
         Width           =   11655
      End
   End
   Begin VB.Timer fretimerselecter 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   720
      Top             =   10680
   End
   Begin VB.Timer fretimer 
      Enabled         =   0   'False
      Interval        =   24
      Left            =   240
      Top             =   10680
   End
   Begin VB.Frame freframe 
      BackColor       =   &H00C7C7C7&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   180
      TabIndex        =   0
      Top             =   9660
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Line frelines 
         Index           =   26
         X1              =   0
         X2              =   960
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line frelines 
         Index           =   25
         X1              =   0
         X2              =   960
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line frelines 
         Index           =   24
         X1              =   0
         X2              =   960
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line frelines 
         Index           =   23
         X1              =   0
         X2              =   960
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line frelines 
         Index           =   22
         X1              =   0
         X2              =   960
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line frelines 
         Index           =   21
         X1              =   0
         X2              =   960
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line frelines 
         Index           =   20
         X1              =   1200
         X2              =   2160
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line frelines 
         Index           =   19
         X1              =   0
         X2              =   960
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line frelines 
         Index           =   18
         X1              =   2280
         X2              =   3240
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line frelines 
         Index           =   17
         X1              =   2280
         X2              =   3240
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line frelines 
         Index           =   16
         X1              =   2280
         X2              =   3240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line frelines 
         Index           =   15
         X1              =   2280
         X2              =   3240
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line frelines 
         Index           =   14
         X1              =   2280
         X2              =   3240
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line frelines 
         Index           =   13
         X1              =   2280
         X2              =   3240
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line frelines 
         Index           =   12
         X1              =   2280
         X2              =   3240
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line frelines 
         Index           =   11
         X1              =   1200
         X2              =   2160
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line frelines 
         Index           =   10
         X1              =   1200
         X2              =   2160
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line frelines 
         Index           =   9
         X1              =   1200
         X2              =   2160
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line frelines 
         Index           =   8
         X1              =   1200
         X2              =   2160
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line frelines 
         Index           =   7
         X1              =   1200
         X2              =   2160
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line frelines 
         Index           =   6
         X1              =   1200
         X2              =   2160
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line frelines 
         Index           =   5
         X1              =   120
         X2              =   1080
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line frelines 
         Index           =   4
         X1              =   120
         X2              =   1080
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line frelines 
         Index           =   3
         X1              =   120
         X2              =   1080
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line frelines 
         Index           =   2
         X1              =   120
         X2              =   1080
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line frelines 
         Index           =   1
         X1              =   120
         X2              =   1080
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line frelines 
         Index           =   0
         X1              =   120
         X2              =   1080
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   27
         Left            =   3240
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   26
         Left            =   3120
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   25
         Left            =   3000
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   24
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   23
         Left            =   2760
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   22
         Left            =   2640
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   21
         Left            =   2520
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   20
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   19
         Left            =   2280
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   18
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   17
         Left            =   2040
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   16
         Left            =   1920
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   15
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   14
         Left            =   1680
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   13
         Left            =   1560
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   12
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   11
         Left            =   1320
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   10
         Left            =   1200
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   9
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   8
         Left            =   960
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   7
         Left            =   840
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   6
         Left            =   720
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   5
         Left            =   600
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   4
         Left            =   480
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   3
         Left            =   360
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   2
         Left            =   240
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   120
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
      Begin VB.Shape frecircles 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   0
         Left            =   0
         Shape           =   3  'Circle
         Top             =   480
         Width           =   75
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000B&
      Height          =   6615
      Left            =   4440
      Top             =   1920
      Width           =   12735
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   101
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   100
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   99
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   98
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   97
      Left            =   6000
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   96
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   95
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   94
      Left            =   7080
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   93
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   92
      Left            =   7800
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   91
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   90
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   89
      Left            =   8880
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   88
      Left            =   9240
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   87
      Left            =   9600
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   86
      Left            =   9960
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   85
      Left            =   10320
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   84
      Left            =   10680
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   83
      Left            =   11040
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   82
      Left            =   11400
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   81
      Left            =   11760
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   80
      Left            =   12120
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   79
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   78
      Left            =   12840
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   77
      Left            =   13200
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   76
      Left            =   13560
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   75
      Left            =   13920
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   74
      Left            =   14280
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   73
      Left            =   14640
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   72
      Left            =   15000
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   71
      Left            =   15360
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   70
      Left            =   15720
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   69
      Left            =   16080
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   68
      Left            =   16440
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   67
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   66
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   7800
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   65
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   7440
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   64
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   63
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   62
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   61
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   6000
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   60
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   59
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   58
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   57
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   4560
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   56
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   55
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   3840
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   54
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   53
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   52
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   51
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   50
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   49
      Left            =   16440
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   48
      Left            =   16080
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   47
      Left            =   15720
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   46
      Left            =   15360
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   45
      Left            =   15000
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   44
      Left            =   14640
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   43
      Left            =   14280
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   42
      Left            =   13920
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   41
      Left            =   13560
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   40
      Left            =   13200
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   39
      Left            =   12840
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   38
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   37
      Left            =   12120
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   36
      Left            =   11760
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   35
      Left            =   11400
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   34
      Left            =   11040
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   33
      Left            =   10680
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   32
      Left            =   10320
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   31
      Left            =   9960
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   30
      Left            =   9600
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   29
      Left            =   9240
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   28
      Left            =   8880
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   27
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   26
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   25
      Left            =   7800
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   24
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   23
      Left            =   7080
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   22
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   21
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   20
      Left            =   6000
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   19
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   18
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   17
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   16
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   15
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   14
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   13
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   12
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   11
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   3840
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   10
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   9
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   4560
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   8
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   6000
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   7440
      Width           =   255
   End
   Begin VB.Shape blubs 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   7800
      Width           =   255
   End
   Begin VB.Label systext 
      BackStyle       =   0  'Transparent
      Caption         =   "SYSTEM CLOCK FREQUENCY"
      BeginProperty Font 
         Name            =   "Separator Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   9360
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Shape frebox 
      BorderColor     =   &H00C7C7C7&
      FillColor       =   &H00E9E9E9&
      Height          =   1095
      Left            =   120
      Top             =   9600
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "Thanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rndtop(27) As Integer
Dim creatorsname(11) As String
Dim inde As Integer
Dim increleng As Integer
Dim rr2 As Integer
Dim rg2 As Integer
Dim rb2 As Integer
Dim bno As Integer
Dim rr3 As Integer
Dim rg3 As Integer
Dim rb3 As Integer
Dim styleno As Integer
Dim style5b1 As Integer
Dim style5b2 As Integer

Private Sub findtop()
Randomize
    For i = 0 To 27
     rndtop(i) = CInt(Int(960 * Rnd()))
    Next i
fretimer.Enabled = True
End Sub

Private Sub dispcolor_Timer()
Dim rr As Integer
Dim rg As Integer
Dim rb As Integer

Randomize
rr = CInt(Int(255 * Rnd()))
rg = CInt(Int(255 * Rnd()))
rb = CInt(Int(255 * Rnd()))

display.BackColor = RGB(rr, rg, rb)
End Sub

Private Sub disptext_Timer()
On Error GoTo changename


disptext.Enabled = False
textani.Enabled = True

inde = inde + 1
Exit Sub
changename:
    inde = 0
    Resume

End Sub



Private Sub Form_Load()

inde = -1
styleno = 1
style5b1 = 0
style5b2 = 51

creatorsname(0) = "CREATORS OF ****** "
creatorsname(1) = "ARE"
creatorsname(2) = "MIDHUNLAL P H"
creatorsname(3) = "ATHUL KRISHNA E S"
creatorsname(4) = "ABIJITH M A"
creatorsname(5) = "AND"
creatorsname(6) = " "
creatorsname(7) = "AND"
creatorsname(8) = " "
creatorsname(9) = " "
creatorsname(10) = " "
creatorsname(11) = "KAMAL E J"

End Sub

Private Sub lightstyle1_Timer()
On Error GoTo changecolorblub
    blubs(bno).FillColor = RGB(rr2, rg2, rb2)
    bno = bno + 1
Exit Sub
changecolorblub:
 bno = 0
 findcolor
End Sub

Private Sub findcolor()
Randomize
rr2 = CInt(Int(255 * Rnd()))
rg2 = CInt(Int(255 * Rnd()))
rb2 = CInt(Int(255 * Rnd()))
rr3 = CInt(Int(255 * Rnd()))
rg3 = CInt(Int(255 * Rnd()))
rb3 = CInt(Int(255 * Rnd()))
End Sub


Private Sub lightstyle2_Timer()
On Error GoTo changecolorblub
 For bno = 0 To 102
 If bno Mod 2 = 0 Then
  blubs(bno).FillColor = RGB(rr2, rg2, rb2)
 Else
  blubs(bno).FillColor = RGB(rr3, rg3, rb3)
 End If
 Next bno
Exit Sub
changecolorblub:
bno = 0
findcolor
End Sub

Private Sub lightstyle3_Timer()
On Error GoTo changecolorblub
 If bno Mod 5 = 0 Then
    blubs(bno).FillColor = RGB(rr2, rg2, rb2)
    bno = bno + 1
 Else
    blubs(bno).FillColor = RGB(rr3, rg3, rb3)
    bno = bno + 1
 End If
Exit Sub
changecolorblub:
bno = 0
findcolor
End Sub

Private Sub lightstyle4_Timer()
On Error GoTo changecolorblub
 If bno Mod 5 = 0 Then
   findcolor
 End If
    blubs(bno).FillColor = RGB(rr3, rg3, rb3)
    bno = bno + 1
Exit Sub
changecolorblub:
bno = 0
End Sub

Private Sub lightstyle5_Timer()
On Error GoTo changecolorblub
If style5b1 <= 50 Then
blubs(style5b1).FillColor = RGB(rr2, rg2, rb2)
style5b1 = style5b1 + 1
End If
If style5b2 <= 102 Then
blubs(style5b2).FillColor = RGB(rr3, rg3, rb3)
style5b2 = style5b2 + 1
End If
Exit Sub
changecolorblub:
style5b1 = 0
style5b2 = 51
findcolor
End Sub

Private Sub selectstyle_Timer()
Select Case styleno
 Case 1
  lightstyle1.Enabled = True
  lightstyle2.Enabled = False
  lightstyle3.Enabled = False
  lightstyle4.Enabled = False
  lightstyle5.Enabled = False
 Case 2
  lightstyle1.Enabled = False
  lightstyle2.Enabled = True
  lightstyle3.Enabled = False
  lightstyle4.Enabled = False
  lightstyle5.Enabled = False
 Case 3
  lightstyle1.Enabled = False
  lightstyle2.Enabled = False
  lightstyle3.Enabled = True
  lightstyle4.Enabled = False
  lightstyle5.Enabled = False
 Case 4
  lightstyle1.Enabled = False
  lightstyle2.Enabled = False
  lightstyle3.Enabled = False
  lightstyle4.Enabled = True
  lightstyle5.Enabled = False
 Case 5
  lightstyle1.Enabled = False
  lightstyle2.Enabled = False
  lightstyle3.Enabled = False
  lightstyle4.Enabled = False
  lightstyle5.Enabled = True
 Case 6
   styleno = 0
End Select
 styleno = styleno + 1
End Sub

Private Sub tempdisp_Click()
systext.Visible = True
frebox.Visible = True
freframe.Visible = True
tempdisp.Visible = False
lightstyle1.Enabled = True
selectstyle.Enabled = True
dispcolor.Enabled = True
disptext.Enabled = True
fretimerselecter.Enabled = True
End Sub

Private Sub textani_Timer()
Dim templen As Integer
On Error GoTo changename
templen = Len(creatorsname(inde))

If increleng <= templen Then
 disptextlabel.Caption = Mid(creatorsname(inde), 1, increleng)
 increleng = increleng + 1
Else
 disptext.Enabled = True
 textani.Enabled = False
 increleng = 0
End If
Exit Sub
changename:
 inde = 0
End Sub

Private Sub fretimer_Timer()
For i = 0 To 27
    If rndtop(i) > frecircles(i).Top Then
        If rndtop(i) >= frecircles(i).Top Then
            frecircles(i).Top = frecircles(i).Top + 15
            findlines (i)
        Else
            findtop
        End If
    Else
        If rndtop(i) <= frecircles(i).Top Then
            frecircles(i).Top = frecircles(i).Top - 15
            findlines (i)
        Else
            findtop
        End If
    End If
Next i
End Sub

Private Sub findlines(j)
    For j = 0 To 26
     frelines(j).X1 = frecircles(j).Left + 25
     frelines(j).Y1 = frecircles(j).Top + 25
     frelines(j).X2 = frecircles(j + 1).Left + 25
     frelines(j).Y2 = frecircles(j + 1).Top + 25
    Next j
End Sub

Private Sub fretimerselecter_Timer()
findtop
End Sub
