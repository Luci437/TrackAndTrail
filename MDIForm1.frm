VERSION 5.00
Begin VB.MDIForm NewMd 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   9630
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   17625
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   11295
      Left            =   0
      ScaleHeight     =   11235
      ScaleWidth      =   17565
      TabIndex        =   0
      Top             =   0
      Width           =   17625
      Begin VB.Frame optionframe 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   5
         Left            =   17760
         TabIndex        =   8
         Top             =   2160
         Width           =   2535
         Begin VB.Label servicelabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SERVICES"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   14
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.Frame optionframe 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   4
         Left            =   15120
         TabIndex        =   7
         Top             =   2160
         Width           =   2535
         Begin VB.Label saleslabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SALES"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   13
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.Frame optionframe 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   3
         Left            =   12480
         TabIndex        =   6
         Top             =   2160
         Width           =   2535
         Begin VB.Label customerlabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "CUSTOMER"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   12
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.Frame optionframe 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   2
         Left            =   9840
         TabIndex        =   5
         Top             =   2160
         Width           =   2535
         Begin VB.Label orderlabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ORDER"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   11
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.Frame optionframe 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   1
         Left            =   7200
         TabIndex        =   4
         Top             =   2160
         Width           =   2535
         Begin VB.Line Line2 
            BorderColor     =   &H00919191&
            X1              =   0
            X2              =   2520
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00919191&
            X1              =   0
            X2              =   2520
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label viewstockbtn 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VIEW STOCKS"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   375
            Left            =   0
            TabIndex        =   16
            Top             =   1320
            Width           =   2535
         End
         Begin VB.Label addstockbtn 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ADD TO STOCK"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   375
            Left            =   0
            TabIndex        =   15
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label stocklabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "STOCKS"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   10
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.Frame optionframe 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   0
         Left            =   4560
         TabIndex        =   3
         Top             =   2160
         Width           =   2535
         Begin VB.Label emplabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EMPLOYEE"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   9
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.Timer slidertimer 
         Interval        =   5000
         Left            =   120
         Top             =   3240
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   855
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   20535
         Begin VB.Image Image1 
            Height          =   855
            Left            =   16560
            Picture         =   "MDIForm1.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   3855
         End
      End
      Begin VB.Label title1 
         BackStyle       =   0  'Transparent
         Caption         =   "Track && Trail"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   4455
      End
      Begin VB.Image slideshow 
         Height          =   8175
         Left            =   0
         Picture         =   "MDIForm1.frx":6D48
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   20415
      End
   End
   Begin VB.Menu sls 
      Caption         =   "Sales"
   End
End
Attribute VB_Name = "NewMd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim imgcnter As Integer

Private Sub addstockbtn_Click()
Unload Me
StockEntry.Show
End Sub

Private Sub customerlabel_Click()
findoption (3)
Unload Me
Customer.Show
End Sub

Private Sub emplabel_Click()
findoption (0)
Unload Me
Employee_form.Show
End Sub

Private Sub findoption(op As Integer)
 For i = 0 To 5
  If i = op Then
   optionframe(i).BackColor = &H378AD7
  Else
   optionframe(i).BackColor = &H8000000F
  End If
 Next i
 If op = 1 Then
 optionframe(1).Height = 1815
 Else
 optionframe(1).Height = 495
 End If

End Sub

Private Sub MDIForm_Load()
imgcnter = 0
End Sub

Private Sub orderlabel_Click()
findoption (2)
Orderform.Show
End Sub

Private Sub saleslabel_Click()
findoption (4)
Unload Me
Sales.Show
End Sub

Private Sub servicelabel_Click()
findoption (5)
Unload Me
Servicedetails.Show
End Sub

Private Sub slidertimer_Timer()
On Error GoTo uturn
slideshow.Picture = LoadPicture("D:\New\MdiImages\img" & imgcnter & ".jpg")
imgcnter = imgcnter + 1
Exit Sub
uturn:
 imgcnter = 0
 Resume
End Sub

Private Sub sls_Click()
Unload Me
Sales.Show
End Sub

Private Sub stocklabel_Click()
findoption (1)
End Sub

Private Sub viewstockbtn_Click()
Unload Me
prddisp.Show
End Sub
