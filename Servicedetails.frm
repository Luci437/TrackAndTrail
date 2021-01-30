VERSION 5.00
Begin VB.Form Servicedetails 
   BackColor       =   &H00680235&
   Caption         =   "Service Room"
   ClientHeight    =   10050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19125
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10050
   ScaleWidth      =   19125
   WindowState     =   2  'Maximized
   Begin VB.CommandButton addtolistbtn 
      Caption         =   "Add"
      Height          =   375
      Left            =   19200
      TabIndex        =   19
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox servicedata 
      BackColor       =   &H00430122&
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
      ForeColor       =   &H00332FEA&
      Height          =   375
      Index           =   3
      Left            =   17760
      TabIndex        =   18
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox servicedata 
      BackColor       =   &H00430122&
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
      ForeColor       =   &H00332FEA&
      Height          =   375
      Index           =   2
      Left            =   7560
      TabIndex        =   17
      Top             =   2760
      Width           =   10095
   End
   Begin VB.ListBox Detailamt 
      Appearance      =   0  'Flat
      BackColor       =   &H00430122&
      Columns         =   3
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   6330
      ItemData        =   "Servicedetails.frx":0000
      Left            =   17760
      List            =   "Servicedetails.frx":0002
      TabIndex        =   16
      Top             =   3240
      Width           =   2055
   End
   Begin VB.ListBox Detaillist 
      Appearance      =   0  'Flat
      BackColor       =   &H00430122&
      Columns         =   3
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   6330
      ItemData        =   "Servicedetails.frx":0004
      Left            =   7560
      List            =   "Servicedetails.frx":0006
      TabIndex        =   15
      Top             =   3240
      Width           =   10095
   End
   Begin VB.ComboBox empnamecombo 
      BackColor       =   &H00430122&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00332FEA&
      Height          =   450
      Left            =   3600
      TabIndex        =   13
      Top             =   6600
      Width           =   3615
   End
   Begin VB.ComboBox prdnamecombo 
      BackColor       =   &H00430122&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00332FEA&
      Height          =   450
      Left            =   3600
      TabIndex        =   11
      Top             =   5760
      Width           =   3615
   End
   Begin VB.ComboBox custidcombo 
      BackColor       =   &H00430122&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00332FEA&
      Height          =   450
      Left            =   3600
      TabIndex        =   8
      Top             =   2760
      Width           =   3615
   End
   Begin VB.ComboBox sridcombo 
      BackColor       =   &H00430122&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00332FEA&
      Height          =   450
      Left            =   3600
      TabIndex        =   7
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox servicedata 
      BackColor       =   &H00430122&
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
      ForeColor       =   &H00332FEA&
      Height          =   495
      Index           =   1
      Left            =   3600
      TabIndex        =   5
      Top             =   4440
      Width           =   3615
   End
   Begin VB.TextBox servicedata 
      BackColor       =   &H00430122&
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
      ForeColor       =   &H00332FEA&
      Height          =   495
      Index           =   0
      Left            =   3600
      TabIndex        =   3
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Label notfillerror 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   5
      Left            =   7200
      TabIndex        =   34
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label notfillerror 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   4
      Left            =   7200
      TabIndex        =   33
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label notfillerror 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   3
      Left            =   7200
      TabIndex        =   32
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label notfillerror 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   17400
      TabIndex        =   31
      Top             =   2700
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label notfillerror 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   7200
      TabIndex        =   30
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label notfillerror 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   7200
      TabIndex        =   29
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label delbtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   240
      TabIndex        =   28
      Top             =   10020
      Width           =   6975
   End
   Begin VB.Label updatebtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   240
      TabIndex        =   27
      Top             =   9200
      Width           =   6975
   End
   Begin VB.Label savebtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   240
      TabIndex        =   26
      Top             =   8320
      Width           =   6975
   End
   Begin VB.Label newbtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   240
      TabIndex        =   25
      Top             =   7510
      Width           =   6975
   End
   Begin VB.Shape Shape8 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00380C21&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   240
      Top             =   9960
      Width           =   6975
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00380C21&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   240
      Top             =   9120
      Width           =   6975
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00380C21&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   240
      Top             =   8280
      Width           =   6975
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00380C21&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   240
      Top             =   7440
      Width           =   6975
   End
   Begin VB.Label amttopay 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00430DE3&
      Height          =   495
      Left            =   17760
      TabIndex        =   24
      Top             =   9600
      Width           =   2655
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00620232&
      FillStyle       =   0  'Solid
      Height          =   3495
      Left            =   120
      Top             =   7320
      Width           =   7215
   End
   Begin VB.Label specificclearbtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   9120
      TabIndex        =   23
      Top             =   9795
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006B6BFE&
      Height          =   255
      Left            =   17760
      TabIndex        =   22
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Service details"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006B6BFE&
      Height          =   375
      Left            =   7560
      TabIndex        =   21
      Top             =   2400
      Width           =   8535
   End
   Begin VB.Label clearallbtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLEAR ALL"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   7560
      TabIndex        =   20
      Top             =   9800
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00300304&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   7560
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Label amttopaylabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Charge"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00430DE3&
      Height          =   495
      Left            =   14880
      TabIndex        =   14
      Top             =   9600
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E4C054&
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   6600
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E4C054&
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Phone"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E4C054&
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E4C054&
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E4C054&
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service ID"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E4C054&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label titleaboutlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "In this form Service details are added"
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
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Titlelabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Details"
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C000&
      Index           =   1
      X1              =   360
      X2              =   5760
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      Index           =   2
      X1              =   360
      X2              =   5760
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00620232&
      FillStyle       =   0  'Solid
      Height          =   2655
      Left            =   240
      Top             =   2520
      Width           =   7215
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00300304&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   9120
      Top             =   9720
      Width           =   1455
   End
End
Attribute VB_Name = "Servicedetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim status As Integer

Private Sub addtolistbtn_Click()
Detaillist.AddItem (servicedata(2).Text)
Detailamt.AddItem (servicedata(3).Text)
adddetails
servicedata(2).Text = ""
servicedata(3).Text = ""
End Sub

Private Sub adddetails()
Dim total As Long
For i = 0 To Detailamt.ListCount
 total = total + Val(Detailamt.List(i))
Next i
amttopay.Caption = total
End Sub

Private Sub clearallbtn_Click()
 Detailamt.Clear
 Detaillist.Clear
 amttopay.Caption = ""
End Sub

Private Sub custidcombo_Click()
If rs.State = 1 Then rs.Close
 rs.Open "select Cr_name,Cr_ph from Customer_details where Cr_id = '" & custidcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
 servicedata(0).Text = rs.Fields(0)
 servicedata(1).Text = rs.Fields(1)
 rs.Close
End Sub

Private Sub delbtn_Click()
If rs.State = 1 Then rs.Close
 rs.Open "select * from SD where Sr_id = '" & sridcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
 rs.Delete
 rs.Close
 rs.Open "select * from Service_table where Sr_id = '" & sridcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
 rs.Delete
 MsgBox "Data Deleted"
 rs.Close
 clearallfields
 autosridcomboload
 autosrid
End Sub

Private Sub Form_Load()
 connect
 autosridcomboload
 autosrid
 autocustid
 autoprdnameload
 autoempnameload
 status = 0
End Sub

Private Sub autosridcomboload()
If rs.State = 1 Then rs.Close
 rs.Open "select Sr_id from Service_table", con, adOpenKeyset, adLockOptimistic
 If rs.RecordCount <> 0 Then
 rs.MoveFirst
  While rs.EOF = False
   sridcombo.AddItem (rs.Fields(0))
   rs.MoveNext
  Wend
 rs.Close
End If
End Sub

Private Sub autoempnameload()
 If rs.State = 1 Then rs.Close
  rs.Open "select Emp_name from Employee_Table", con, adOpenKeyset, adLockOptimistic
  rs.MoveFirst
   While rs.EOF = False
    empnamecombo.AddItem (rs.Fields(0))
    rs.MoveNext
   Wend
  rs.Close
End Sub

Private Sub autoprdnameload()
 If rs.State = 1 Then rs.Close
  rs.Open "select Pr_name from Stock_table", con, adOpenKeyset, adLockOptimistic
   rs.MoveFirst
   While rs.EOF = False
    prdnamecombo.AddItem (rs.Fields(0))
    rs.MoveNext
   Wend
  rs.Close
End Sub

Private Sub autocustid()
If rs.State = 1 Then rs.Close
 rs.Open "select Cr_id from Customer_details", con, adOpenKeyset, adLockOptimistic
  While rs.EOF = False
   custidcombo.AddItem (rs.Fields(0))
   rs.MoveNext
  Wend
rs.Close
End Sub
    
Private Sub autosrid()
 Dim l As Integer, id As Integer
 If rs.State = 1 Then rs.Close
  rs.Open "select * from Service_table", con, adOpenKeyset, adLockOptimistic
  If rs.RecordCount = 0 Then
   sridcombo.Text = "Sr-1"
  Else
   l = 0
   While rs.EOF = False
    id = Val(Mid(rs.Fields(0), 4))
    If l < id Then
     l = id
    End If
    rs.MoveNext
   Wend
   l = l + 1
   sridcombo.Text = "Sr-" & l
  End If
 rs.Close
End Sub

Private Sub newbtn_Click()
clearallfields
autosrid
End Sub

Private Sub savebtn_Click()
checkallfields
If status = 0 Then
'-----------------------------------
 If rs.State = 1 Then rs.Close
  rs.Open "select * from Service_table", con, adOpenKeyset, adLockOptimistic
  With rs
   .AddNew
   .Fields(0) = sridcombo.Text
   .Fields(1) = custidcombo.Text
   .Fields(2) = prdnamecombo.Text
   .Fields(3) = empnamecombo.Text
   .Fields(4) = Val(amttopay.Caption)
   .Update
  End With
 rs.Close
  For i = 0 To Detaillist.ListCount - 1
   rs.Open "select * from SD", con, adOpenKeyset, adLockOptimistic
   With rs
   .AddNew
   rs.Fields(0) = sridcombo.Text
   rs.Fields(1) = Detaillist.List(i)
   rs.Fields(2) = Val(Detailamt.List(i))
   .Update
   End With
    rs.Close
  Next i
 MsgBox "Data Saved"
'-----------------------------------
autosridcomboload
End If 'status
End Sub

Private Sub specificclearbtn_Click()
 Dim one As Integer
 If Detaillist.ListIndex >= 0 Then
 one = Detaillist.ListIndex
 Detaillist.RemoveItem (one)
 Detailamt.RemoveItem (one)
 adddetails
 Else
  MsgBox "Please Select any ", vbCritical, "Option Error"
 End If
End Sub

Private Sub clearallfields()
sridcombo.Text = ""
custidcombo.Text = ""
servicedata(0).Text = ""
servicedata(1).Text = ""
prdnamecombo.Text = ""
empnamecombo.Text = ""
servicedata(2).Text = ""
servicedata(3).Text = ""
clearallbtn_Click
End Sub

Private Sub checkallfields()
status = 0
For i = 0 To 1
 If servicedata(i).Text = "" Then
  notfillerror(i).Visible = True
  status = 1
 Else
  notfillerror(i).Visible = False
 End If
Next i
 If custidcombo.Text = "" Then
 notfillerror(3).Visible = True
  status = 1
 Else
  notfillerror(3).Visible = False
 End If
 If prdnamecombo.Text = "" Then
  notfillerror(4).Visible = True
  status = 1
 Else
  notfillerror(4).Visible = False
 End If
 If empnamecombo.Text = "" Then
  notfillerror(5).Visible = True
  status = 1
 Else
  notfillerror(5).Visible = False
 End If
End Sub

Private Sub sridcombo_Click()
Detaillist.Clear
Detailamt.Clear
If rs.State = 1 Then rs.Close
 rs.Open "select * from Service_table where Sr_id = '" & sridcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
 With rs
  custidcombo.Text = .Fields(1)
  prdnamecombo.Text = .Fields(2)
  empnamecombo.Text = .Fields(3)
  amttopay.Caption = .Fields(4)
 End With
rs.Close

rs.Open "select Cr_name,Cr_ph from Customer_details where Cr_id = '" & custidcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
 servicedata(0).Text = rs.Fields(0)
 servicedata(1).Text = rs.Fields(1)
rs.Close

 rs.Open "select * from SD where Sr_id = '" & sridcombo.Text & "'", con, adOpenDynamic, adLockOptimistic
 While rs.EOF = False
  Detaillist.AddItem (rs.Fields(1))
  Detailamt.AddItem (rs.Fields(1))
  rs.MoveNext
 Wend
 rs.Close
End Sub

Private Sub updatebtn_Click()
If rs.State = 1 Then rs.Close
 rs.Open "select * from Service_table where Sr_id = '" & sridcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
 rs.Fields(1) = custidcombo.Text
 rs.Fields(2) = prdnamecombo.Text
 rs.Fields(3) = empnamecombo.Text
 rs.Fields(4) = Val(amttopay.Caption)
 rs.Update
rs.Close
For i = 0 To Detaillist.ListCount - 1
 rs.Open "select * from SD where Sr_id ='" & sridcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
  With rs
   .Fields(1) = Detaillist.List(i)
   .Fields(2) = Detailamt.List(i)
   .Update
  End With
 rs.Close
Next i
End Sub
