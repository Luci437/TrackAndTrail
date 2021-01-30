VERSION 5.00
Begin VB.Form Customer 
   BackColor       =   &H001F1403&
   Caption         =   "Customer Details"
   ClientHeight    =   9705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   9705
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox custdetailtextbx 
      BackColor       =   &H000F0A02&
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
      ForeColor       =   &H8000000D&
      Height          =   495
      Index           =   4
      Left            =   11040
      TabIndex        =   14
      Top             =   9000
      Width           =   2895
   End
   Begin VB.TextBox custdetailtextbx 
      BackColor       =   &H000F0A02&
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
      ForeColor       =   &H8000000D&
      Height          =   495
      Index           =   3
      Left            =   11040
      TabIndex        =   13
      Top             =   8160
      Width           =   2895
   End
   Begin VB.TextBox custdetailtextbx 
      BackColor       =   &H000F0A02&
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
      ForeColor       =   &H8000000D&
      Height          =   495
      Index           =   2
      Left            =   11040
      MaxLength       =   10
      TabIndex        =   10
      Top             =   9840
      Width           =   2895
   End
   Begin VB.TextBox custdetailtextbx 
      BackColor       =   &H000F0A02&
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
      ForeColor       =   &H8000000D&
      Height          =   495
      Index           =   1
      Left            =   4080
      TabIndex        =   8
      Top             =   9840
      Width           =   2895
   End
   Begin VB.TextBox custdetailtextbx 
      BackColor       =   &H000F0A02&
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
      ForeColor       =   &H8000000D&
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   6
      Top             =   9000
      Width           =   2895
   End
   Begin VB.ComboBox ctidcombo 
      Appearance      =   0  'Flat
      BackColor       =   &H000F0A02&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   4080
      TabIndex        =   4
      Top             =   8160
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000C0C0C&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5895
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   19935
      Begin VB.Timer imageselecter 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   600
         Top             =   5400
      End
      Begin VB.Timer imagemove 
         Interval        =   6000
         Left            =   120
         Top             =   5400
      End
      Begin VB.Image slideshow 
         Height          =   5655
         Left            =   120
         Picture         =   "Customer.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   19695
      End
      Begin VB.Label quoteslabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   5160
         Width           =   19815
      End
   End
   Begin VB.Timer title2 
      Enabled         =   0   'False
      Interval        =   24
      Left            =   4920
      Top             =   600
   End
   Begin VB.Timer titleanima 
      Interval        =   24
      Left            =   4200
      Top             =   360
   End
   Begin VB.Label fielderrorlabel 
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
      Left            =   14040
      TabIndex        =   26
      Top             =   9000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label fielderrorlabel 
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
      Left            =   14040
      TabIndex        =   25
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label fielderrorlabel 
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
      Left            =   14040
      TabIndex        =   24
      Top             =   9840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label fielderrorlabel 
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
      Left            =   7080
      TabIndex        =   23
      Top             =   9840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label fielderrorlabel 
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
      Left            =   7080
      TabIndex        =   22
      Top             =   9000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label numbererror 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Only Numeric values"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11040
      TabIndex        =   21
      Top             =   10320
      Visible         =   0   'False
      Width           =   2895
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
      ForeColor       =   &H00FDD1AE&
      Height          =   375
      Left            =   16080
      TabIndex        =   19
      Top             =   10200
      Width           =   3975
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
      ForeColor       =   &H00FDD1AE&
      Height          =   375
      Left            =   16080
      TabIndex        =   18
      Top             =   9360
      Width           =   3975
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
      ForeColor       =   &H00FDD1AE&
      Height          =   375
      Left            =   16080
      TabIndex        =   17
      Top             =   8520
      Width           =   3975
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
      ForeColor       =   &H00FDD1AE&
      Height          =   375
      Left            =   16080
      TabIndex        =   16
      Top             =   7680
      Width           =   3975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0050CB05&
      Height          =   255
      Left            =   7680
      TabIndex        =   15
      Top             =   7800
      Width           =   3375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000F0A02&
      X1              =   14520
      X2              =   14520
      Y1              =   7320
      Y2              =   10575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Place Name"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7680
      TabIndex        =   12
      Top             =   9000
      Width           =   3375
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7680
      TabIndex        =   11
      Top             =   8160
      Width           =   3375
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000F0A02&
      X1              =   7560
      X2              =   7560
      Y1              =   7320
      Y2              =   10575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Number"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7680
      TabIndex        =   9
      Top             =   9840
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Email ID"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   9840
      Width           =   3375
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   9000
      Width           =   3015
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   8160
      Width           =   3015
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000F0A02&
      X1              =   240
      X2              =   20040
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00702B01&
      Index           =   2
      X1              =   240
      X2              =   5640
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00CA4E02&
      Index           =   1
      X1              =   240
      X2              =   5640
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Titlelabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Details"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label titleaboutlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "In this form Admin Stores informations about customers"
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
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Shape delbtnshape 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00B64A10&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   15960
      Top             =   10080
      Width           =   4215
   End
   Begin VB.Shape savebtnshape 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00B64A10&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   15960
      Top             =   8400
      Width           =   4215
   End
   Begin VB.Shape updatebtnshape 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00B64A10&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   15960
      Top             =   9240
      Width           =   4215
   End
   Begin VB.Shape newbtnshape 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00B64A10&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   15960
      Top             =   7560
      Width           =   4215
   End
End
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim titlelength As Integer
Dim titleval As String
Dim titleincre As Integer
Dim incre As Integer
Dim imageTop As Integer
Dim imagecounter As Integer
Dim totalimages As Integer
Dim totalquotes As Integer
Dim quotes
Dim fieldstatus As Integer

Private Sub customeridload()
ctidcombo.Clear
If rs.State = 1 Then rs.Close
 rs.Open "select * from Customer_details", con, adOpenKeyset, adLockOptimistic
 rs.MoveFirst
 While rs.EOF = False
  ctidcombo.AddItem (rs.Fields(0))
  rs.MoveNext
 Wend
rs.Close
autoctid
End Sub

Private Sub rehidex()
 For i = 0 To 4
  fielderrorlabel(i).Visible = False
 Next i
 numbererror.Visible = False
End Sub

Private Sub ctidcombo_Click()
If rs.State = 1 Then rs.Close
 rs.Open "select * from Customer_details where Cr_id = '" & ctidcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
  ctidcombo.Text = rs.Fields(0)
  custdetailtextbx(0).Text = rs.Fields(1)
  custdetailtextbx(3).Text = rs.Fields(2)
  custdetailtextbx(4).Text = rs.Fields(3)
  custdetailtextbx(1).Text = rs.Fields(4)
  custdetailtextbx(2).Text = rs.Fields(5)
 rs.Close
End Sub

Private Sub custdetailtextbx_Change(Index As Integer)
If Index = 2 Then
 If IsNumeric(custdetailtextbx(2).Text) = False Then
  numbererror.Visible = True
  fieldstatus = 1
 Else
  numbererror.Visible = False
 End If
End If
End Sub


Private Sub Form_Load()
connect
autoctid
customeridload
titlelength = Len(Titlelabel.Caption)
titleval = Titlelabel.Caption
Titlelabel.Caption = ""
titleincre = 0
'---------------
incre = 2
imageTop = 1
imagecounter = 0
totalimages = 5
totalquotes = 3
quotes = Array("Always do your best. What you plant now, you will harvest later. -Og Mandino", "Don't watch the clock; do what it does. Keep going. -Sam Levenson", "The secret of getting ahead is getting started. -Mark Twain", "Act as if what you do makes a difference. It does. -William James")
End Sub

Private Sub autoctid()
Dim l As Integer, id As Integer
 If rs.State = 1 Then rs.Close
  rs.Open "select * from Customer_details", con, adOpenKeyset, adLockOptimistic
  If rs.RecordCount <> 0 Then
  rs.MoveFirst
  While rs.EOF = False
   l = 0
   id = Val(Mid(rs.Fields(0), 4))
   If id > l Then
   l = id
   End If
   rs.MoveNext
  Wend
  l = l + 1
  ctidcombo.Text = "CT-" & l
 Else
  ctidcombo.Text = "CT-1"
 End If
 rs.Close
End Sub

Private Sub imagemove_Timer()
Dim rndno As Integer
Randomize
rndno = CInt(Int(totalquotes * Rnd()))
imageselecter.Enabled = True
quoteslabel.Caption = quotes(rndno)
End Sub

Private Sub imageselecter_Timer()
If imageTop = 1 Then
 If slideshow.Top >= -5700 Then
  slideshow.Top = slideshow.Top - 100
 Else
  imageTop = 0
  
  If imagecounter <= totalimages Then
   slideshow.Picture = LoadPicture("F:\Project46\MiniProject\Images\Sales\cust" & imagecounter & ".jpg")
   imagecounter = imagecounter + 1
 Else
  imagecounter = 0
  End If
  
 End If
Else
 If slideshow.Top < 100 Then
  slideshow.Top = slideshow.Top + 100
 Else
  imageselecter.Enabled = False
  imageTop = 1
 End If
End If
End Sub

'BUTTON PRESS COLOR CHANGE CODE
Private Sub newbtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
newbtnshape.FillColor = &H632601
newbtn.ForeColor = &H8000000B
clearallfields
customeridload
rehidex
End Sub

Private Sub newbtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
newbtnshape.FillColor = &HB64A10
newbtn.ForeColor = &HFDD1AE
End Sub

Private Sub Savebtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
savebtnshape.FillColor = &H632601
savebtn.ForeColor = &H8000000B
checkallfields
'----------------------SAVE DETAILS COMES UNDER THIS
 If fieldstatus = 0 Then
  If rs.State = 1 Then rs.Close
   rs.Open "select * from Customer_details", con, adOpenKeyset, adLockOptimistic
   rs.AddNew
   rs.Fields(0) = ctidcombo.Text
   rs.Fields(1) = custdetailtextbx(0).Text
   rs.Fields(2) = custdetailtextbx(3).Text
   rs.Fields(3) = custdetailtextbx(4).Text
   rs.Fields(4) = custdetailtextbx(1).Text
   rs.Fields(5) = custdetailtextbx(2).Text
   rs.Update
   MsgBox "Customer Details Added"
  rs.Close
  
 End If
'---------------------------------------------------
customeridload
End Sub

Private Sub Savebtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
savebtnshape.FillColor = &HB64A10
savebtn.ForeColor = &HFDD1AE
End Sub

Private Sub updatebtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
updatebtnshape.FillColor = &H632601
updatebtn.ForeColor = &H8000000B
End Sub

Private Sub updatebtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
updatebtnshape.FillColor = &HB64A10
updatebtn.ForeColor = &HFDD1AE
If rs.State = 1 Then rs.Close
 rs.Open "select * from Customer_details where Cr_id = '" & ctidcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
  rs.Fields(1) = custdetailtextbx(0).Text
  rs.Fields(2) = custdetailtextbx(3).Text
  rs.Fields(3) = custdetailtextbx(4).Text
  rs.Fields(4) = custdetailtextbx(1).Text
  rs.Fields(5) = custdetailtextbx(2).Text
  rs.Update
  MsgBox "Record Updated"
rs.Close
End Sub

Private Sub delbtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
delbtnshape.FillColor = &H632601
delbtn.ForeColor = &H8000000B
End Sub

Private Sub delbtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
delbtnshape.FillColor = &HB64A10
delbtn.ForeColor = &HFDD1AE
If rs.State = 1 Then rs.Close
 rs.Open "select * from Customer_details where Cr_id = '" & ctidcombo.Text & "'", con, adOpenKeyset, adLockOptimistic
 rs.Delete
 MsgBox "Customer Details deleted"
 rs.Close
 clearallfields
End Sub

'---------------------------------------------------------------------------------------------------------

Private Sub checkallfields()
fieldstatus = 0
For i = 0 To 4
 If custdetailtextbx(i).Text = "" Then
  fieldstatus = 1
  fielderrorlabel(i).Visible = True
 Else
  fielderrorlabel(i).Visible = False
 End If
Next i
 If IsNumeric(custdetailtextbx(2).Text) = False Then
  fieldstatus = 1
End If
End Sub

Private Sub clearallfields()
 For i = 0 To 4
  custdetailtextbx(i).Text = ""
 Next i
End Sub


Private Sub titleanima_Timer()
If titleincre <= titlelength Then
Titlelabel.Caption = Mid(titleval, 1, titleincre)
titleincre = titleincre + 1
Else
titleanima.Enabled = False
titlelength = Len(titleaboutlabel.Caption)
titleval = titleaboutlabel.Caption
titleaboutlabel.Caption = ""
titleincre = 1
titleaboutlabel.Visible = True
title2.Enabled = True
End If
End Sub

Private Sub title2_Timer()
If titleincre <= titlelength Then
titleaboutlabel.Caption = Mid(titleval, 1, titleincre)
titleincre = titleincre + 1
Else
title2.Enabled = False
End If
End Sub

