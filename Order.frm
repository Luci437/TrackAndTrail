VERSION 5.00
Begin VB.Form Orderform 
   BackColor       =   &H00271203&
   Caption         =   "Order List"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19050
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   19050
   WindowState     =   2  'Maximized
   Begin VB.ListBox ordnolist 
      Appearance      =   0  'Flat
      BackColor       =   &H00110100&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   8445
      ItemData        =   "Order.frx":0000
      Left            =   720
      List            =   "Order.frx":0002
      TabIndex        =   17
      Top             =   2040
      Width           =   975
   End
   Begin VB.ComboBox qntcombo 
      Appearance      =   0  'Flat
      BackColor       =   &H00271203&
      ForeColor       =   &H000000FF&
      Height          =   315
      ItemData        =   "Order.frx":0004
      Left            =   10680
      List            =   "Order.frx":0029
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox statuscombo 
      Appearance      =   0  'Flat
      BackColor       =   &H00271203&
      ForeColor       =   &H000000FF&
      Height          =   315
      ItemData        =   "Order.frx":0050
      Left            =   11880
      List            =   "Order.frx":005A
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.ListBox prdstatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00110100&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   8445
      ItemData        =   "Order.frx":0078
      Left            =   11880
      List            =   "Order.frx":007A
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ListBox prdqntlist 
      Appearance      =   0  'Flat
      BackColor       =   &H00110100&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   8445
      ItemData        =   "Order.frx":007C
      Left            =   10680
      List            =   "Order.frx":007E
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ListBox prdnamelist 
      Appearance      =   0  'Flat
      BackColor       =   &H00110100&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   8445
      ItemData        =   "Order.frx":0080
      Left            =   2640
      List            =   "Order.frx":0082
      TabIndex        =   2
      Top             =   2040
      Width           =   8055
   End
   Begin VB.ListBox prdlist 
      Appearance      =   0  'Flat
      BackColor       =   &H00110100&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   8445
      ItemData        =   "Order.frx":0084
      Left            =   1680
      List            =   "Order.frx":0086
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin VB.ListBox srlno 
      Appearance      =   0  'Flat
      BackColor       =   &H00110100&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   8445
      ItemData        =   "Order.frx":0088
      Left            =   120
      List            =   "Order.frx":008A
      TabIndex        =   0
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label shbtn 
      BackStyle       =   0  'Transparent
      Caption         =   "SHOW DATA"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   14760
      TabIndex        =   20
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label updbtn 
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE "
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   14760
      TabIndex        =   19
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Or No"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3F3F3&
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label clearallbtn 
      BackStyle       =   0  'Transparent
      Caption         =   "CLEAR ALL"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   14760
      TabIndex        =   16
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label clrbtn 
      BackStyle       =   0  'Transparent
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   14760
      TabIndex        =   15
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label savebtn 
      BackStyle       =   0  'Transparent
      Caption         =   "SAVE DATA"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   14760
      TabIndex        =   14
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14040
      Top             =   4920
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14640
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14040
      Top             =   4200
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14640
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14040
      Top             =   2760
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14640
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Items which reached reload limits are shown here."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3F3F3&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3F3F3&
      Height          =   255
      Left            =   11880
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3F3F3&
      Height          =   255
      Left            =   10680
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Product  Name"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3F3F3&
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Product  ID"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3F3F3&
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sr No"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3F3F3&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDER LIST"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3F3F3&
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3495
   End
   Begin VB.Shape Shape8 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14640
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14040
      Top             =   3480
      Width           =   615
   End
   Begin VB.Shape Shape10 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14640
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14040
      Top             =   2040
      Width           =   615
   End
End
Attribute VB_Name = "Orderform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim orid As String

Private Sub clrbtn_Click()
If rs.State = 1 Then rs.Close
 rs.Open "select * from Order_table", con, adOpenKeyset, adLockOptimistic
  rs.Delete
rs.Close
End Sub

Private Sub Form_Load()
connect
End Sub

Private Sub shbtn_Click()
     findorid
 If rs.State = 1 Then rs.Close
  rs.Open "select * from Stock_table s where s.Pr_stock <= s.Pr_rol", con, adOpenKeyset, adLockOptimistic
  While rs.EOF = False
     ordnolist.AddItem (orid)
     prdlist.AddItem (rs.Fields(0))
     prdnamelist.AddItem (rs.Fields(1))
     prdqntlist.AddItem (rs.Fields(5))
     prdstatus.AddItem ("NOT PURCHASED")
     rs.MoveNext
  Wend
End Sub

Private Sub findorid()
Dim l As Integer, id As Integer
 If rs.State = 1 Then rs.Close
  rs.Open "select * from Order_table", con, adOpenKeyset, adLockOptimistic
  If rs.EOF = True Then
   orid = "Or-1"
  Else
   l = 0
   While rs.EOF = False
    id = Val(Mid(rs.Fields(0), 4))
    If id > l Then
     l = id
    End If
    rs.MoveNext
   Wend
   orid = "Or-" & l + 1
  End If
End Sub
