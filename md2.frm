VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5160
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13440
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu sls 
      Caption         =   "Sales"
   End
   Begin VB.Menu ord 
      Caption         =   "Order"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ord_Click()
Unload Me
Orderform.Show
End Sub

Private Sub sls_Click()
Unload Me
Sales.Show
End Sub
