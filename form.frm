VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8070
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10890
   LinkTopic       =   "Form2"
   ScaleHeight     =   8070
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   8055
      Left            =   0
      Picture         =   "form.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10905
   End
   Begin VB.Menu mnuCustomer 
      Caption         =   "Customer"
      Begin VB.Menu mnuAddCustomer 
         Caption         =   "AddCustomer"
      End
      Begin VB.Menu mnuEditCustomer 
         Caption         =   "EditCustomer"
      End
   End
   Begin VB.Menu mnuGasAgency 
      Caption         =   "GasAgency"
      Begin VB.Menu mnuSetPrice 
         Caption         =   "SetPrice"
      End
   End
   Begin VB.Menu mnuStocks 
      Caption         =   "Stocks"
      Begin VB.Menu mnuStocskRecord 
         Caption         =   "StocskRecord"
      End
      Begin VB.Menu mnuCylinderBooking 
         Caption         =   "CylinderBooking"
      End
      Begin VB.Menu mnuCylinderDelivery 
         Caption         =   "CylinderDelivery"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Begin VB.Menu mnuStockRecord 
         Caption         =   "StockRecord"
      End
      Begin VB.Menu mnuCustomerReport 
         Caption         =   "CustomerReport"
      End
      Begin VB.Menu mnuDeliveryReport 
         Caption         =   "DeliveryReport"
      End
      Begin VB.Menu mnuRequestRecord 
         Caption         =   "RequestRecord"
      End
   End
   Begin VB.Menu mnuAboutUs 
      Caption         =   "AboutUs"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form1.Hide
End Sub

Private Sub mnuChangePassword_Click()
Form3.Show
Form2.Hide
End Sub


Private Sub mnuAboutUs_Click()

End Sub

Private Sub mnuAddCustomer_Click()
Form3.Show
Form1.Hide
End Sub


Private Sub mnuEditCustomer_Click()
Form4.Show
Form2.Hide
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuLogout_Click()
Form1.Show
Form2.Hide
End Sub

Private Sub mnuStock_Click()

End Sub

Private Sub mnuSetPrice_Click()
Form5.Show
End Sub
