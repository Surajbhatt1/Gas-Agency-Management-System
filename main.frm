VERSION 5.00
Begin VB.Form main 
   Caption         =   "Form2"
   ClientHeight    =   8070
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   ScaleHeight     =   8070
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   8055
      Left            =   0
      Picture         =   "main.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11985
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
   Begin VB.Menu mnuBookig 
      Caption         =   "Booking"
      Begin VB.Menu mnuBookRequest 
         Caption         =   "BookRequest"
      End
      Begin VB.Menu mnuOtherRequest 
         Caption         =   "OtherRequest"
      End
   End
   Begin VB.Menu mnuStocks 
      Caption         =   "Stocks"
      Begin VB.Menu mnuStocskRecord 
         Caption         =   "StocskRecord"
      End
      Begin VB.Menu mnuCylinderDelivery 
         Caption         =   "CylinderDelivery"
      End
   End
   Begin VB.Menu a 
      Caption         =   "Search"
      Begin VB.Menu as 
         Caption         =   "Search Customer"
      End
      Begin VB.Menu fdsv 
         Caption         =   "Search Booking"
      End
      Begin VB.Menu uhj 
         Caption         =   "Search Other Bookings"
      End
      Begin VB.Menu fv 
         Caption         =   "Search Admin"
      End
      Begin VB.Menu we 
         Caption         =   "Search Employee"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Begin VB.Menu mnuStockRecord 
         Caption         =   "StockRecord Report"
      End
      Begin VB.Menu mnuCustomerReport 
         Caption         =   "CustomerReport"
      End
      Begin VB.Menu mnuDeliveryReport 
         Caption         =   "DeliveryReport"
      End
      Begin VB.Menu mnuRequestRecord 
         Caption         =   "Booking Record Report"
      End
      Begin VB.Menu rth 
         Caption         =   " Other Booking Report"
      End
   End
   Begin VB.Menu mnuBill 
      Caption         =   "Bill"
   End
   Begin VB.Menu mnuAboutUs 
      Caption         =   "AboutUs"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub as_Click()
searchcustomer.Show
Me.Hide

End Sub

Private Sub fdsv_Click()
searchbooking.Show
Me.Hide
End Sub

Private Sub Form_Load()
Login.Hide
End Sub

Private Sub mnuChangePassword_Click()
Form3.Show
Form2.Hide
End Sub


Private Sub fv_Click()
serachadmlogin.Show
Me.Hide
End Sub

Private Sub mnuAboutUs_Click()
frmAbout.Show
End Sub

Private Sub mnuAddCustomer_Click()
customer.Show
main.Hide
End Sub


Private Sub mnuBill_Click()
Bill.Show
main.Hide
End Sub

Private Sub mnuBookRequest_Click()
cylinderbooking.Show
Me.Hide
End Sub

Private Sub mnuCustomerReport_Click()
DataReport2.Show
End Sub

Private Sub mnuCylinderBooking_Click()
cylinderbooking.Show
main.Hide
End Sub

Private Sub mnuCylinderDelivery_Click()
cylinderdelivery.Show
main.Hide

End Sub

Private Sub mnuCylinderRequest_Click()

End Sub

Private Sub mnuDeliveryReport_Click()
DataReport3.Show
End Sub

Private Sub mnuEditCustomer_Click()
editcust.Show
main.Hide
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuLogout_Click()
Login.Show
main.Hide
End Sub

Private Sub mnuStock_Click()
searchemplogin.Show
Me.Hide
End Sub

Private Sub mnuOtherRequest_Click()
otherbooking.Show
main.Hide
End Sub

Private Sub mnuRequestRecord_Click()
DataReport4.Show
End Sub

Private Sub mnuSetPrice_Click()
setprise.Show
Me.Hide
End Sub

Private Sub mnuStockRecord_Click()
DataReport1.Show

End Sub

Private Sub mnuStocskRecord_Click()
stockrecord.Show
Me.Hide
End Sub

Private Sub rth_Click()
DataReport5.Show
End Sub

Private Sub uhj_Click()
serachotherbooking.Show
Me.Hide
End Sub

Private Sub we_Click()
searchemplogin.Show
Me.Hide
End Sub
