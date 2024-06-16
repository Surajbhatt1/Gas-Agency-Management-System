VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form editcust 
   Caption         =   "Form4"
   ClientHeight    =   8820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11835
   LinkTopic       =   "Form4"
   ScaleHeight     =   8820
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Search"
      Height          =   495
      Left            =   7680
      TabIndex        =   18
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4560
      TabIndex        =   17
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   16
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   15
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   14
      Top             =   7560
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4440
      TabIndex        =   13
      Top             =   6480
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4440
      TabIndex        =   12
      Top             =   5760
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   4080
      Width           =   2775
   End
   Begin VB.ComboBox Combo2 
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Text            =   "Combo2"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   2520
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6840
      Top             =   7560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form4.frx":0000
      OLEDBString     =   $"Form4.frx":0097
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "customer"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   11
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Pincode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Connection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   8760
      Left            =   0
      Picture         =   "Form4.frx":012E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11880
   End
End
Attribute VB_Name = "editcust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.Update
'Adodc1.Recordset.Fields("ID") = Text1.Text
Adodc1.Recordset.Fields("custname") = Text1.Text
Adodc1.Recordset.Fields("Type") = Combo2.Text
Adodc1.Recordset.Fields("Connect") = Text2.Text
Adodc1.Recordset.Fields("Address") = Text3.Text
Adodc1.Recordset.Fields("Pincode") = Text4.Text
Adodc1.Recordset.Fields("Phno") = Text5.Text
Adodc1.Recordset.Update
MsgBox "Data Update"
Adodc1.Recordset.MoveLast
Text1.Text = Adodc1.Recordset.Fields("Custno") + 1
Text1.Enabled = False
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

Private Sub Command2_Click()
On Error GoTo errmsg

Adodc1.Recordset.Find "Custno=" & Val(Text6.Text)
Adodc1.Recordset.Delete
MsgBox "Successfully Deleted"


Exit Sub
errmsg:
MsgBox "Error In Deleting"
End Sub

Private Sub Command3_Click()
main.Show
editcust.Hide
End Sub

Private Sub Command4_Click()
On Error GoTo errmsg
Adodc1.Refresh
Command1.Visible = True
 Adodc1.Recordset.Find "Custno=" & Val(Text6.Text)


    

'Text1.Text = Adodc1.Recordset.Fields("Custno")
Text1.Text = Adodc1.Recordset.Fields("custname")
Combo2.Text = Adodc1.Recordset.Fields("Type")
Text2.Text = Adodc1.Recordset.Fields("Connect")
Text3.Text = Adodc1.Recordset.Fields("Address")
Text4.Text = Adodc1.Recordset.Fields("Pincode")
Text5.Text = Adodc1.Recordset.Fields("Phno")
 Exit Sub
errmsg:
MsgBox "record not exist"
End Sub
