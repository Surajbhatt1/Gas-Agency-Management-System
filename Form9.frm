VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form cylinderdelivery 
   BackColor       =   &H80000016&
   Caption         =   "Form9"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12120
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   7440
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   7800
      Top             =   6360
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
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
      Connect         =   $"Form9.frx":941B
      OLEDBString     =   $"Form9.frx":94B2
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "delivery"
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
   Begin VB.TextBox Text6 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   7080
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
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
      Left            =   5520
      TabIndex        =   5
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delivered"
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
      Left            =   3720
      TabIndex        =   4
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cylinder Delivery"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   13
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   2400
      TabIndex        =   12
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Delivery"
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
      Left            =   2280
      TabIndex        =   11
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of booking"
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
      Left            =   2280
      TabIndex        =   10
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
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
      Left            =   2400
      TabIndex        =   9
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   7380
      Left            =   0
      Picture         =   "Form9.frx":9549
      Stretch         =   -1  'True
      Top             =   120
      Width           =   12120
   End
End
Attribute VB_Name = "cylinderdelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("ID") = Text6.Text
Adodc1.Recordset.Fields("C_ID") = Text1.Text
Adodc1.Recordset.Fields("dob") = Text2.Text
Adodc1.Recordset.Fields("dod") = Text3.Text
Adodc1.Recordset.Fields("status") = Text4.Text
Adodc1.Recordset.Fields("type") = Text5.Text

Adodc1.Recordset.Update
MsgBox "Added Successfully "

Adodc1.Refresh
End Sub

Private Sub Command2_Click()
main.Show
Me.Hide
End Sub

Private Sub Form_Load()
Adodc1.Recordset.MoveLast
Text6.Text = Adodc1.Recordset.Fields("ID") + 1
Text6.Enabled = False
End Sub

