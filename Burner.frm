VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form otherbooking 
   BackColor       =   &H00FFFF80&
   Caption         =   "Form1"
   ClientHeight    =   10050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   ScaleHeight     =   10050
   ScaleWidth      =   14430
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9360
      Top             =   7440
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
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
      Connect         =   $"Burner.frx":0000
      OLEDBString     =   $"Burner.frx":0097
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "otherbooking"
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
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   4080
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   5
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   6840
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   4920
      TabIndex        =   3
      Text            =   "Select Type"
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   9480
      Top             =   6360
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Connect         =   $"Burner.frx":012E
      OLEDBString     =   $"Burner.frx":01C5
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
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      TabIndex        =   11
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer phno"
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
      Left            =   1440
      TabIndex        =   10
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cust.No"
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
      TabIndex        =   8
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "New other Booking"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   7
      Top             =   240
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   13530
      Left            =   -120
      Picture         =   "Burner.frx":025C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19500
   End
End
Attribute VB_Name = "otherbooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Combo1_Change()
Label2.Caption = Combo1.Text
End Sub

Private Sub Combo2_Change()
Label5.Caption = Combo2.Text
End Sub

Private Sub Command1_Click()
Adodc1.Recordset.MoveLast
Dim a As Integer
a = Adodc1.Recordset.Fields("ID") + 1

Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("cid") = Text1.Text
Adodc1.Recordset.Fields("name") = Text2.Text
Adodc1.Recordset.Fields("phno") = Text3.Text
Adodc1.Recordset.Fields("add") = Text4.Text
Adodc1.Recordset.Fields("type") = Combo2.Text
Adodc1.Recordset.Update
MsgBox "Added Successfully "

Adodc1.Refresh


End Sub

Private Sub Command2_Click()
main.Show
End Sub

Private Sub Form_Load()

Combo2.AddItem ("Burners")
Combo2.AddItem ("Burner_Stove")
Combo2.AddItem ("Gas Pipe")
Combo2.AddItem ("Lighter")
Combo2.AddItem ("Other")
End Sub


Private Sub Text2_Click()
On Error GoTo errmsg
Adodc2.Refresh

 Adodc2.Recordset.Find "Custno=" & Val(Text1.Text)


    

'Text1.Text = Adodc1.Recordset.Fields("Custno")
Text2.Text = Adodc2.Recordset.Fields("custname")

Text3.Text = Adodc2.Recordset.Fields("Phno")
Text4.Text = Adodc2.Recordset.Fields("Address")

 Exit Sub
errmsg:
MsgBox "record not exist"
End Sub
