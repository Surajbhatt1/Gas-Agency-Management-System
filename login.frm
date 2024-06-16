VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Login 
   BackColor       =   &H00FF00FF&
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   Picture         =   "login.frx":0000
   ScaleHeight     =   11020
   ScaleMode       =   0  'User
   ScaleWidth      =   14280
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc loginado 
      Height          =   855
      Left            =   5880
      Top             =   5640
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   $"login.frx":32A6
      OLEDBString     =   $"login.frx":333D
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * From Login"
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
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   6720
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6720
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancle"
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
      Left            =   7320
      TabIndex        =   1
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
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
      Left            =   5520
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "    LOGIN"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   4680
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
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
      Left            =   4680
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
loginado.RecordSource = "Select * From Login Where UserID='" + Text1.Text + "' and Password='" + Text2.Text + "'"
loginado.Refresh
If loginado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "massege"
Else
MsgBox ("WellCome!!! " & loginado.Recordset.Fields(1))
 Me.Hide
 main.Show
 Exit Sub
errmsg:
MsgBox "Invalid Crediantals"
loginado.Caption = loginado.RecordSource
End If


End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
End Sub
