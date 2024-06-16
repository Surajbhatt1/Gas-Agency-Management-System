VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form SLogin 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   13560
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc loginado 
      Height          =   735
      Left            =   9960
      Top             =   7560
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
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
      Connect         =   $"SLogin.frx":0000
      OLEDBString     =   $"SLogin.frx":0097
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Emplog"
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
      Left            =   7440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3480
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   2520
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      Picture         =   "SLogin.frx":012E
      TabIndex        =   1
      Top             =   7440
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000080&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      TabIndex        =   0
      Top             =   7440
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "   STAFF LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   4800
      TabIndex        =   6
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "UserId"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   -360
      Picture         =   "SLogin.frx":2415
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13920
   End
End
Attribute VB_Name = "SLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
loginado.RecordSource = "select * from Emplog where UserId='" + Text1.Text + "'and Password='" + Text2.Text + "'"
loginado.Refresh
If loginado.Recordset.EOF Then
MsgBox "Login failed,Try Again..!!!", vbCritical, "Please Enter correct UserId and Password"
Else
MsgBox "Login Successful.", vbInformation, "Successful Attempt"
main.Show
Login.Hide
End If
End Sub


Private Sub Command2_Click()
Wellcome.Show
SLogin.Hide
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
End Sub
