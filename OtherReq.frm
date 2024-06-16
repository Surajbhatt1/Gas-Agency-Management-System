VERSION 5.00
Begin VB.Form OtherReq 
   BackColor       =   &H80000014&
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Burner Stove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080FFFF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000018&
      Caption         =   "Lighter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Other's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Gas-Pip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Regulator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      MaskColor       =   &H00004040&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   6840
      Left            =   0
      Picture         =   "OtherReq.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "OtherReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Gaspip.Show
OtherReq.Hide
End Sub

Private Sub Command3_Click()
Burner.Show
OtherReq.Hide
End Sub

Private Sub Command7_Click()
main.Show
OtherReq.Hide
End Sub
