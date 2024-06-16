VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form Splash 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Progress"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   17340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   12840
      Top             =   120
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   6360
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   1508
      _Version        =   327682
      Appearance      =   1
      Max             =   105
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   4095
      Left            =   4920
      Picture         =   "Splash.frx":0000
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "                           GAS AGENCY                                                       MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   15615
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
Label2.Caption = "Loading..."
Label3.Caption = ProgressBar1.Value & "%"
If (ProgressBar1.Value = ProgressBar1.Max) Then
Timer1.Enabled = False
Unload Me
Wellcome.Show
Splash.Hide
End If
End Sub

