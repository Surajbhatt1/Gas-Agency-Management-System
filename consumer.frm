VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   14355
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   3120
      TabIndex        =   1
      Top             =   1200
      Width           =   7935
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   4440
         TabIndex        =   16
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Update"
         Height          =   495
         Left            =   1200
         TabIndex        =   15
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   5400
         MaxLength       =   11
         TabIndex        =   14
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1440
         MaxLength       =   49
         TabIndex        =   13
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1440
         MaxLength       =   49
         TabIndex        =   12
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1440
         MaxLength       =   19
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1440
         MaxLength       =   19
         TabIndex        =   10
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   495
         Left            =   3000
         TabIndex        =   8
         Top             =   3480
         Width           =   1215
      End
      Begin VB.PictureBox DTPicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd.MM.yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   5400
         ScaleHeight     =   315
         ScaleWidth      =   1875
         TabIndex        =   7
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   5400
         MaxLength       =   6
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   5280
         TabIndex        =   3
         Top             =   2640
         Width           =   2535
         Begin VB.OptionButton Option2 
            Caption         =   "Commercial"
            Height          =   255
            Left            =   1200
            TabIndex        =   5
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Domestic"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.PictureBox DTPicker2 
         Height          =   375
         Left            =   5400
         ScaleHeight     =   315
         ScaleWidth      =   1875
         TabIndex        =   2
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Cylinder Type"
         Height          =   255
         Left            =   4200
         TabIndex        =   25
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Date of Connection"
         Height          =   255
         Left            =   3840
         TabIndex        =   24
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Phone"
         Height          =   255
         Left            =   4560
         TabIndex        =   23
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Address Line2"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Address Line1"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "First Name"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Consumer No:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Pin"
         Height          =   255
         Left            =   4800
         TabIndex        =   18
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Last Booking"
         Height          =   375
         Left            =   4200
         TabIndex        =   17
         Top             =   2160
         Width           =   975
      End
   End
   Begin VB.Label Vconsumer 
      Caption         =   "Consumer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If (Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "") Then
MsgBox "Enter all Fields", vbOKOnly, "Error"
Else
rs(0) = Text1.Text
rs(1) = Text2.Text
rs(2) = Text3.Text
rs(3) = Text4.Text
rs(4) = Text5.Text
rs(5) = Text6.Text
rs(6) = Text7.Text
rs(7) = DTPicker1.Value
If Option1.Value = True Then
rs(9) = Option1.Caption
Else
rs(9) = Option2.Caption
End If
rs.Update
MsgBox "Record updated Sucessfully", vbOKOnly, "Update"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If (Text1.Text = "") Then
MsgBox "No Consumer", vbOKOnly, "Error"
Else
rs(9) = "Deleted"
rs.Update
MsgBox "Consumer deleted", vbOKOnly, "Delete"
Unload Me
End If
End Sub
