VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4065
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   4000
         Left            =   3000
         Top             =   1860
      End
      Begin VB.Frame Frame2 
         Height          =   4050
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7080
         Begin VB.Label lblCopyright 
            BackStyle       =   0  'Transparent
            Caption         =   "@Copyright"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4560
            TabIndex        =   8
            Top             =   3060
            Width           =   2415
         End
         Begin VB.Label lblCompany 
            BackStyle       =   0  'Transparent
            Caption         =   "All Rights Are Reserved"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4560
            TabIndex        =   7
            Top             =   3270
            Width           =   2415
         End
         Begin VB.Label lblWarning 
            BackStyle       =   0  'Transparent
            Caption         =   "Warning:Do Not Try Unauthorised Access. It Is Harmfull For System"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   150
            TabIndex        =   6
            Top             =   3660
            Width           =   6855
         End
         Begin VB.Label lblVersion 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Version :1.0.0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5400
            TabIndex        =   5
            Top             =   2700
            Width           =   1560
         End
         Begin VB.Label lblPlatform 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Platform : Visual Basic 6.0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2925
            TabIndex        =   4
            Top             =   2340
            Width           =   3930
         End
         Begin VB.Label lblLicenseTo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "LicenseTo Mr.Suraj Bhatt &.Mr.Shivratan Khambat"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   6855
         End
         Begin VB.Label lblCompanyProduct 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gas Agency ManagementSystem"
            BeginProperty Font 
               Name            =   "Baskerville Old Face"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   0
            TabIndex        =   2
            Top             =   480
            Width           =   6585
         End
         Begin VB.Image imgLogo 
            Height          =   3870
            Left            =   0
            Picture         =   "frmSplash.frx":000C
            Stretch         =   -1  'True
            Top             =   75
            Width           =   7095
         End
      End
      Begin VB.Timer Timer2 
         Interval        =   4000
         Left            =   2160
         Top             =   1920
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
If Timer2.Interval = 4000 Then
Wellcome.Show

Unload Me
End If
End Sub
