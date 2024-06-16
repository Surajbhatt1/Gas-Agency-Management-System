VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10740
   LinkTopic       =   "Form10"
   ScaleHeight     =   5850
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
Form10.Hide

End Sub
