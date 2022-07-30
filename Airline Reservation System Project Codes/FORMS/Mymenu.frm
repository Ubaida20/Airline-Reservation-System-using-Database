VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daxesh Software Solution"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu Master 
      Caption         =   "&Master"
      Begin VB.Menu Password 
         Caption         =   "&Password"
      End
   End
   Begin VB.Menu Exit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Exit_Click()
z = MsgBox("Do Want to Exit System", vbYesNo, "Alert")
If z = vbYes Then
   End
End If
End Sub
Private Sub Password_Click()
If frmLogin.TMP = "A" Then
   Form1.Show
Else
  z = MsgBox("This Option Is Only For Administrator", vbOKOnly, "Alert")
End If
End Sub
