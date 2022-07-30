VERSION 5.00
Begin VB.MDIForm MDIRESERVATION 
   BackColor       =   &H00C0E0FF&
   Caption         =   "OVERSEAS TRAVELS PVT LTD."
   ClientHeight    =   3195
   ClientLeft      =   1665
   ClientTop       =   1140
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu MASTER 
      Caption         =   "MASTER"
      Begin VB.Menu LOGIN 
         Caption         =   "LOGIN"
         Begin VB.Menu ADDUSER 
            Caption         =   "ADD USER"
         End
         Begin VB.Menu DELETEUSER 
            Caption         =   "DELETE USER"
         End
         Begin VB.Menu CHANGEPASSWORD 
            Caption         =   "CHANGE PASSWORD"
         End
      End
      Begin VB.Menu BRANCH 
         Caption         =   "BRANCH"
      End
      Begin VB.Menu SERVICE 
         Caption         =   "SERVICE"
      End
      Begin VB.Menu AIRBUS 
         Caption         =   "AIRBUS"
      End
      Begin VB.Menu FARE 
         Caption         =   "FARE"
      End
      Begin VB.Menu FLIGHT 
         Caption         =   "FLIGHT INFO"
      End
      Begin VB.Menu FLIGHTSCHEDULE 
         Caption         =   "FLIGHT SCHEDULE"
      End
      Begin VB.Menu CONTROL 
         Caption         =   "CONTROL"
      End
      Begin VB.Menu ROUTE 
         Caption         =   "ROUTE"
      End
      Begin VB.Menu EXIT 
         Caption         =   "EXIT"
      End
   End
   Begin VB.Menu TRANSACTION 
      Caption         =   "TRANSACTION"
      Begin VB.Menu RESERVATION 
         Caption         =   "RESERVATION"
      End
      Begin VB.Menu CANCELLATION 
         Caption         =   "CANCELLATION"
      End
   End
   Begin VB.Menu REPORT 
      Caption         =   "REPORT"
      Begin VB.Menu RESERVATION_REPORT 
         Caption         =   "RESERVATION"
      End
      Begin VB.Menu CANCELLATION_REPORT 
         Caption         =   "CANCELLATION"
      End
      Begin VB.Menu FLIGHTBYBRANCH_REPORT 
         Caption         =   "FLIGHT INFORMATION BY BRANCH"
      End
   End
End
Attribute VB_Name = "MDIRESERVATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ADDUSER_Click()
FRMUSER.Show
FRMUSER.SSTab1.Tab = 1
FRMUSER.ADDUSER.SetFocus
End Sub

Private Sub AIRBUS_Click()
FRMAIRBUS.Show
End Sub

Private Sub BRANCH_Click()
FRMBRANCH.Show
End Sub

Private Sub CANCELLATION_Click()
FRMCANC.Show
End Sub

Private Sub CANCELLATION_REPORT_Click()
CANCEL.Show
End Sub

Private Sub CHANGEPASSWORD_Click()
FRMUSER.Show
FRMUSER.SSTab1.Tab = 0
FRMUSER.uname.SetFocus
End Sub

Private Sub CONTROL_Click()
FRMCONTROL.Show
End Sub

Private Sub DELETEUSER_Click()
FRMUSER.Show
FRMUSER.SSTab1.Tab = 2
FRMUSER.U_NAME.SetFocus
End Sub

Private Sub EXIT_Click()
End
End Sub

Private Sub FARE_Click()
FRMFARE.Show
End Sub

Private Sub FLIGHT_Click()
FRMFLIGHT.Show
End Sub

Private Sub FLIGHTBYBRANCH_REPORT_Click()
BREPORT.Show
End Sub

Private Sub FLIGHTSCHEDULE_Click()
FRMFLIGHT_SCH.Show
End Sub

Private Sub RESERVATION_Click()
FRMRESERVATION.Show
End Sub

Private Sub RESERVATION_REPORT_Click()
RESERVE.Show
End Sub

Private Sub ROUTE_Click()
FRMROUTE.Show
End Sub

Private Sub SERVICE_Click()
FRMSERVICE.Show
End Sub

Public Sub comm_enb()

End Sub
