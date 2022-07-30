VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMCONTROL 
   BackColor       =   &H00C0E0FF&
   Caption         =   "CONTROL INFORMATION"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   1935
      Left            =   960
      TabIndex        =   26
      Top             =   360
      Width           =   9495
      Begin VB.Image Image1 
         Height          =   1680
         Left            =   120
         Picture         =   "FRMCONTROL.frx":0000
         Top             =   120
         Width           =   1920
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AIRLINE RESERVATION SYSTEM"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3120
         TabIndex        =   28
         Top             =   480
         Width           =   5085
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CONTROL INFORMATION"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         TabIndex        =   27
         Top             =   1200
         Width           =   3045
      End
   End
   Begin VB.Frame FRMBUTTON 
      BackColor       =   &H0080C0FF&
      Height          =   855
      Left            =   960
      TabIndex        =   15
      Top             =   6360
      Width           =   9495
      Begin VB.CommandButton CMDADD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CMDSAVE 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CMDEDIT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "EDIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CMDEXIT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CMDCANCEL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox CAN_D_3 
      DataField       =   "CANC_DEDUC_3"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   8160
      MaxLength       =   3
      TabIndex        =   7
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox CAN_D_6 
      DataField       =   "CANC_DEDUC_6"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   8160
      MaxLength       =   3
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox CAN_D_12 
      DataField       =   "CANC_DEDUC_12"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   8160
      MaxLength       =   3
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox E_BG_LIMIT 
      DataField       =   "ECO_BG_LIMIT"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox B_BG_LIMIT 
      DataField       =   "BUS_BG_LIMIT"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox F_BG_LIMIT 
      DataField       =   "FIRST_BG_LIMIT"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox EX_BG_LIMIT 
      DataField       =   "EXCESS_BG_LIMIT"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox AIR_TAX 
      DataField       =   "AIR_TAX"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3720
      MaxLength       =   4
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame FRMINFO 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   960
      TabIndex        =   13
      Top             =   2280
      Width           =   9495
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   4920
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   16777215
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=AIRLINE"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "AIRLINE"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   "TYBCA34"
         RecordSource    =   "CONTROL"
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
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "AIR TAX :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "BAG LIMIT :--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   24
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "EXCESS BAG LIMIT :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "FIRST BAG LIMIT :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "BUS  BAG LIMIT :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   21
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ECO BAG LIMIT :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080C0FF&
         Caption         =   "CANCEL DEDUCTION 12 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   19
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080C0FF&
         Caption         =   "CANCEL DEDUCTION 6 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080C0FF&
         Caption         =   "CANCEL DEDUCTION 3 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   17
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080C0FF&
         Caption         =   "CANCELLATION DEDUCTION :--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   16
         Top             =   1200
         Width           =   2895
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   5400
      TabIndex        =   14
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "FRMCONTROL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub B_BG_LIMIT_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       E_BG_LIMIT.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER NUMERIC VALUE ONLY", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
End Sub

Private Sub CAN_D_12_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       CAN_D_6.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER NUMERIC VALUE ONLY", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
End Sub

Private Sub CAN_D_3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       CMDSAVE.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER NUMERIC VALUE ONLY", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
End Sub

Private Sub CAN_D_6_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       CAN_D_3.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER NUMERIC VALUE ONLY", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
End Sub

Private Sub CMDADD_Click()
Adodc1.Refresh
Adodc1.Recordset.AddNew
EX_BG_LIMIT.SetFocus
CMDADD.Enabled = False
CMDEDIT.Enabled = False
CMDCANCEL.Enabled = True
CMDSAVE.Enabled = True
CMDEXIT.Enabled = True
End Sub

Private Sub CMDCANCEL_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
CMDCANCEL.Enabled = False
Adodc1.Recordset.CANCEL
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
CMDADD.Enabled = True
CMDSAVE.Enabled = False
CMDEXIT.Enabled = True
CMDEDIT.Enabled = True
End Sub




Private Sub CMDEDIT_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
Adodc1.Recordset.Update
CMDADD.Enabled = False
CMDEDIT.Enabled = False
CMDCANCEL.Enabled = True
CMDSAVE.Enabled = True
CMDEXIT.Enabled = True
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub



Private Sub CMDSAVE_Click()
If AIR_TAX.Text = "" Or EX_BG_LIMIT.Text = "" Or F_BG_LIMIT.Text = "" Or B_BG_LIMIT.Text = "" Or E_BG_LIMIT.Text = "" Or CAN_D_12.Text = "" Or CAN_D_6.Text = "" Or CAN_D_3.Text = "" Then
  MsgBox ("INCOMPLETE RECORD")
  Exit Sub
End If
CMDADD.Enabled = False
CMDSAVE.Enabled = False
CMDCANCEL.Enabled = False
CMDEXIT.Enabled = True
CMDEDIT.Enabled = True
A = MsgBox("DO YOU WANT TO SAVE RECORD", vbYesNo, "MESSAGE")
If A = 6 Then
   Adodc1.Recordset.Save
Else
   CLEAR
End If
End Sub

Public Sub CLEAR()
AIR_TAX.Text = ""
EX_BG_LIMIT.Text = ""
F_BG_LIMIT.Text = ""
B_BG_LIMIT.Text = ""
E_BG_LIMIT.Text = ""
CAN_D_12.Text = ""
CAN_D_6.Text = ""
CAN_D_3.Text = ""
End Sub


Private Sub E_BG_LIMIT_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       CAN_D_12.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER NUMERIC VALUE ONLY", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
End Sub

Private Sub EX_BG_LIMIT_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       F_BG_LIMIT.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER NUMERIC VALUE ONLY", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
End Sub


Private Sub F_BG_LIMIT_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       B_BG_LIMIT.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER NUMERIC VALUE ONLY", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
End Sub

Private Sub Form_Load()
Adodc1.Refresh
CMDSAVE.Enabled = False
CMDCANCEL.Enabled = False
If Adodc1.Recordset.RecordCount > 0 Then
  CMDADD.Enabled = False
End If



End Sub

