VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMLOGIN 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN"
   ClientHeight    =   3405
   ClientLeft      =   1830
   ClientTop       =   1335
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2730
      Top             =   2730
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=AIRLINE"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "AIRLINE"
      OtherAttributes =   ""
      UserName        =   "scott"
      Password        =   "tiger"
      RecordSource    =   "LOGIN"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   5055
      Begin VB.CommandButton CMDCANCEL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&CANCEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CMDLOGIN 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&LOGIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   1815
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5055
      Begin VB.TextBox UNAME 
         DataField       =   "USER_ID"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox PWD 
         DataField       =   "PASSWORD"
         DataSource      =   "Adodc1"
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "USER NAME :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1230
      End
   End
End
Attribute VB_Name = "FRMLOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDCANCEL_Click()
End
End Sub

Private Sub CMDLOGIN_Click()
Adodc1.RecordSource = "select * from login"
'Adodc1.Refresh
For I = 1 To Adodc1.Recordset.RecordCount
If UCase(UNAME.Text) = UCase(Adodc1.Recordset.Fields.Item(0).Value) Then
        If PWD.Text = Adodc1.Recordset.Fields(1) Then
        MDIRESERVATION.Show
        FRMLOGIN.Hide
        user = UNAME.Text
        Exit Sub
        Else
        MsgBox ("Invalid Password")
        Exit Sub
        End If
Else
A = 1

End If
Adodc1.Recordset.MoveNext
Next
If A = 1 Then
MsgBox ("Invalid Username")
End If
End Sub

Private Sub Form_Activate()
UNAME.Text = ""
PWD.Text = ""
UNAME.SetFocus
End Sub

