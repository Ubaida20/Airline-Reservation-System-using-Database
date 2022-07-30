VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FRMUSER 
   BackColor       =   &H00C0E0FF&
   Caption         =   "USER FORM"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1680
      Top             =   6600
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   1440
      TabIndex        =   17
      Top             =   1440
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7858
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "&CHANGE PASSWORD"
      TabPicture(0)   =   "FRMUSER.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OLDPWD"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "NEWPWD"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CONPWD"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CMDOK"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CMDCANCEL"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CMD_EXIT(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "uname"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "ADD USER"
      TabPicture(1)   =   "FRMUSER.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CMD_EXIT(1)"
      Tab(1).Control(1)=   "CMD_CANCEL"
      Tab(1).Control(2)=   "CMD_OK"
      Tab(1).Control(3)=   "CON_PWD"
      Tab(1).Control(4)=   "NEW_PWD"
      Tab(1).Control(5)=   "ADDUSER"
      Tab(1).Control(6)=   "Label1(5)"
      Tab(1).Control(7)=   "Label1(4)"
      Tab(1).Control(8)=   "Label1(3)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "DELETE USER"
      TabPicture(2)   =   "FRMUSER.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CMD_EXIT(0)"
      Tab(2).Control(1)=   "CCANCEL"
      Tab(2).Control(2)=   "COK"
      Tab(2).Control(3)=   "U_NAME"
      Tab(2).Control(4)=   "Label1(6)"
      Tab(2).ControlCount=   5
      Begin VB.TextBox uname 
         DataField       =   "USER_ID"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   2880
         TabIndex        =   0
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton CMD_EXIT 
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
         Index           =   2
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CommandButton CMD_EXIT 
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
         Index           =   1
         Left            =   -70560
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton CMD_EXIT 
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
         Index           =   0
         Left            =   -70920
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton CCANCEL 
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
         Height          =   495
         Left            =   -72720
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton COK 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&OK"
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
         Left            =   -74520
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox U_NAME 
         DataField       =   "USER_ID"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -72000
         TabIndex        =   13
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CommandButton CMD_CANCEL 
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
         Height          =   495
         Left            =   -72360
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton CMD_OK 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&OK"
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
         Left            =   -74160
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox CON_PWD 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -72000
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox NEW_PWD 
         DataField       =   "PASSWORD"
         DataSource      =   "Adodc1"
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   -72000
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox ADDUSER 
         DataField       =   "USER_ID"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -72000
         TabIndex        =   7
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton CMDCANCEL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "C&ANCEL"
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
         TabIndex        =   5
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CommandButton CMDOK 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&OK"
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox CONPWD 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox NEWPWD 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox OLDPWD 
         DataField       =   "PASSWORD"
         DataSource      =   "Adodc1"
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Left            =   720
         TabIndex        =   25
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "DELETE USER :"
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
         Index           =   6
         Left            =   -74040
         TabIndex        =   24
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CONFIRM PASSWORD :"
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
         Index           =   5
         Left            =   -74400
         TabIndex        =   23
         Top             =   2280
         Width           =   2115
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "NEW PASSWORD :"
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
         Index           =   4
         Left            =   -74325
         TabIndex        =   22
         Top             =   1680
         Width           =   1725
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ADD USER :"
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
         Index           =   3
         Left            =   -74295
         TabIndex        =   21
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CONFIRM PASSWORD :"
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
         Index           =   2
         Left            =   585
         TabIndex        =   20
         Top             =   2760
         Width           =   2115
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "NEW PASSWORD :"
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
         Left            =   675
         TabIndex        =   19
         Top             =   2160
         Width           =   1725
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "OLD PASSWORD  :"
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
         Left            =   645
         TabIndex        =   18
         Top             =   1560
         Width           =   1725
      End
   End
End
Attribute VB_Name = "FRMUSER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CCANCEL_Click()
U_NAME.Text = ""
U_NAME.SetFocus
End Sub

Private Sub CMD_CANCEL_Click()
ADDUSER.Text = ""
NEW_PWD.Text = ""
CON_PWD.Text = ""
ADDUSER.SetFocus
End Sub

Private Sub CMD_EXIT_Click(Index As Integer)
Unload Me
End Sub

Private Sub CMD_OK_Click()
If user = "admin" Then
Adodc1.RecordSource = "SELECT * FROM LOGIN"
'Adodc1.Refresh
For I = 1 To Adodc1.Recordset.RecordCount
If UCase(ADDUSER.Text) = UCase(Adodc1.Recordset.Fields("username")) Then
MsgBox ("Username Already Entered")
Exit Sub
End If
Adodc1.Recordset.MoveNext
Next
If NEW_PWD.Text = CON_PWD.Text Then
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("USERNAME") = ADDUSER.Text
Adodc1.Recordset.Fields("PASS") = NEW_PWD.Text
Adodc1.Recordset.Update
MsgBox ("New user created sucessfully")
ADDUSER.Text = ""
NEW_PWD.Text = ""
CON_PWD.Text = ""
Else
MsgBox ("PASSWORD NOT MATCH")
End If
Else
MsgBox ("Only Admin User Can Create User")
Exit Sub
End If
End Sub

Private Sub CMDCANCEL_Click()
OLDPWD.Text = ""
NEWPWD.Text = ""
CONPWD.Text = ""
OLDPWD.SetFocus
End Sub

Private Sub CMDOK_Click()
Adodc1.RecordSource = "select * from login"
'Adodc1.Refresh
For I = 1 To Adodc1.Recordset.RecordCount
If UCase(uname.Text) = UCase(Adodc1.Recordset.Fields("username")) And OLDPWD.Text = Adodc1.Recordset.Fields("pass") Then
    If NEWPWD.Text = CONPWD.Text Then
    Adodc1.Recordset.Fields("pass") = NEWPWD.Text
    Adodc1.Recordset.Update
    MsgBox ("Password Has Been Changed")
    uname.Text = ""
    OLDPWD.Text = ""
    NEWPWD.Text = ""
    CONPWD.Text = ""
    uname.SetFocus
    Exit Sub
    Else
    MsgBox ("New and Confirm Password not match")
    Exit Sub
    End If
Else
X = 1
End If
Adodc1.Recordset.MoveNext
Next
If X = 1 Then
MsgBox ("User name or Old Password Invalid")
End If
End Sub

Private Sub COK_Click()
If user = "admin" Then
    Adodc1.RecordSource = "SELECT * FROM LOGIN"
 '   Adodc1.Refresh
    If UCase(U_NAME.Text = "admin") Then
    MsgBox ("Admin user Cannot be deleted")
    Exit Sub
    End If
    For I = 1 To Adodc1.Recordset.RecordCount
    
        If UCase(U_NAME.Text) = UCase(Adodc1.Recordset.Fields("username")) Then
        NO = MsgBox("Are you sure you want to delete User", vbYesNo)
            If NO = 6 Then
            Adodc1.Recordset.Delete
            U_NAME.Text = ""
            U_NAME.SetFocus
            Exit Sub
            End If
        Else
            A = 1
    
        End If
    Adodc1.Recordset.MoveNext
    Next
Else
MsgBox ("Only Admin User Can Create User")
Exit Sub
End If
If A = 1 Then
MsgBox ("No Such User Found")
End If
End Sub


Private Sub Form_Load()
'uname.Text = ""
'OLDPWD.Text = ""
'NEWPWD.Text = ""
'CONPWD.Text = ""

'ADDUSER.Text = ""
'NEW_PWD.Text = ""
'CON_PWD.Text = ""

'U_NAME.Text = ""

End Sub


