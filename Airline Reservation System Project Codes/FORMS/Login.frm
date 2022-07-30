VERSION 5.00
Begin VB.Form FRMLOGIN 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TMP 
      Height          =   285
      Left            =   0
      MaxLength       =   1
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0080C0FF&
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    Me.Hide
End Sub
Private Sub cmdOK_Click()
Dim z As Integer
rs.Open "select * from SYS_PASS WHERE UPPER(USER_ID)='" & Trim(UCase(txtUserName.Text)) & "'", cn, adOpenDynamic, adLockReadOnly
If rs.RecordCount = 0 Then
     z = MsgBox("User Not Found", vbOKOnly, "WARRNING")
     txtPassword.Text = ""
     txtUserName.Text = ""
     txtUserName.SetFocus
     rs.Close
     Exit Sub
Else
    'check for correct password
    Encrip (txtPassword)
    If txtPassword = rs(2) Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        TMP = rs(1)
        rs.Close
        Me.Hide
        Form2.Show
    Else
        rs.Close
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End If
End Sub

Private Sub Form_Load()
cn.Open "provider=msdaora.1;user id=SUJAL;password=SUJAL;data source=aisectnet"
cn.CursorLocation = adUseClient
End Sub
Public Sub Encrip(s As String)
Dim a, l, i As Integer
Dim str As String
Dim c As String
l = Len(s)
For i = 1 To l
   c = Mid(s, i, 1)
   str = str & Chr(Asc(c) + i)
Next
txtPassword.Text = str
End Sub
