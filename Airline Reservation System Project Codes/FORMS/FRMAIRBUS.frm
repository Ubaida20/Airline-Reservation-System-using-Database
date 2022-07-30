VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMAIRBUS 
   BackColor       =   &H00C0E0FF&
   Caption         =   "AIRBUS INFORMATION"
   ClientHeight    =   8100
   ClientLeft      =   210
   ClientTop       =   1005
   ClientWidth     =   11685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9783.799
   ScaleMode       =   0  'User
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   1935
      Left            =   480
      TabIndex        =   29
      Top             =   120
      Width           =   11055
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
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
         Left            =   3240
         TabIndex        =   31
         Top             =   480
         Width           =   5085
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "AIRBUS INFORMATION"
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
         Left            =   4320
         TabIndex        =   30
         Top             =   1320
         Width           =   2745
      End
      Begin VB.Image Image1 
         Height          =   1680
         Left            =   120
         Picture         =   "FRMAIRBUS.frx":0000
         Top             =   150
         Width           =   1920
      End
   End
   Begin VB.Frame FRMBUTTON 
      BackColor       =   &H0080C0FF&
      Height          =   1455
      Left            =   480
      TabIndex        =   19
      Top             =   6480
      Width           =   11055
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
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
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
         Left            =   1680
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CMDDELETE 
         BackColor       =   &H00E0E0E0&
         Caption         =   "DELETE"
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
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CMDFIRST 
         BackColor       =   &H00E0E0E0&
         Caption         =   "FIRST"
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
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton CMDNEXT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "NEXT"
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton CMDPREVIOUS 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PREVIOUS"
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton CMDLAST 
         BackColor       =   &H00E0E0E0&
         Caption         =   "LAST"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   840
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
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton CMDFIND 
         BackColor       =   &H00E0E0E0&
         Caption         =   "FIND"
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
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
      Height          =   4455
      Left            =   480
      TabIndex        =   18
      Top             =   2040
      Width           =   11055
      Begin VB.TextBox FIRST_WL_CAP 
         DataField       =   "FIRST_WL_CAP"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   8520
         MaxLength       =   3
         TabIndex        =   4
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox BUS_WL_CAP 
         DataField       =   "BUS_WL_CAP"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   8520
         MaxLength       =   3
         TabIndex        =   5
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox ECO_WL_CAP 
         DataField       =   "ECO_WL_CAP"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   8520
         MaxLength       =   3
         TabIndex        =   6
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox FIRST_CAP 
         DataField       =   "FIRST_CAP"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3360
         MaxLength       =   3
         TabIndex        =   1
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox BUS_CAP 
         DataField       =   "BUS_CAP"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3360
         MaxLength       =   3
         TabIndex        =   2
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox ECO_CAP 
         DataField       =   "ECO_CAP"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3360
         MaxLength       =   3
         TabIndex        =   3
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox airbus_no 
         DataField       =   "AIRBUSNO"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3360
         MaxLength       =   5
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   7560
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
         UserName        =   ""
         Password        =   "TYBCA34"
         RecordSource    =   "AIRBUS"
         Caption         =   "Adodc2"
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   7560
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
         UserName        =   ""
         Password        =   "TYBCA34"
         RecordSource    =   "AIRBUS"
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
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "WAITING LIST CAPACITY :-"
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
         Left            =   5445
         TabIndex        =   28
         Top             =   1560
         Width           =   2565
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "FIRST CLASS W.L CAP :"
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
         Left            =   5520
         TabIndex        =   27
         Top             =   2280
         Width           =   2145
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "BUSINESS CLASS W.L CAP :"
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
         Left            =   5520
         TabIndex        =   26
         Top             =   3000
         Width           =   2535
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ECONOMICAL CLASS W. L CAP :"
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
         Left            =   5520
         TabIndex        =   25
         Top             =   3720
         Width           =   2865
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "AIRBUS NO :"
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
         Left            =   570
         TabIndex        =   24
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER CAPACITY : -"
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
         Left            =   360
         TabIndex        =   23
         Top             =   1440
         Width           =   2265
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "FIRST CLASS CAP :"
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
         Left            =   480
         TabIndex        =   22
         Top             =   2280
         Width           =   1740
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "BUSINESS CLASS CAP :"
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
         Left            =   480
         TabIndex        =   21
         Top             =   3000
         Width           =   2130
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ECONOMICAL CLASS CAP :"
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
         Left            =   480
         TabIndex        =   20
         Top             =   3720
         Width           =   2400
      End
   End
End
Attribute VB_Name = "FRMAIRBUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim abus As String
Dim chk As Integer
Private Sub AIRBUS_NO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
If KeyAscii = 45 Then Exit Sub
  If KeyAscii < 48 Or KeyAscii > 57 And KeyAscii < 65 Or KeyAscii > 90 Then
     If KeyAscii = 13 Then
        FIRST_CAP.SetFocus
        Exit Sub
     End If
     If KeyAscii <> 8 Then
        I = MsgBox("ENTER NUMERIC VALUE", vbOKOnly, "ALERT")
        KeyAscii = 0
     End If
End If
End Sub

Private Sub BUS_CAP_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
        ECO_CAP.SetFocus
        Exit Sub
     End If
     If KeyAscii <> 8 Then
        I = MsgBox("ENTER NUMERIC VALUE", vbOKOnly, "ALERT")
        KeyAscii = 0
     End If
End If
End Sub

Private Sub BUS_WL_CAP_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       ECO_WL_CAP.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER NUMERIC VALUE", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
End Sub

Private Sub CMDADD_Click()
chk = 1
TXT_LOCK (1)
Adodc1.Refresh
Adodc1.Recordset.AddNew
AIRBUS_NO.SetFocus
CMDCANCEL.Enabled = True
CMDADD.Enabled = False
CMDEDIT.Enabled = False
CMDDELETE.Enabled = False
CMDFIRST.Enabled = False
CMDNEXT.Enabled = False
CMDPREVIOUS.Enabled = False
CMDLAST.Enabled = False
CMDFIND.Enabled = False
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
CMDDELETE.Enabled = True
CMDADD.Enabled = True
CMDSAVE.Enabled = False
CMDEXIT.Enabled = True
CMDEDIT.Enabled = True
CMDFIRST.Enabled = True
CMDPREVIOUS.Enabled = True
CMDLAST.Enabled = True
CMDFIND.Enabled = True
CMDNEXT.Enabled = True
End Sub

Private Sub CMDDELETE_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
CMDADD.Enabled = True
CMDSAVE.Enabled = False
CMDEDIT.Enabled = True
CMDFIRST.Enabled = True
CMDPREVIOUS.Enabled = True
CMDLAST.Enabled = True
CMDFIND.Enabled = True
CMDNEXT.Enabled = True
CMDCANCEL.Enabled = False
A = MsgBox("DO YOU WANT TO DELETE", vbYesNo, "MESSAGE")
If A = 6 Then
  With Adodc1.Recordset
      .Delete
      .MoveNext
      If .EOF = True Then .MoveLast
  End With
End If
End Sub

Private Sub CMDEDIT_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
chk = 2
CMDADD.Enabled = False
CMDEDIT.Enabled = False
CMDDELETE.Enabled = False
CMDFIRST.Enabled = False
CMDNEXT.Enabled = False
CMDPREVIOUS.Enabled = False
CMDLAST.Enabled = False
CMDFIND.Enabled = False
CMDCANCEL.Enabled = True
CMDSAVE.Enabled = True
CMDEXIT.Enabled = True
TXT_LOCK (1)
End Sub
Private Sub CMDEXIT_Click()
Unload Me
End Sub
Private Sub CMDFIND_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
A = InputBox("ENTER AIRBUS NO:")
If Trim(A) = "" Then
  Exit Sub
End If
Adodc1.Recordset.MoveFirst
For I = 1 To Adodc1.Recordset.RecordCount
  If UCase(A) = UCase(Adodc1.Recordset.Fields("AIRBUSNO")) Then
     b = 1
Exit For
  End If
  Adodc1.Recordset.MoveNext
Next
If b = 1 Then
  AIRBUS_NO.Text = Adodc1.Recordset.Fields("AIRBUSNO")
  FIRST_CAP.Text = Adodc1.Recordset.Fields("FIRST_CAP")
  BUS_CAP.Text = Adodc1.Recordset.Fields("BUS_CAP")
  ECO_CAP.Text = Adodc1.Recordset.Fields("ECO_CAP")
  FIRST_WL_CAP.Text = Adodc1.Recordset.Fields("FIRST_WL_CAP")
  BUS_WL_CAP.Text = Adodc1.Recordset.Fields("BUS_WL_CAP")
  ECO_WL_CAP.Text = Adodc1.Recordset.Fields("ECO_WL_CAP")
Else
  MsgBox ("NO RECORD FOUND")
  Adodc1.Refresh
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub CMDFIRST_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
Adodc1.Recordset.MoveFirst
CMDDELETE.Enabled = True
CMDADD.Enabled = True
CMDCANCEL.Enabled = False
CMDSAVE.Enabled = False
CMDEXIT.Enabled = True
CMDEDIT.Enabled = True
CMDFIRST.Enabled = True
CMDPREVIOUS.Enabled = True
CMDLAST.Enabled = True
CMDFIND.Enabled = True
CMDNEXT.Enabled = True
End Sub

Private Sub CMDLAST_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
Adodc1.Recordset.MoveLast
CMDDELETE.Enabled = True
CMDADD.Enabled = True
CMDSAVE.Enabled = False
CMDCANCEL.Enabled = False
CMDEXIT.Enabled = True
CMDEDIT.Enabled = True
CMDFIRST.Enabled = True
CMDPREVIOUS.Enabled = True
CMDLAST.Enabled = True
CMDFIND.Enabled = True
CMDNEXT.Enabled = True
End Sub

Private Sub CMDNEXT_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
   Adodc1.Recordset.MoveLast
   MsgBox ("POINTER IS ON LAST RECORD")
End If
CMDDELETE.Enabled = True
CMDADD.Enabled = True
CMDSAVE.Enabled = False
CMDCANCEL.Enabled = False
CMDEXIT.Enabled = True
CMDEDIT.Enabled = True
CMDFIRST.Enabled = True
CMDPREVIOUS.Enabled = True
CMDLAST.Enabled = True
CMDFIND.Enabled = True
CMDNEXT.Enabled = True
End Sub

Private Sub CMDPREVIOUS_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
   Adodc1.Recordset.MoveFirst
   MsgBox ("PONTER IS ON THE FIRST RECORD")
End If
CMDDELETE.Enabled = True
CMDADD.Enabled = True
CMDSAVE.Enabled = False
CMDCANCEL.Enabled = False
CMDEXIT.Enabled = True
CMDEDIT.Enabled = True
CMDFIRST.Enabled = True
CMDPREVIOUS.Enabled = True
CMDLAST.Enabled = True
CMDFIND.Enabled = True
CMDNEXT.Enabled = True
End Sub

Private Sub CMDSAVE_Click()
If AIRBUS_NO.Text = "" Or FIRST_CAP.Text = "" Or BUS_CAP.Text = "" Or ECO_CAP.Text = "" Or FIRST_WL_CAP.Text = "" Or BUS_WL_CAP.Text = "" Or ECO_WL_CAP.Text = "" Then
   MsgBox ("INCOMLETE RECORD")
Exit Sub
End If
If Adodc1.Recordset.RecordCount > Adodc2.Recordset.RecordCount Then
   If Adodc2.Recordset.RecordCount > 0 Then Adodc2.Recordset.MoveFirst
      For I = 1 To Adodc2.Recordset.RecordCount
         If Adodc2.Recordset(0) = AIRBUS_NO Then
            J = MsgBox("This Airbus No is already Exist", vbOKOnly + vbCritical, "Airbus Reservation")
            SendKeys "{Home}+{End}"
           AIRBUS_NO.SetFocus
            Exit Sub
         End If
       Adodc2.Recordset.MoveNext
    Next
End If
CMDDELETE.Enabled = True
CMDADD.Enabled = True
CMDSAVE.Enabled = False
CMDCANCEL.Enabled = False
CMDEXIT.Enabled = True
CMDEDIT.Enabled = True
CMDFIRST.Enabled = True
CMDPREVIOUS.Enabled = True
CMDLAST.Enabled = True
CMDFIND.Enabled = True
CMDNEXT.Enabled = True
A = MsgBox("DO YOU WANT TO SAVE RECORD", vbYesNo, "MESSAGE")
If A = 6 Then
    If chk = 1 Then
    abus = AIRBUS_NO.Text
    Adodc1.Recordset.Update
    FRMFARE.ADD
    FRMAIRBUS.Hide
    FRMFARE.Show
    FRMFARE.AIRBUS_NO = FRMAIRBUS.AIRBUS_NO
    End If
    If chk = 2 Then
    Adodc1.Recordset.Update
    Adodc1.Recordset.MoveFirst
    End If
End If
TXT_LOCK (0)
End Sub

Public Sub CLEAR()
AIRBUS_NO.Text = ""
FIRST_CAP.Text = ""
BUS_CAP.Text = ""
ECO_CAP.Text = ""
FIRST_WL_CAP.Text = ""
BUS_WL_CAP.Text = ""
ECO_WL_CAP.Text = ""
End Sub

Private Sub ECO_CAP_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       FIRST_WL_CAP.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER NUMERIC VALUE", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
End Sub
Private Sub ECO_WL_CAP_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       CMDSAVE.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER NUMERIC VALUE", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
End Sub

Private Sub FIRST_CAP_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       BUS_CAP.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER NUMERIC VALUE", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
End Sub
Private Sub FIRST_WL_CAP_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       BUS_WL_CAP.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER NUMERIC VALUE", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
End Sub

Private Sub Form_Load()
Adodc2.Refresh
CMDCANCEL.Enabled = False
CMDSAVE.Enabled = False
TXT_LOCK (0)
End Sub

Public Sub TXT_LOCK(NO As Integer)
If NO = 0 Then
  AIRBUS_NO.Locked = True
  ECO_CAP.Locked = True
  FIRST_CAP.Locked = True
  BUS_CAP.Locked = True
  ECO_WL_CAP.Locked = True
  BUS_WL_CAP.Locked = True
  FIRST_WL_CAP.Locked = True
Else
  AIRBUS_NO.Locked = False
  ECO_CAP.Locked = False
  FIRST_CAP.Locked = False
  BUS_CAP.Locked = False
  ECO_WL_CAP.Locked = False
  BUS_WL_CAP.Locked = False
  FIRST_WL_CAP.Locked = False
End If
End Sub

