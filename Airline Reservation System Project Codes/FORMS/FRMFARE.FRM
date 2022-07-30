VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMFARE 
   BackColor       =   &H00C0E0FF&
   Caption         =   "FARE INFORMATION"
   ClientHeight    =   7440
   ClientLeft      =   900
   ClientTop       =   675
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7543.727
   ScaleMode       =   0  'User
   ScaleWidth      =   15072.5
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   1935
      Left            =   1200
      TabIndex        =   22
      Top             =   480
      Width           =   9135
      Begin VB.Image Image1 
         Height          =   1680
         Left            =   120
         Picture         =   "FRMFARE.frx":0000
         Top             =   120
         Width           =   1920
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "FARE INFORMATION"
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
         Left            =   4080
         TabIndex        =   24
         Top             =   1320
         Width           =   2475
      End
      Begin VB.Label Label2 
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
         Left            =   2640
         TabIndex        =   23
         Top             =   600
         Width           =   5085
      End
   End
   Begin VB.Frame FRMBUTTON 
      BackColor       =   &H0080C0FF&
      Height          =   1575
      Left            =   1200
      TabIndex        =   5
      Top             =   5760
      Width           =   9135
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
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
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
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   1215
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
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   960
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   960
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   960
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
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   960
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
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
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame FRMINFO 
      BackColor       =   &H0080C0FF&
      Height          =   3375
      Left            =   1200
      TabIndex        =   16
      Top             =   2400
      Width           =   9135
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   7080
         Top             =   2160
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         RecordSource    =   ""
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
      Begin VB.ComboBox route_code 
         DataField       =   "ROUTE_CODE"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   3000
         TabIndex        =   1
         Top             =   1080
         Width           =   2415
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   7080
         Top             =   2520
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         RecordSource    =   "FARE"
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
      Begin VB.TextBox AIRBUS_NO 
         DataField       =   "AIRBUSNO"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3000
         MaxLength       =   5
         TabIndex        =   0
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox F_FARE 
         DataField       =   "FIRST_FARE"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox B_FARE 
         DataField       =   "BUS_FARE"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   3
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox E_FARE 
         DataField       =   "ECO_FARE"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   4
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ROUTE CODE :"
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
         Left            =   600
         TabIndex        =   21
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "FIRST FARE :"
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
         Left            =   600
         TabIndex        =   20
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "BUSINESS FARE :"
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
         Left            =   600
         TabIndex        =   19
         Top             =   2280
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ECONOMICAL FARE :"
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
         Left            =   600
         TabIndex        =   18
         Top             =   2880
         Width           =   1875
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "AIRBUS NO : "
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
         Left            =   600
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FRMFARE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ROUTE As String
Dim AIRBUS As String
Private Sub B_FARE_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       E_FARE.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER NUMERIC VALUE ONLY", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
End Sub

Private Sub B_FARE_LostFocus()
If Val(B_FARE.Text) >= Val(F_FARE.Text) Then
MsgBox ("BUSINESS FARE NOT GREATER THAN FIRST FARE")
B_FARE.SetFocus
End If
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
CMDDELETE.Enabled = True
CMDSAVE.Enabled = False
CMDCANCEL.Enabled = False
CMDEXIT.Enabled = True
CMDEDIT.Enabled = True
CMDFIRST.Enabled = True
CMDPREVIOUS.Enabled = True
CMDLAST.Enabled = True
CMDFIND.Enabled = True
CMDNEXT.Enabled = True
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
TXT_LOCK (1)
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
Adodc1.Recordset.Update
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub CMDFIND_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If

 A = InputBox("ENTER ROUTE CODE:-")
 If Trim(A) = "" Then
  Exit Sub
End If

  Adodc1.Refresh
Adodc1.Recordset.MoveFirst
For I = 1 To Adodc1.Recordset.RecordCount
  If UCase(A) = UCase(Adodc1.Recordset.Fields("ROUTE_CODE")) Then
     b = 1
Exit For
  End If
  Adodc1.Recordset.MoveNext
Next
If b = 1 Then
   F_FARE.Text = Adodc1.Recordset.Fields("FIRST_FARE")
   B_FARE.Text = Adodc1.Recordset.Fields("BUS_FARE")
   E_FARE.Text = Adodc1.Recordset.Fields("ECO_FARE")
  Else
   MsgBox ("NO RECORD FOUND")
End If

End Sub

Private Sub CMDFIRST_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
CMDDELETE.Enabled = True
CMDCANCEL.Enabled = False
CMDSAVE.Enabled = False
CMDEXIT.Enabled = True
CMDEDIT.Enabled = True
CMDFIRST.Enabled = True
CMDPREVIOUS.Enabled = True
CMDLAST.Enabled = True
CMDFIND.Enabled = True
CMDNEXT.Enabled = True
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
End Sub

Private Sub CMDLAST_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
CMDDELETE.Enabled = True
CMDSAVE.Enabled = False
CMDCANCEL.Enabled = False
CMDEXIT.Enabled = True
CMDEDIT.Enabled = True
CMDFIRST.Enabled = True
CMDPREVIOUS.Enabled = True
CMDLAST.Enabled = True
CMDFIND.Enabled = True
CMDNEXT.Enabled = True
Adodc1.Recordset.MoveLast
End Sub

Private Sub CMDNEXT_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
CMDDELETE.Enabled = True
CMDCANCEL.Enabled = False
CMDSAVE.Enabled = False
CMDEXIT.Enabled = True
CMDEDIT.Enabled = True
CMDFIRST.Enabled = True
CMDPREVIOUS.Enabled = True
CMDLAST.Enabled = True
CMDFIND.Enabled = True
CMDNEXT.Enabled = True
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
 Adodc1.Recordset.MoveLast
 MsgBox ("POINTER IS ON LAST RECORD")
End If
End Sub

Private Sub CMDPREVIOUS_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
CMDDELETE.Enabled = True
CMDCANCEL.Enabled = False
CMDSAVE.Enabled = False
CMDEXIT.Enabled = True
CMDEDIT.Enabled = True
CMDFIRST.Enabled = True
CMDPREVIOUS.Enabled = True
CMDLAST.Enabled = True
CMDFIND.Enabled = True
CMDNEXT.Enabled = True
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
 Adodc1.Recordset.MoveFirst
 MsgBox ("POINTER IS ON FIRST RECORD")
End If
End Sub

Private Sub CMDSAVE_Click()
If F_FARE.Text = "" Or B_FARE.Text = "" Or E_FARE.Text = "" Then
   MsgBox ("INCOMPLETE RECORD")
Exit Sub
End If
CMDDELETE.Enabled = True
CMDCANCEL.Enabled = False
CMDSAVE.Enabled = False
CMDEXIT.Enabled = True
CMDEDIT.Enabled = True
CMDFIRST.Enabled = True
CMDPREVIOUS.Enabled = True
CMDLAST.Enabled = True
CMDFIND.Enabled = True
CMDNEXT.Enabled = True
A = MsgBox("DO YOU WANT TO SAVE RECORD?", vbYesNo, "MESSAGE")
If A = 6 Then
    ROUTE = ROUTE_CODE.Text
    AIRBUS = AIRBUS_NO.Text
    Adodc1.Recordset.Update
    FRMFARE.Hide
    FRMFLIGHT_SCH.Show
    FRMFLIGHT_SCH.ADD
    FRMFLIGHT_SCH.AIRBUS_NO = FRMFARE.AIRBUS_NO
    FRMFLIGHT_SCH.ROUTE_CODE = FRMFARE.ROUTE_CODE
Else
    CLEAR
End If
TXT_LOCK (0)
End Sub

Public Sub CLEAR()
F_FARE.Text = ""
B_FARE.Text = ""
E_FARE.Text = ""
End Sub

Private Sub E_FARE_KeyPress(KeyAscii As Integer)
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

Private Sub E_FARE_LostFocus()
If Val(E_FARE.Text) >= Val(B_FARE.Text) And Val(E_FARE.Text) >= Val(F_FARE) Then
  MsgBox ("ECONOMICAL FARE NOT GREATER THAN FIRST FARE AND BUSINESS FARE")
  E_FARE.SetFocus
End If
End Sub


Private Sub F_FARE_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       B_FARE.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER NUMERIC VALUE ONLY", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
End Sub

Private Sub Form_Load()
ROUT
Adodc1.Refresh
CMDCANCEL.Enabled = False
CMDSAVE.Enabled = False
TXT_LOCK (0)
End Sub


Public Sub TXT_LOCK(NO As Integer)
If NO = 0 Then
  F_FARE.Locked = True
  B_FARE.Locked = True
  E_FARE.Locked = True
Else
  F_FARE.Locked = False
  B_FARE.Locked = False
  E_FARE.Locked = False
End If
End Sub

Public Sub ADD()

TXT_LOCK (1)
Adodc1.Refresh
Adodc1.Recordset.AddNew
CMDEDIT.Enabled = False
CMDDELETE.Enabled = False
CMDFIRST.Enabled = False
CMDNEXT.Enabled = False
CMDLAST.Enabled = False
CMDPREVIOUS.Enabled = False
CMDFIND.Enabled = False
CMDCANCEL.Enabled = True
CMDSAVE.Enabled = True
CMDEXIT.Enabled = True
End Sub

Public Sub ROUT()
Adodc2.RecordSource = "select * from route"
Adodc2.Refresh
For I = 1 To Adodc2.Recordset.RecordCount
ROUTE_CODE.AddItem (Adodc2.Recordset.Fields(0))
Adodc2.Recordset.MoveNext
Next
End Sub

Private Sub route_code_DblClick()
ROUT
End Sub
