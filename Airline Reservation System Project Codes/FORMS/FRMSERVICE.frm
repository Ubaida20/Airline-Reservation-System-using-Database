VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMSERVICE 
   BackColor       =   &H00C0E0FF&
   Caption         =   "SERVICE INFORMATION"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   1935
      Left            =   840
      TabIndex        =   19
      Top             =   600
      Width           =   9015
      Begin VB.Image Image1 
         Height          =   1680
         Left            =   120
         Picture         =   "FRMSERVICE.frx":0000
         Top             =   120
         Width           =   1920
      End
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
         Left            =   3480
         TabIndex        =   21
         Top             =   480
         Width           =   5085
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "SERVICE INFORMATION"
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
         TabIndex        =   20
         Top             =   1080
         Width           =   2895
      End
   End
   Begin VB.Frame FRMBUTTON 
      BackColor       =   &H0080C0FF&
      Height          =   1455
      Left            =   840
      TabIndex        =   15
      Top             =   5400
      Width           =   9015
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
         TabIndex        =   3
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
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
         TabIndex        =   12
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
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FRMINFO 
      BackColor       =   &H0080C0FF&
      Height          =   2895
      Left            =   840
      TabIndex        =   14
      Top             =   2520
      Width           =   9015
      Begin VB.TextBox S_CODE 
         BackColor       =   &H00FFFFFF&
         DataField       =   "SS_CODE"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   0
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox S_DESC 
         DataField       =   "SS_DESC"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2760
         MaxLength       =   15
         TabIndex        =   1
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox S_FARE 
         DataField       =   "SS_FARE"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2760
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1920
         Width           =   1815
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   5520
         Top             =   1440
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
         UserName        =   ""
         Password        =   "TYBCA34"
         RecordSource    =   "SERVICE"
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
         Caption         =   "SERVICE CODE :"
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
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   "SERVICE DESCRIPTION :"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   "SERVICE FARE :"
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
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FRMSERVICE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDADD_Click()
TXT_LOCK (1)
Adodc1.Refresh
Adodc1.Recordset.AddNew
S_CODE.SetFocus
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
CMDCANCEL.Enabled = False
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
Adodc1.Recordset.Update
CMDADD.Enabled = False
CMDEDIT.Enabled = False
CMDDELETE.Enabled = False
CMDFIRST.Enabled = False
CMDNEXT.Enabled = False
CMDPREVIOUS.Enabled = False
CMDLAST.Enabled = False
CMDFIND.Enabled = False
CMDSAVE.Enabled = True
CMDCANCEL.Enabled = True
CMDEXIT.Enabled = True
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub CMDFIND_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
A = InputBox("ENTER SERVICE CODE:")
If Trim(A) = "" Then
  Exit Sub
End If

A = UCase(A)
 Adodc1.Refresh
 Adodc1.Recordset.MoveFirst
 For I = 1 To Adodc1.Recordset.RecordCount
   If UCase(A) = UCase(Adodc1.Recordset.Fields("SS_CODE")) Then
     b = 1
 Exit For
   End If
   Adodc1.Recordset.MoveNext
 Next
 If b = 1 Then
   S_CODE.Text = Adodc1.Recordset.Fields("SS_CODE")
   S_DESC.Text = Adodc1.Recordset.Fields("SS_DESC")
   S_FARE.Text = Adodc1.Recordset.Fields("SS_FARE")
Else
   MsgBox ("NO RECORD FOUND")
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
   MsgBox ("POINTER IS ON FIRST RECORD")
End If
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

Private Sub CMDSAVE_Click()
If S_CODE.Text = "" Or S_DESC.Text = "" Or S_FARE.Text = "" Then
  MsgBox ("INCOMPLETE RECORD")
  Exit Sub
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
   Adodc1.Recordset.Save
Else
   CLEAR
End If
TXT_LOCK (0)
End Sub

Private Sub Form_Load()
Adodc1.Refresh
CMDCANCEL.Enabled = False
CMDSAVE.Enabled = False
TXT_LOCK (0)
End Sub

Public Sub CLEAR()
S_CODE.Text = ""
S_DESC.Text = ""
S_FARE.Text = ""
End Sub



Private Sub S_CODE_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
If KeyAscii = 45 Then GoTo A
  If KeyAscii < 65 Or KeyAscii > 90 Then
         If KeyAscii = 13 Then
           S_DESC.SetFocus
         Exit Sub
  End If
  If KeyAscii <> 8 Then
     I = MsgBox("ENTER ALPHABET CHARACTER", vbOKOnly, "ALERT")
     KeyAscii = 0
  End If
End If
A:
End Sub

Private Sub S_DESC_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
If KeyAscii = 32 Then GoTo A
  If KeyAscii < 65 Or KeyAscii > 90 Then
     If KeyAscii = 13 Then
      S_FARE.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER ALPHABET CHARACTER", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
A:
End Sub

Private Sub S_FARE_KeyPress(KeyAscii As Integer)
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

Public Sub TXT_LOCK(NO As Integer)
If NO = 0 Then
  S_CODE.Locked = True
  S_DESC.Locked = True
  S_FARE.Locked = True
Else
  S_CODE.Locked = False
  S_DESC.Locked = False
  S_FARE.Locked = False
End If
End Sub
