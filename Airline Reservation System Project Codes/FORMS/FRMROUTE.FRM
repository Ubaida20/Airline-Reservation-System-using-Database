VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMROUTE 
   BackColor       =   &H00C0E0FF&
   Caption         =   "ROUTE INFORMATION"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   1935
      Left            =   1440
      TabIndex        =   21
      Top             =   720
      Width           =   8775
      Begin VB.Image Image1 
         Height          =   1680
         Left            =   120
         Picture         =   "FRMROUTE.frx":0000
         Top             =   120
         Width           =   1920
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
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
         Left            =   2760
         TabIndex        =   23
         Top             =   360
         Width           =   5070
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ROUTE INFORMATION"
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
         Left            =   3735
         TabIndex        =   22
         Top             =   960
         Width           =   2655
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2880
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "ROUTE"
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
   Begin VB.Frame BUTTON 
      BackColor       =   &H0080C0FF&
      Height          =   1455
      Left            =   1440
      TabIndex        =   20
      Top             =   5640
      Width           =   8775
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
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   840
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FRMROUTE 
      BackColor       =   &H0080C0FF&
      Height          =   3015
      Left            =   1440
      TabIndex        =   15
      Top             =   2640
      Width           =   8775
      Begin VB.TextBox DEST 
         DataField       =   "DESTINATION"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   3
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox ORIGIN 
         DataField       =   "ORIGIN"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox ROUTE_DESC 
         DataField       =   "ROUTE_DESC"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3960
         MaxLength       =   25
         TabIndex        =   1
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox ROUTE_CODE 
         DataField       =   "ROUTE_CODE"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3960
         MaxLength       =   7
         TabIndex        =   0
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESTINATION : "
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
         Left            =   1200
         TabIndex        =   19
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ORIGIN :"
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
         Left            =   1200
         TabIndex        =   18
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ROUTE DESCRIPTION : "
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
         Left            =   1200
         TabIndex        =   17
         Top             =   1080
         Width           =   2160
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Left            =   1200
         TabIndex        =   16
         Top             =   480
         Width           =   1350
      End
   End
End
Attribute VB_Name = "FRMROUTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDADD_Click()
TXT_LOCK (1)
Adodc1.Refresh
Adodc1.Recordset.AddNew
ROUTE_CODE.SetFocus
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
A = InputBox("ENTER ROUTE CODE:")
If Trim(A) = "" Then
  Exit Sub
End If

A = UCase(A)
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
   ROUTE_CODE.Text = Adodc1.Recordset.Fields("ROUTE_CODE")
   ROUTE_DESC.Text = Adodc1.Recordset.Fields("ROUTE_DESC")
   ORIGIN.Text = Adodc1.Recordset.Fields("ORIGIN")
   DEST.Text = Adodc1.Recordset.Fields("DESTINATION")
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
If ROUTE_CODE.Text = "" Or ROUTE_DESC.Text = "" Or ORIGIN.Text = "" Or DEST.Text = "" Then
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

Public Sub TXT_LOCK(NO As Integer)
If NO = 0 Then
  ROUTE_CODE.Locked = True
  ROUTE_DESC.Locked = True
  ORIGIN.Locked = True
  DEST.Locked = True
Else
  ROUTE_CODE.Locked = False
  ROUTE_DESC.Locked = False
  ORIGIN.Locked = False
  DEST.Locked = False
End If

End Sub


Private Sub DEST_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
If KeyAscii = 32 Then GoTo A
  If KeyAscii < 65 Or KeyAscii > 90 Then
     If KeyAscii = 13 Then
       CMDSAVE.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER ALPHABET CHARACTER", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
A:
End Sub

Private Sub Form_Load()
Adodc1.Refresh
CMDCANCEL.Enabled = False
CMDSAVE.Enabled = False
TXT_LOCK (0)
End Sub

Public Sub CLEAR()
ROUTE_CODE.Text = ""
ROUTE_DESC.Text = ""
ORIGIN.Text = ""
DEST.Text = ""
End Sub

Private Sub ORIGIN_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
If KeyAscii = 32 Then GoTo A
  If KeyAscii < 65 Or KeyAscii > 90 Then
     If KeyAscii = 13 Then
       DEST.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER ALPHABET CHARACTER", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
A:
End Sub

Private Sub ROUTE_CODE_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
If KeyAscii = 45 Then Exit Sub
  If KeyAscii < 65 Or KeyAscii > 90 Then
     If KeyAscii = 13 Then
        ROUTE_DESC.SetFocus
        Exit Sub
     End If
     If KeyAscii <> 8 Then
        I = MsgBox("ENTER ANY VALUE", vbOKOnly, "ALERT")
        KeyAscii = 0
     End If
End If
End Sub

Private Sub ROUTE_DESC_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
If KeyAscii = 32 Then GoTo A
  If KeyAscii < 65 Or KeyAscii > 90 Then
     If KeyAscii = 13 Then
       ORIGIN.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER ALPHABET CHARACTER", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
A:
End Sub
