VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMBRANCH 
   BackColor       =   &H00C0E0FF&
   Caption         =   "BRANCH INFORMATION"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   1935
      Left            =   1680
      TabIndex        =   26
      Top             =   240
      Width           =   9615
      Begin VB.Image Image1 
         Height          =   1680
         Left            =   120
         Picture         =   "FRMBRANCH.frx":0000
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
         Left            =   3360
         TabIndex        =   28
         Top             =   480
         Width           =   5085
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BRANCH INFORMATION"
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
         Left            =   4335
         TabIndex        =   27
         Top             =   1200
         Width           =   2895
      End
   End
   Begin VB.Frame FRMBUTTON 
      BackColor       =   &H0080C0FF&
      Height          =   1455
      Left            =   1680
      TabIndex        =   17
      Top             =   6480
      Width           =   9615
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FRMINFO 
      BackColor       =   &H0080C0FF&
      Height          =   4335
      Left            =   1680
      TabIndex        =   16
      Top             =   2160
      Width           =   9615
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   375
         Left            =   6720
         Top             =   3120
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
         RecordSource    =   "SELECT * FROM BRANCH"
         Caption         =   "Adodc3"
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
      Begin VB.Frame FRM2 
         BackColor       =   &H0080C0FF&
         Height          =   1455
         Left            =   6360
         TabIndex        =   23
         Top             =   1560
         Visible         =   0   'False
         Width           =   2295
         Begin VB.OptionButton BR_CITY 
            BackColor       =   &H0080C0FF&
            Caption         =   "CITY"
            Height          =   375
            Left            =   360
            TabIndex        =   25
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton BR_CODE 
            BackColor       =   &H0080C0FF&
            Caption         =   "BRANCH CODE"
            Height          =   375
            Left            =   360
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   24
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.TextBox B_CODE 
         DataField       =   "BRANCH_CODE"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   0
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox B_ADD1 
         DataField       =   "ADD1"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox B_ADD2 
         DataField       =   "ADD2"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   2
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox B_CITY 
         DataField       =   "CITY"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   3
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox B_TELEPHONE 
         DataField       =   "TELEPHONE"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   4
         Top             =   3240
         Width           =   1935
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   6720
         Top             =   3600
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
         RecordSource    =   "BRANCH"
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
         Left            =   6720
         Top             =   480
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
         RecordSource    =   "SELECT * FROM BRANCH ORDER BY 1"
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
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "BRANCH CODE:"
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
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS1:"
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
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS2:"
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
         Top             =   2160
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CITY:"
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
         TabIndex        =   19
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "TELEPHONE NO:"
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
         TabIndex        =   18
         Top             =   3360
         Width           =   1515
      End
   End
End
Attribute VB_Name = "FRMBRANCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim chk As Integer
Private Sub B_ADD1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub B_ADD2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub B_CITY_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 65 Or KeyAscii > 90 Then
     If KeyAscii = 13 Then
       B_TELEPHONE.SetFocus
     Exit Sub
      End If
     If KeyAscii <> 8 Then
        I = MsgBox("ENTER ALPHABET CHARACTER", vbOKOnly, "ALERT")
        KeyAscii = 0
     End If
End If
End Sub

Private Sub B_TELEPHONE_KeyPress(KeyAscii As Integer)
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

Private Sub BR_CITY_Click()
If BR_CITY.Value = True Then
 A = InputBox("ENTER BRANCH CITY:-")
 If A = "" Then
 FRM2.Visible = False
  Exit Sub
  End If
 Adodc2.Refresh
Adodc2.Recordset.MoveFirst
For I = 1 To Adodc2.Recordset.RecordCount
  If UCase(A) = UCase(Adodc2.Recordset.Fields("CITY")) Then
     b = 1
Exit For
  End If
  Adodc2.Recordset.MoveNext
Next
If b = 1 Then
   B_CODE.Text = Adodc2.Recordset.Fields("BRANCH_CODE")
   B_ADD1.Text = Adodc2.Recordset.Fields("ADD1")
   B_ADD2.Text = Adodc2.Recordset.Fields("ADD2")
   B_CITY.Text = Adodc2.Recordset.Fields("CITY")
   B_TELEPHONE.Text = Adodc2.Recordset.Fields("TELEPHONE")
Else
   MsgBox ("NO RECORD FOUND")
   Adodc1.Refresh
End If
 End If
 FRM2.Visible = False
End Sub


Private Sub BR_CODE_Click()
If BR_CODE.Value = True Then
 A = InputBox("ENTER BRANCH CODE:-")
 If A = "" Then
 FRM2.Visible = False
  Exit Sub
  End If
 Adodc1.Refresh
Adodc1.Recordset.MoveFirst
For I = 1 To Adodc1.Recordset.RecordCount
  If UCase(A) = UCase(Adodc1.Recordset.Fields("BRANCH_CODE")) Then
     b = 1
Exit For
  End If
  Adodc1.Recordset.MoveNext
Next
If b = 1 Then
   B_CODE.Text = Adodc1.Recordset.Fields("BRANCH_CODE")
   B_ADD1.Text = Adodc1.Recordset.Fields("ADD1")
   B_ADD2.Text = Adodc1.Recordset.Fields("ADD2")
   B_CITY.Text = Adodc1.Recordset.Fields("CITY")
   B_TELEPHONE.Text = Adodc1.Recordset.Fields("TELEPHONE")
Else
   MsgBox ("NO RECORD FOUND")
   Adodc1.Refresh
End If
 End If
 FRM2.Visible = False
End Sub

Private Sub CMDADD_Click()
chk = 1
TXT_LOCK (1)
Adodc1.Refresh
Adodc1.Recordset.AddNew
AUTONUM
B_CODE.SetFocus
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

Public Sub TXT_LOCK(NO As Integer)
If NO = 0 Then
  B_CODE.Locked = True
  B_ADD1.Locked = True
  B_ADD2.Locked = True
  B_CITY.Locked = True
  B_TELEPHONE.Locked = True
Else
  B_CODE.Locked = False
  B_ADD1.Locked = False
  B_ADD2.Locked = False
  B_CITY.Locked = False
  B_TELEPHONE.Locked = False
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
FRM2.Visible = True
BR_CODE.Value = False
BR_CITY.Value = False
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
If chk = 2 Then
    A = MsgBox("DO YOU WANT TO SAVE RECORD", vbYesNo, "MESSAGE")
        If A = 6 Then
            Adodc1.Recordset.Update
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
            Adodc1.Refresh
            Adodc1.Recordset.MoveFirst
            TXT_LOCK (0)
            Exit Sub
        Else
        If A <> 6 Then
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
            Adodc1.Refresh
            Adodc1.Recordset.MoveFirst
            TXT_LOCK (0)
        End If
    End If
    Exit Sub
End If

If B_CODE.Text = "" Or B_ADD1.Text = "" Or B_ADD2.Text = "" Or B_CITY.Text = "" Or B_TELEPHONE.Text = "" Then
   MsgBox ("INCOMLETE RECORD")
    Exit Sub
End If

Adodc2.Refresh

For I = 1 To Adodc2.Recordset.RecordCount
If Adodc2.Recordset.Fields(0) = B_CODE.Text Then
MsgBox ("BRANCH CODE ALREADY ENTRED")
Z = 1
Exit For
End If
Adodc2.Recordset.MoveNext
Next

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
                 Adodc1.Recordset.Update
            Else
                 Adodc1.Refresh
                Adodc1.Recordset.MoveFirst
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
B_CODE.Text = ""
B_ADD1.Text = ""
B_ADD2.Text = ""
B_CITY.Text = ""
B_TELEPHONE.Text = ""
End Sub
Public Sub AUTONUM()
Dim A, b As Integer
Adodc3.Refresh
If Adodc3.Recordset.RecordCount = 0 Then
  B_CODE.Text = "A001"
  Exit Sub
  End If
 
Adodc3.RecordSource = "SELECT MAX(BRANCH_CODE) FROM BRANCH"
 Adodc3.Refresh
If IsNull(Adodc3.Recordset(0)) Then
  B_CODE.Text = "A001"
  Exit Sub
Else
  A = Val(Trim(Mid(Trim(Adodc3.Recordset(0)), 2))) + 1
  B_CODE.Text = Left(Trim(Adodc3.Recordset(0)), 4 - Len(Trim(Str(A)))) & Trim(Str(A))
End If

End Sub
