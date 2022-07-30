VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMFLIGHT_SCH 
   BackColor       =   &H00C0E0FF&
   Caption         =   "FLIGHT SCHEDULE "
   ClientHeight    =   6615
   ClientLeft      =   1575
   ClientTop       =   930
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   7650
   WindowState     =   2  'Maximized
   Begin VB.Frame FRMINFO 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   960
      TabIndex        =   23
      Top             =   1880
      Width           =   9015
      Begin VB.Frame FRM2 
         BackColor       =   &H0080C0FF&
         Height          =   1575
         Left            =   5160
         TabIndex        =   24
         Top             =   2160
         Visible         =   0   'False
         Width           =   2175
         Begin VB.OptionButton FLIG_NO 
            BackColor       =   &H0080C0FF&
            Caption         =   "FLIGHT NO"
            Height          =   375
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton AIR_NO 
            BackColor       =   &H0080C0FF&
            Caption         =   "AIRBUS NO"
            Height          =   375
            Left            =   240
            TabIndex        =   25
            Top             =   960
            Width           =   1575
         End
      End
      Begin VB.TextBox ROUTE_CODE 
         DataField       =   "ROUTE_CODE"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2160
         MaxLength       =   7
         TabIndex        =   3
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox FLI_NO 
         DataField       =   "FLIGHT_NO"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2160
         MaxLength       =   7
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox AIRBUS_NO 
         DataField       =   "AIRBUSNO"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   1
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox AIRBUS_NM 
         DataField       =   "AIRBUS_NM"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox DEPRT_TIME 
         DataField       =   "DEPRT_TIME"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   4
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox JOURNEY_HRS 
         DataField       =   "JOURNEY_HRS"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   5
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox FLI_DAY1 
         DataField       =   "FLIGHT_DAY1"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   6
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox FLI_DAY2 
         DataField       =   "FLIGHT_DAY2"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   7
         Top             =   4560
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   5040
         Top             =   4080
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
         RecordSource    =   "FLIGHT_SCH"
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
         Left            =   5280
         Top             =   720
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
         RecordSource    =   "FLIGHT_SCH"
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
      Begin VB.Label DEPRTTIME 
         BackColor       =   &H0080C0FF&
         Height          =   375
         Left            =   3960
         TabIndex        =   35
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Left            =   900
         TabIndex        =   34
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "AIRBUS NAME :"
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
         Left            =   645
         TabIndex        =   33
         Top             =   1680
         Width           =   1410
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         Left            =   705
         TabIndex        =   32
         Top             =   2280
         Width           =   1350
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "DEPARTURE TIME :"
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
         Left            =   270
         TabIndex        =   31
         Top             =   2880
         Width           =   1785
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "JOURNEY HOURS :"
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
         Left            =   330
         TabIndex        =   30
         Top             =   3480
         Width           =   1725
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "FLIGHT DAY1 :"
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
         TabIndex        =   29
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "FLIGHT DAY2 :"
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
         TabIndex        =   28
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "FLIGHT NO :"
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
         Left            =   930
         TabIndex        =   27
         Top             =   480
         Width           =   1125
      End
   End
   Begin VB.Frame FRMBUTTON 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   960
      TabIndex        =   21
      Top             =   6940
      Width           =   9015
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
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CMDSAVE 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SAVE "
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
         Left            =   3960
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
         Left            =   5400
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
         Left            =   360
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
         Left            =   1800
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
         Left            =   3240
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
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   1900
      Left            =   960
      TabIndex        =   18
      Top             =   -50
      Width           =   9015
      Begin VB.Image Image1 
         Height          =   1680
         Left            =   120
         Picture         =   "FRMFLIGHT_SCH.frx":0000
         Top             =   150
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
         Left            =   3240
         TabIndex        =   20
         Top             =   480
         Width           =   5085
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "FLIGHT SCHEDULE INFORMATION"
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
         Left            =   3600
         TabIndex        =   19
         Top             =   960
         Width           =   4065
      End
   End
End
Attribute VB_Name = "FRMFLIGHT_SCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AIR_NO_Click()
If AIR_NO.Value = True Then
 A = InputBox("ENTER AIRBUS NO:-")
  If A = "" Then
 FRM2.Visible = False
  Exit Sub
  End If
 Adodc1.Refresh
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
   FLI_NO.Text = Adodc1.Recordset.Fields("FLIGHT_NO")
   AIRBUS_NM.Text = Adodc1.Recordset.Fields("AIRBUS_NM")
   ROUTE_CODE.Text = Adodc1.Recordset.Fields("ROUTE_CODE")
   DEPRT_TIME.Text = Adodc1.Recordset.Fields("DEPRT_TIME")
   JOURNEY_HRS.Text = Adodc1.Recordset.Fields("JOURNEY_HRS")
   FLI_DAY1.Text = Adodc1.Recordset.Fields("FLIGHT_DAY1")
   FLI_DAY2.Text = Adodc1.Recordset.Fields("FLIGHT_DAY2")
Else
   MsgBox ("NO RECORD FOUND")
End If
 End If
 FRM2.Visible = False
 End Sub

Private Sub AIRBUS_NM_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
If KeyAscii = 32 Then GoTo A
  If KeyAscii < 65 Or KeyAscii > 90 Then
     If KeyAscii = 13 Then
       ROUTE_CODE.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER ALPHABET CHARACTER", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
A:
End Sub


Private Sub AIRBUS_NO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
If KeyAscii = 45 Then Exit Sub
  If KeyAscii < 48 Or KeyAscii > 57 And KeyAscii < 65 Or KeyAscii > 90 Then
     If KeyAscii = 13 Then
        AIRBUS_NM.SetFocus
        Exit Sub
     End If
     If KeyAscii <> 8 Then
        I = MsgBox("ENTER NUMERIC VALUE", vbOKOnly, "ALERT")
        KeyAscii = 0
     End If
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
CMDCANCEL.Enabled = True
CMDSAVE.Enabled = True
CMDEXIT.Enabled = True
End Sub
Private Sub CMDEXIT_Click()
FRMAIRBUS.Show
Unload Me
End Sub

Private Sub CMDFIND_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
FRM2.Visible = True
FLIG_NO.Value = False
AIR_NO.Value = False
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
On Error Resume Next
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
Adodc1.Recordset.MoveLast
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
If AIRBUS_NO.Text = "" Or FLI_NO.Text = "" Or AIRBUS_NM.Text = "" Or ROUTE_CODE.Text = "" Or DEPRT_TIME.Text = "" Or JOURNEY_HRS.Text = "" Or FLI_DAY1.Text = "" Or FLI_DAY2.Text = "" Then
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

Public Sub CLEAR()
AIRBUS_NO.Text = ""
FLI_NO.Text = ""
AIRBUS_NM.Text = ""
ROUTE_CODE.Text = ""
DEPRT_TIME.Text = ""
JOURNEY_HRS.Text = ""
FLI_DAY1.Text = ""
FLI_DAY2.Text = ""
End Sub

Private Sub DEPRT_TIME_GotFocus()
DEPRTTIME.Caption = "ENTER TIME IN 24HRS FORMAT"
End Sub

Private Sub DEPRT_TIME_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
If KeyAscii = 58 Or KeyAscii = 8 Then Exit Sub
  If KeyAscii < 48 Or KeyAscii > 57 Then
  MsgBox ("ENTER NUMRIC VALUE")
  KeyAscii = 0
       End If
End Sub

Private Sub DEPRT_TIME_LostFocus()
Dim HRS As Integer
Dim MIN As Integer
A = 0
For I = 1 To 3
  J = Mid$(DEPRT_TIME.Text, I, 1)
  If J = ":" Then
  Exit For
  Else
    A = A + 1
    
 End If
 Next
If DEPRT_TIME.Text = "" Then
  MsgBox ("ENTER HOURS BETWEEN 00 TO 23 & MINUTES IN 00 TO 59")
  DEPRT_TIME.SetFocus
Exit Sub
End If
HRS = Mid$(DEPRT_TIME.Text, 1, A)
MIN = Mid$(DEPRT_TIME.Text, 4)
If HRS > 23 Or MIN > 59 Then
  MsgBox ("ENTER HOURS BETWEEN 00 TO 23 & MINUTES IN 00 TO 59")
  FLI_DAY1.SetFocus
Else
DEPRTTIME.Caption = ""
End If
End Sub

Private Sub FLI_DAY1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       FLI_DAY2.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER NUMERIC VALUE", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
End Sub

Private Sub FLI_DAY1_LostFocus()

If Val(FLI_DAY1.Text) > 7 Then
  MsgBox ("ENTER DAY BETWEEN 1 TO 7")
  FLI_DAY1.Text = ""
  FLI_DAY1.SetFocus
End If
End Sub

Private Sub FLI_DAY2_KeyPress(KeyAscii As Integer)
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

Private Sub FLI_DAY2_LostFocus()
If Val(FLI_DAY2.Text) > 7 Then
  MsgBox ("ENTER DAY BETWEEN 1 TO 7")
  FLI_DAY2.Text = ""
  FLI_DAY2.SetFocus
End If
If FLI_DAY1.Text = FLI_DAY2.Text Then
  MsgBox ("ENTER ANOTHER DAY")
  FLI_DAY2.SetFocus
End If
  
End Sub

Private Sub FLI_NO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
If KeyAscii = 45 Then Exit Sub
  If KeyAscii < 48 Or KeyAscii > 57 And KeyAscii < 65 Or KeyAscii > 90 Then
     If KeyAscii = 13 Then
        AIRBUS_NO.SetFocus
        Exit Sub
     End If
     If KeyAscii <> 8 Then
        I = MsgBox("ENTER NUMERIC VALUE", vbOKOnly, "ALERT")
        KeyAscii = 0
     End If
End If
End Sub

Private Sub FLIG_NO_Click()
If FLIG_NO.Value = True Then
 A = InputBox("ENTER FLIGHT NO:-")
 If A = "" Then
 FRM2.Visible = False
  Exit Sub
  End If
 Adodc2.Refresh
Adodc2.Recordset.MoveFirst
For I = 1 To Adodc2.Recordset.RecordCount
  If UCase(A) = UCase(Adodc2.Recordset.Fields("FLIGHT_NO")) Then
     b = 1
Exit For
  End If
  Adodc2.Recordset.MoveNext
Next
If b = 1 Then
   AIRBUS_NO.Text = Adodc2.Recordset.Fields("AIRBUSNO")
   FLI_NO.Text = Adodc2.Recordset.Fields("FLIGHT_NO")
   AIRBUS_NM.Text = Adodc2.Recordset.Fields("AIRBUS_NM")
   ROUTE_CODE.Text = Adodc2.Recordset.Fields("ROUTE_CODE")
   DEPRT_TIME.Text = Adodc2.Recordset.Fields("DEPRT_TIME")
   JOURNEY_HRS.Text = Adodc2.Recordset.Fields("JOURNEY_HRS")
   FLI_DAY1.Text = Adodc2.Recordset.Fields("FLIGHT_DAY1")
   FLI_DAY2.Text = Adodc2.Recordset.Fields("FLIGHT_DAY2")
Else
   MsgBox ("NO RECORD FOUND")
End If
 End If
 FRM2.Visible = False
End Sub

Private Sub Form_Load()
Adodc1.Refresh
CMDCANCEL.Enabled = False
CMDSAVE.Enabled = False
TXT_LOCK (0)
End Sub


Private Sub JOURNEY_HRS_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
       JOURNEY_HRS.SetFocus
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
 FLI_NO.Locked = True
 AIRBUS_NO.Locked = True
 AIRBUS_NM.Locked = True
 ROUTE_CODE.Locked = True
 DEPRT_TIME.Locked = True
 JOURNEY_HRS.Locked = True
 FLI_DAY1.Locked = True
 FLI_DAY2.Locked = True
Else
 FLI_NO.Locked = False
 AIRBUS_NO.Locked = False
 AIRBUS_NM.Locked = False
 ROUTE_CODE.Locked = False
 DEPRT_TIME.Locked = False
 JOURNEY_HRS.Locked = False
 FLI_DAY1.Locked = False
 FLI_DAY2.Locked = False
End If
End Sub

Public Sub ADD()
TXT_LOCK (1)
Adodc1.Refresh
Adodc1.Recordset.AddNew
'FLI_NO.SetFocus
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

