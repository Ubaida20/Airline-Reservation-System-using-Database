VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMFLIGHT 
   BackColor       =   &H00C0E0FF&
   Caption         =   "FLIGHT INFORMATION"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   2055
      Left            =   960
      TabIndex        =   19
      Top             =   480
      Width           =   9135
      Begin VB.Image Image1 
         Height          =   1680
         Left            =   120
         Picture         =   "FRMFLIGHT.frx":0000
         Top             =   240
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
         Left            =   2880
         TabIndex        =   21
         Top             =   360
         Width           =   5085
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "FLIGHT INFORMATION"
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
         Left            =   3840
         TabIndex        =   20
         Top             =   1080
         Width           =   2745
      End
   End
   Begin VB.Frame FRMBUTTON 
      BackColor       =   &H0080C0FF&
      Height          =   1095
      Left            =   960
      TabIndex        =   12
      Top             =   6840
      Width           =   9135
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
         TabIndex        =   5
         Top             =   360
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
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
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
         TabIndex        =   8
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
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
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
      Height          =   4335
      Left            =   960
      TabIndex        =   10
      Top             =   2520
      Width           =   9135
      Begin VB.TextBox F_NO 
         DataField       =   "FLIGHT_NO"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         MaxLength       =   7
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox F_DATE 
         DataField       =   "FLIGHT_DATE"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox F_BK_C 
         DataField       =   "FIRST_SEATS_BK"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   2
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox B_BK_C 
         DataField       =   "BUS_SEATS_BK"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   3
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox E_BK_C 
         DataField       =   "ECO_SEATS_BK"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   4
         Top             =   3720
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   5400
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
         RecordSource    =   "FLIGHT"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label2 
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
         Left            =   360
         TabIndex        =   18
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "FLIGHT DATE :"
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
         Left            =   360
         TabIndex        =   17
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "BOOKING SEATS :"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   1635
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "FIRST CLASS :"
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
         Left            =   360
         TabIndex        =   15
         Top             =   2640
         Width           =   1320
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "BUSINESS CLASS :"
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
         Left            =   360
         TabIndex        =   14
         Top             =   3240
         Width           =   1710
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ECONOMICAL CLASS :"
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
         Left            =   360
         TabIndex        =   13
         Top             =   3840
         Width           =   1980
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "FRMFLIGHT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDCHECK_Click()
Adodc1.RecordSource = "SELECT * FROM RESERVATION"
Adodc1.Refresh
A = F_NO.Text
dt = DateValue(Adodc1.Recordset.Fields("FLIGHT_DATE"))
dt1 = DateValue(F_DATE.Text)
For I = 1 To Adodc1.Recordset.RecordCount
b = 0
  If UCase(A) = UCase(Adodc1.Recordset.Fields("FLIGHT_NO")) And dt = dt1 Then
     b = 1
     Else
     b = 0
  End If
    If b = 1 Then
            If Adodc1.Recordset.Fields("CLASS") = "F" Then F = F + 1
            If Adodc1.Recordset.Fields("CLASS") = "B" Then BC = BC + 1
            If Adodc1.Recordset.Fields("CLASS") = "E" Then e = e + 1
    End If
    Adodc1.Recordset.MoveNext
    Next
F_BK_C.Text = F
B_BK_C.Text = BC
E_BK_C.Text = e

End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub MV_DateClick(ByVal DateClicked As Date)
dt = mv.Day & "/" & mv.Month & "/" & mv.Year
F_DATE.Text = dt
End Sub

Private Sub CMDFIRST_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
Adodc1.Recordset.MoveFirst
End Sub

Private Sub CMDLAST_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
Adodc1.Recordset.MoveLast
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
End Sub

