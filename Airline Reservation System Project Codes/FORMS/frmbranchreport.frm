VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form BREPORT 
   BackColor       =   &H00C0E0FF&
   Caption         =   "BRANCH REPORT"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   50
      TabIndex        =   7
      Top             =   1850
      Width           =   4815
      Begin VB.CommandButton CMDREPORT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "REPORT"
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
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CLOSE"
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
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   1200
         Top             =   240
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
         RecordSource    =   "select * from branch order by branch_code"
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
   End
   Begin MSDataListLib.DataCombo BCODE 
      Bindings        =   "frmbranchreport.frx":0000
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "BRANCH_CODE"
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker TO 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1470
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   24510467
      CurrentDate     =   37667
   End
   Begin MSComCtl2.DTPicker FROM 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   990
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   24510467
      CurrentDate     =   37667
   End
   Begin MSDataListLib.DataCombo BNAME 
      Bindings        =   "frmbranchreport.frx":0015
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CITY"
      Text            =   ""
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Branch :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   855
      TabIndex        =   6
      Top             =   660
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   735
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   495
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "BREPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BNAME_Change()
Dim s As String
If (Trim(BNAME.Text) <> "") Then
   Adodc1.Refresh
   s = "city='" & BNAME.Text & "'"
   Adodc1.Recordset.Find (s)
   BCODE.Text = Adodc1.Recordset.Fields(0)
End If
End Sub

Private Sub CMDREPORT_Click()
If Trim(BNAME.Text) = "" Then
   MsgBox "SELECT BRANCH", vbCritical + vbOKOnly, "ALERT"
   BNAME.SetFocus
   Exit Sub
End If
BRANCH_BOOK.Show
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

