VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmflight_list 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Flight List"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDREFRESH 
      BackColor       =   &H00E0E0E0&
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   1080
      Top             =   7080
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
   Begin VB.CommandButton check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CLICK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   2280
      Top             =   7080
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
      RecordSource    =   "SELECT * FROM FLIGHT_SCH"
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
   Begin VB.ComboBox ROUTE 
      Height          =   315
      ItemData        =   "frmflight_list.frx":0000
      Left            =   2280
      List            =   "frmflight_list.frx":0002
      TabIndex        =   3
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox FLDATE 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin MSComCtl2.MonthView mv 
      Height          =   2370
      Left            =   7320
      TabIndex        =   1
      Top             =   480
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   24641537
      CurrentDate     =   37615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3600
      Top             =   7080
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmflight_list.frx":0004
      Height          =   1935
      Left            =   840
      TabIndex        =   0
      Top             =   3480
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmflight_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim chk As Integer
Private Sub check_Click()
Adodc2.RecordSource = "SELECT * FROM FLIGHT_SCH WHERE ROUTE_CODE ='" & Trim(ROUTE.Text) & "' and (flight_day1 = " & mv.DayOfWeek & " or  flight_day2= " & mv.DayOfWeek & ")"
Adodc2.Refresh
Adodc2.Refresh
chk = chk + 1
If chk < 2 Then
check_Click
End If

End Sub

Private Sub CMDREFRESH_Click()
Adodc2.RecordSource = "SELECT * FROM FLIGHT_SCH "
Adodc2.Refresh
FLDATE.Text = ""
ROUTE.Text = ""
End Sub

Private Sub DataGrid1_Click()
FRMRESERVATION.Show
FRMRESERVATION.FLIGHT_DATE = frmflight_list.FLDATE.Text
FRMRESERVATION.FLIGHT_NO = frmflight_list.Adodc2.Recordset.Fields(0)
FRMRESERVATION.ORDEST.Text = frmflight_list.Adodc2.Recordset.Fields(3)
frmflight_list.Hide
End Sub
Private Sub Form_Load()
ROU
Adodc2.Refresh
End Sub
Public Sub ROUT()
Adodc2.Refresh
For I = 0 To Adodc2.Recordset.RecordCount
A = 1
For J = 0 To ROUTE.ListCount
If Adodc2.Recordset.Fields(3) = ROUTE.List(J) Then
A = 0
Exit For
End If
Next
If A = 0 Then
Adodc2.Recordset.MoveNext
Else
ROUTE.AddItem (Adodc2.Recordset.Fields(3))
End If
Next
End Sub

Private Sub MV_DateClick(ByVal DateClicked As Date)
dt = mv.Month & "/" & mv.Day & "/" & mv.Year
FLDATE.Text = Format(dt, "dd/MM/yyyy")
End Sub

Public Sub ROU()
Adodc3.RecordSource = "SELECT * FROM ROUTE"
Adodc3.Refresh
For I = 1 To Adodc3.Recordset.RecordCount

ROUTE.AddItem (Adodc3.Recordset.Fields("ROUTE_CODE"))
Adodc3.Recordset.MoveNext
Next

End Sub
