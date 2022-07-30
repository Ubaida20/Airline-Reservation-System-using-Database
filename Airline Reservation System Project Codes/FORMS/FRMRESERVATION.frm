VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMRESERVATION 
   BackColor       =   &H00C0E0FF&
   Caption         =   "RESERVATION INFORMATION"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Height          =   1455
      Left            =   480
      TabIndex        =   46
      Top             =   -120
      Width           =   10695
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RESERVATION FORM"
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
         Left            =   4920
         TabIndex        =   48
         Top             =   960
         Width           =   2340
      End
      Begin VB.Image Image1 
         Height          =   1260
         Left            =   120
         Picture         =   "FRMRESERVATION.frx":0000
         Top             =   120
         Width           =   1440
      End
      Begin VB.Label Label16 
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
         Left            =   3840
         TabIndex        =   47
         Top             =   480
         Width           =   5085
      End
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   8160
      Top             =   5400
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
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
      RecordSource    =   "RESERVATION"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8160
      Top             =   5040
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "select * from BRANCH"
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
   Begin VB.Frame FRMBUTTON 
      BackColor       =   &H0080C0FF&
      Height          =   1455
      Left            =   480
      TabIndex        =   27
      Top             =   6840
      Width           =   10695
      Begin VB.CommandButton CMDPRINT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PRINT"
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
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   4560
         TabIndex        =   45
         Top             =   1440
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
         TabIndex        =   26
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   21
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "PASSENGER DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   480
      TabIndex        =   29
      Top             =   4440
      Width           =   10695
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   330
         Left            =   7680
         Top             =   1680
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
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
         Caption         =   "Adodc5"
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
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   330
         Left            =   7680
         Top             =   1320
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
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
         RecordSource    =   "SELECT * FROM RESERVATION ORDER BY PNR"
         Caption         =   "Adodc4"
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
         Height          =   330
         Left            =   7680
         Top             =   240
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
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
         RecordSource    =   "select * from reservation order by pnr"
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
      Begin VB.TextBox PASS_NM 
         Height          =   375
         Left            =   2880
         MaxLength       =   20
         TabIndex        =   11
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox PASS_ADD 
         Height          =   645
         Left            =   2880
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox PASS_NO 
         Height          =   375
         Left            =   2880
         MaxLength       =   8
         TabIndex        =   13
         Top             =   1440
         Width           =   3015
      End
      Begin VB.ComboBox PASS_STATUS 
         Height          =   315
         ItemData        =   "FRMRESERVATION.frx":0E6F
         Left            =   2880
         List            =   "FRMRESERVATION.frx":0E79
         TabIndex        =   14
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PASSENGER STATUS :"
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
         Left            =   240
         TabIndex        =   38
         Top             =   1960
         Width           =   2055
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PASSPORT NO :"
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
         Left            =   240
         TabIndex        =   37
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PASSENGER ADDRESS :"
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
         Left            =   240
         TabIndex        =   36
         Top             =   960
         Width           =   2205
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PASSENGER NAME :"
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
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   1845
      End
   End
   Begin VB.Frame FRMINFO 
      BackColor       =   &H0080C0FF&
      Caption         =   "RESERVATION DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   480
      TabIndex        =   28
      Top             =   1320
      Width           =   10695
      Begin VB.TextBox TOT_AMT 
         Height          =   375
         Left            =   7560
         MaxLength       =   6
         TabIndex        =   10
         Top             =   2160
         Width           =   2535
      End
      Begin VB.ComboBox BRANCH_CODE 
         Height          =   315
         ItemData        =   "FRMRESERVATION.frx":0E8F
         Left            =   7560
         List            =   "FRMRESERVATION.frx":0E91
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
      Begin VB.ComboBox FLIGHT_TYPE 
         Height          =   315
         ItemData        =   "FRMRESERVATION.frx":0E93
         Left            =   7560
         List            =   "FRMRESERVATION.frx":0E9D
         TabIndex        =   9
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox PNRNO 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox TOT_FARE 
         Height          =   375
         Left            =   7560
         MaxLength       =   6
         TabIndex        =   7
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox SS_CODE 
         Height          =   375
         Left            =   7560
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ComboBox CLASS 
         Height          =   315
         ItemData        =   "FRMRESERVATION.frx":0EC8
         Left            =   2160
         List            =   "FRMRESERVATION.frx":0ED5
         TabIndex        =   5
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox ORDEST 
         Height          =   375
         Left            =   2160
         MaxLength       =   7
         TabIndex        =   4
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox RESERV_DATE 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox FLIGHT_DATE 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox FLIGHT_NO 
         Height          =   405
         Left            =   2160
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL FARE :"
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
         Left            =   5040
         TabIndex        =   44
         Top             =   2280
         Width           =   1275
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FLIGHT TYPE :"
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
         Left            =   5040
         TabIndex        =   43
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PNR NO :"
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
         Left            =   240
         TabIndex        =   42
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   5040
         TabIndex        =   41
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TICKET FARE :"
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
         Left            =   5040
         TabIndex        =   40
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BRANCH CODE :"
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
         Left            =   5040
         TabIndex        =   39
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Left            =   240
         TabIndex        =   34
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Left            =   240
         TabIndex        =   33
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLASS :"
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
         Left            =   240
         TabIndex        =   32
         Top             =   2760
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RESERV. DATE :-"
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
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ORIGIN DEST : "
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
         Left            =   240
         TabIndex        =   30
         Top             =   2280
         Width           =   1410
      End
   End
End
Attribute VB_Name = "FRMRESERVATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim chk As Integer
Public Sub CLEAR()
PNRNO.Text = ""
ORDEST.Text = ""
FLIGHT_NO.Text = ""
CLASS.Text = ""
FLIGHT_TYPE.Text = ""
SS_CODE.Text = ""
FLIGHT_DATE.Text = ""
RESERV_DATE.Text = ""
TOT_FARE.Text = ""
TOT_AMT.Text = ""
PASS_NM.Text = ""
PASS_ADD.Text = ""
PASS_NO.Text = ""
PASS_STATUS.Text = ""
End Sub


Private Sub CLASS_Click()
FARE
End Sub

Private Sub CMDADD_Click()
chk = 1
CLEAR
RESERV_DATE.Text = Day(Date) & "/" & Month(Date) & "/" & Year(Date)
auto_gen
FLIGHT_DATE.SetFocus
TXT_LOCK (1)
CMDADD.Enabled = False
CMDCANCEL.Enabled = True
CMDSAVE.Enabled = True
End Sub

Private Sub CMDCANCEL_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If

CMDCANCEL.Enabled = False
CMDADD.Enabled = True
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
Text_val
End Sub
Private Sub CMDDELETE_Click()
If UCase(user) = UCase("admin") Then
A = MsgBox("DO YOU WANT TO DELETE", vbYesNo, "MESSAGE")
If A = 6 Then
  With Adodc1.Recordset
      .Delete
      .MoveNext
      If .EOF = True Then .MoveLast
  End With
  Text_val
Else
   MsgBox ("Admin User Only Delete Record")
End If
End If
End Sub
Private Sub CMDEDIT_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
chk = 2
TXT_LOCK (1)
CMDEDIT.Enabled = False
CMDSAVE.Enabled = True
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub CMDFIND_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
A = InputBox("ENTER PNR NUMBER:-")
If Trim(A) = "" Then
  Exit Sub
End If

Adodc1.Refresh
For I = 1 To Adodc1.Recordset.RecordCount
If Val(A) = Adodc1.Recordset.Fields("pnr") Then
    If Adodc1.Recordset.Fields("CANC_FLAG") = "Y" Then
  MsgBox ("ALREADY CANCELLED")
Else
  Text_val
End If
Exit Sub
End If
If Val(A) > Adodc1.Recordset.RecordCount Then
  MsgBox ("NO RECORD FOUND")
Exit Sub
End If
Adodc1.Recordset.MoveNext
Next
End Sub

Private Sub CMDFIRST_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
For I = 1 To Adodc1.Recordset.RecordCount
If Adodc1.Recordset.Fields("CANC_FLAG") = "Y" Then
Adodc1.Recordset.MoveNext
Else
Exit For
End If
Next
Text_val
End Sub

Private Sub CMDLAST_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
Adodc1.Recordset.MoveLast
Text_val
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
Exit Sub
End If
If Adodc1.Recordset.Fields("CANC_FLAG") = "Y" Then Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
Adodc1.Recordset.MoveLast
MsgBox ("POINTER IS ON LAST RECORD")
Exit Sub
End If
Text_val
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
Exit Sub
End If
If Adodc1.Recordset.Fields("CANC_FLAG") = "Y" Then Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
Adodc1.Recordset.MoveFirst
MsgBox ("POINTER IS ON FIRST RECORD")
Exit Sub
End If
Text_val
End Sub

Private Sub CMDPRINT_Click()
RESERVE.Show
End Sub

Private Sub CMDSAVE_Click()
If RESERV_DATE.Text = "" Or PNRNO.Text = "" Or FLIGHT_DATE.Text = "" Or FLIGHT_NO.Text = "" Or ORDEST.Text = "" Or CLASS.Text = "" Or BRANCH_CODE.Text = "" Or TOT_FARE.Text = "" Or _
SS_CODE.Text = "" Or FLIGHT_TYPE.Text = "" Or TOT_AMT.Text = "" Or PASS_NM.Text = "" Or PASS_ADD.Text = "" Or PASS_NO.Text = "" Or PASS_STATUS.Text = "" Then
   MsgBox ("INCOMLETE RECORD")
Exit Sub
End If
If chk = 2 Then
  CMDEDIT.Enabled = True
  GoTo X:
  End If
If CMDADD.Enabled = False Then
Adodc1.Refresh
Adodc1.Recordset.AddNew
End If
X:
A = MsgBox("DO YOU WANT TO SAVE RECORD", vbYesNo, "message")
If A = 6 Then
    If chk = 1 Then
    FLIGHT_UPDATE
    End If
FIELD_VAL
Adodc1.Recordset.Update
'BRANCH_CODE.CLEAR
TXT_LOCK (0)
End If
CMDADD.Enabled = True
CMDSAVE.Enabled = False
frmflight_list.Adodc1.Refresh
frmflight_list.DataGrid1.Refresh

End Sub
Private Sub FLIGHT_DATE_KeyDown(KeyCode As Integer, Shift As Integer)
If vbKeyF1 = KeyCode Then
frmflight_list.Show
End If
End Sub
Public Sub auto_gen()
Adodc3.Refresh
If Adodc3.Recordset.RecordCount = 0 Then
   PNRNO.Text = 1
 Exit Sub
 End If
 A = 0
For I = 1 To Adodc3.Recordset.RecordCount
If Adodc3.Recordset.Fields("pnr") > A Then
  A = Adodc3.Recordset.Fields("pnr")
   End If
 Adodc3.Recordset.MoveNext
Next
PNRNO.Text = A + 1
End Sub


Private Sub Form_Load()
Adodc2.RecordSource = "SELECT * FROM BRANCH"
Adodc1.Refresh
Adodc2.Refresh
Adodc4.Refresh
CMDSAVE.Enabled = False
CMDCANCEL.Enabled = False
BRANCHCODE
If Adodc1.Recordset.RecordCount > 0 Then
X:
If Adodc1.Recordset.EOF = True Then Exit Sub
If Adodc1.Recordset.Fields("canc_flag") = "Y" Then
Adodc1.Recordset.MoveNext
GoTo X:
Else
Text_val
End If
TXT_LOCK (0)
End If
End Sub

Public Sub BRANCHCODE()
Adodc2.RecordSource = "SELECT * FROM BRANCH"
Adodc2.Refresh
For I = 1 To Adodc2.Recordset.RecordCount
BRANCH_CODE.AddItem (Adodc2.Recordset.Fields(0) + "  " + Adodc2.Recordset.Fields(3))
Adodc2.Recordset.MoveNext
Next
End Sub

Public Sub FARE()
Adodc2.RecordSource = "select * from flight_sch"
Adodc2.Refresh
For I = 1 To Adodc2.Recordset.RecordCount
If Adodc2.Recordset.Fields(0) = FLIGHT_NO.Text Then
A = Adodc2.Recordset.Fields(1)
Exit For
End If
Adodc2.Recordset.MoveNext
Next
Adodc2.RecordSource = "select * from fare"
Adodc2.Refresh
For I = 1 To Adodc2.Recordset.RecordCount
If Adodc2.Recordset.Fields(1) = A Then
    Select Case (UCase(CLASS.Text))
        Case "FIRST CLASS":
        TOT_FARE.Text = Adodc2.Recordset.Fields(2)
        Case "BUSINESS CLASS":
        TOT_FARE.Text = Adodc2.Recordset.Fields(3)
        Case "ECONOMIC CLASS":
        TOT_FARE.Text = Adodc2.Recordset.Fields(4)
        End Select
Exit For
End If
Adodc2.Recordset.MoveNext
Next
End Sub


Private Sub PASS_ADD_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub PASS_NM_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
If KeyAscii = 32 Then GoTo A
  If KeyAscii < 65 Or KeyAscii > 90 Then
     If KeyAscii = 13 Then
       PASS_ADD.SetFocus
       Exit Sub
      End If
     If KeyAscii <> 8 Then
            I = MsgBox("ENTER ALPHABET CHARACTER", vbOKOnly, "ALERT")
       KeyAscii = 0
     End If
End If
A:
End Sub

Private Sub PASS_NO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  If KeyAscii < 48 Or KeyAscii > 57 Then
     If KeyAscii = 13 Then
        PASS_STATUS.SetFocus
        Exit Sub
     End If
     If KeyAscii <> 8 Then
        I = MsgBox("ENTER NUMERIC VALUE", vbOKOnly, "ALERT")
        KeyAscii = 0
     End If
End If
End Sub


Private Sub SS_CODE_KeyDown(KeyCode As Integer, Shift As Integer)
If vbKeyF1 = KeyCode Then
FRMSSCODE.Show
End If
End Sub

Public Sub FIELD_VAL()
Adodc1.Recordset.Fields("pnr") = PNRNO.Text
Adodc1.Recordset.Fields("flight_date") = FLIGHT_DATE.Text
Adodc1.Recordset.Fields("flight_no") = FLIGHT_NO.Text
If Trim(CLASS.Text) = "FIRST CLASS" Then
    Adodc1.Recordset.Fields("class") = "F"
        Else
        If Trim(CLASS.Text) = "BUSINESS CLASS" Then
        Adodc1.Recordset.Fields("class") = "B"
            Else
            If Trim(CLASS.Text) = "ECONOMIC CLASS" Then
            Adodc1.Recordset.Fields("class") = "E"
            End If
        End If
End If
If Trim(FLIGHT_TYPE) = "DOMISTIC FLIGHT" Then
   Adodc1.Recordset.Fields("flight_type") = "D"
Else
   Adodc1.Recordset.Fields("flight_type") = "I"
End If
Adodc1.Recordset.Fields("reserv_date") = Day(RESERV_DATE.Text) & "/" & Month(RESERV_DATE.Text) & "/" & Year(RESERV_DATE.Text)
Adodc1.Recordset.Fields("route_code") = ORDEST.Text
Adodc1.Recordset.Fields("pass_name") = PASS_NM.Text
Adodc1.Recordset.Fields("pass_add") = PASS_ADD.Text
Adodc1.Recordset.Fields("passport_no") = PASS_NO.Text
Adodc1.Recordset.Fields("ss_code") = SS_CODE.Text
Adodc1.Recordset.Fields("pass_status") = PASS_STATUS.Text
Adodc1.Recordset.Fields("total_fare") = TOT_FARE.Text
Adodc1.Recordset.Fields("branch_code") = Left(BRANCH_CODE.Text, 4)
Adodc1.Recordset.Fields("tot_fare") = TOT_AMT.Text
End Sub

Public Sub Text_val()
Adodc2.Refresh
PNRNO.Text = Adodc1.Recordset.Fields("pnr")
FLIGHT_DATE.Text = Format(Adodc1.Recordset.Fields("flight_date"), "dd/MM/yyyy")
FLIGHT_NO.Text = Adodc1.Recordset.Fields("flight_no")
If Adodc1.Recordset.Fields("class") = "F" Then
CLASS.Text = "FIRST CLASS"
    Else
        If Adodc1.Recordset.Fields("class") = "B" Then
        CLASS.Text = "BUSINESS CLASS"
        Else
            If Adodc1.Recordset.Fields("class") = "E" Then
            CLASS.Text = "ECONOMIC CLASS"
            End If
        End If
End If
If Adodc1.Recordset.Fields("flight_type") = "D" Then
   FLIGHT_TYPE.Text = "DOMISTIC FLIGHT"
Else
   FLIGHT_TYPE.Text = "INTERNATIONAL"
End If
RESERV_DATE.Text = Format(Adodc1.Recordset.Fields("reserv_date"), "dd/MM/yyyy")
ORDEST.Text = Adodc1.Recordset.Fields("route_code")
PASS_NM.Text = Adodc1.Recordset.Fields("pass_name")
PASS_ADD.Text = Adodc1.Recordset.Fields("pass_add")
PASS_NO.Text = Adodc1.Recordset.Fields("passport_no")
SS_CODE.Text = Adodc1.Recordset.Fields("ss_code")
PASS_STATUS.Text = Adodc1.Recordset.Fields("pass_status")
TOT_FARE.Text = Adodc1.Recordset.Fields("total_fare")
Adodc2.RecordSource = "SELECT * FROM BRANCH"
Adodc2.Refresh
For I = 1 To Adodc2.Recordset.RecordCount
  If Adodc1.Recordset.Fields("branch_code") = Adodc2.Recordset.Fields(0) Then
     BRANCH_CODE.Text = Adodc2.Recordset.Fields(0) + "  " + Adodc2.Recordset.Fields(3)
     Exit For
  End If
Adodc2.Recordset.MoveNext
Next
TOT_AMT.Text = Adodc1.Recordset.Fields("tot_fare")
End Sub

Public Sub TXT_LOCK(NO As Integer)
If NO = 0 Then
  PNRNO.Locked = True
  ORDEST.Locked = True
  FLIGHT_NO.Locked = True
  CLASS.Locked = True
  FLIGHT_TYPE.Locked = True
  SS_CODE.Locked = True
  FLIGHT_DATE.Locked = True
  RESERV_DATE.Locked = True
  TOT_FARE.Locked = True
  TOT_AMT.Locked = True
  PASS_NM.Locked = True
  PASS_ADD.Locked = True
  PASS_NO.Locked = True
  PASS_STATUS.Locked = True
  BRANCH_CODE.Locked = True
Else
  PNRNO.Locked = False
  ORDEST.Locked = False
  FLIGHT_NO.Locked = False
  CLASS.Locked = False
  FLIGHT_TYPE.Locked = False
  SS_CODE.Locked = False
  FLIGHT_DATE.Locked = False
  RESERV_DATE.Locked = False
  TOT_FARE.Locked = False
  TOT_AMT.Locked = False
  PASS_NM.Locked = False
  PASS_ADD.Locked = False
  PASS_NO.Locked = False
  PASS_STATUS.Locked = False
  BRANCH_CODE.Locked = False
End If
End Sub

Public Sub FLIGHT_UPDATE()
If PASS_STATUS.Text = UCase("WAITING") Then Exit Sub

Adodc5.RecordSource = "SELECT * FROM flight"
Adodc5.Refresh
Z = 0
If Adodc5.Recordset.RecordCount = 0 Then
Z = 1
GoTo Y:
Else
Z = 1
End If
dt1 = DateValue(FLIGHT_DATE.Text)
For I = 1 To Adodc5.Recordset.RecordCount
dt = DateValue(Adodc5.Recordset.Fields("FLIGHT_DATE"))

b = 0
  If UCase(FLIGHT_NO.Text) = UCase(Adodc5.Recordset.Fields("FLIGHT_NO")) And dt = dt1 Then
     
     b = 1
     Else
     b = 0
  End If
      
    If b = 1 Then
        If UCase(CLASS.Text = "FIRST CLASS") Then Adodc5.Recordset.Fields("FIRST_SEATS_BK") = Adodc5.Recordset.Fields("FIRST_SEATS_BK") + 1
        If UCase(CLASS.Text = "BUSINESS CLASS") Then Adodc5.Recordset.Fields("BUS_SEATS_BK") = Adodc5.Recordset.Fields("BUS_SEATS_BK") + 1
        If UCase(CLASS.Text = "ECONOMIC CLASS") Then Adodc5.Recordset.Fields("ECO_SEATS_BK   ") = Adodc5.Recordset.Fields("ECO_SEATS_BK") + 1
        Adodc5.Recordset.Update
        Exit Sub
    End If
    Adodc5.Recordset.MoveNext
    
    Next
Y:
    If Z = 1 Then
    Adodc5.Refresh
    Adodc5.Recordset.AddNew
    Adodc5.Recordset.Fields("FLIGHT_DATE") = FLIGHT_DATE.Text
    Adodc5.Recordset.Fields("FLIGHT_NO") = FLIGHT_NO.Text
         If UCase(CLASS.Text = "FIRST CLASS") Then
         Adodc5.Recordset.Fields("FIRST_SEATS_BK") = 1
         Adodc5.Recordset.Fields("BUS_SEATS_BK") = 0
         Adodc5.Recordset.Fields("ECO_SEATS_BK") = 0
         End If
        If UCase(CLASS.Text = "BUSINESS CLASS") Then
        Adodc5.Recordset.Fields("FIRST_SEATS_BK") = 0
         Adodc5.Recordset.Fields("BUS_SEATS_BK") = 1
         Adodc5.Recordset.Fields("ECO_SEATS_BK") = 0
         End If
        If UCase(CLASS.Text = "ECONOMIC CLASS") Then
        Adodc5.Recordset.Fields("FIRST_SEATS_BK") = 0
         Adodc5.Recordset.Fields("BUS_SEATS_BK") = 0
         Adodc5.Recordset.Fields("ECO_SEATS_BK") = 1
        End If
        Adodc5.Recordset.Update
        End If
End Sub

Public Sub cmd_enb()
CMDADD.Enabled = True
CMDSAVE.Enabled = True
CMDDELETE.Enabled = True
CMDNEXT.Enabled = True
CMDCANCEL.Enabled = True
CMDFIRST.Enabled = True
CMDPREVIOUS.Enabled = True
CMDLAST.Enabled = True
CMDEDIT.Enabled = True
CMDFIND.Enabled = True
CMDPRINT.Enabled = True

End Sub
