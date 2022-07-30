VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMCANC 
   BackColor       =   &H00C0E0FF&
   Caption         =   "CANCELLATION INFORMATION"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   975
      Left            =   360
      TabIndex        =   48
      Top             =   0
      Width           =   11055
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANCELLATION FORM"
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
         Left            =   4680
         TabIndex        =   50
         Top             =   600
         Width           =   2685
      End
      Begin VB.Image Image1 
         Height          =   765
         Left            =   120
         Picture         =   "FRMCANCELLATION.frx":0000
         Top             =   120
         Width           =   870
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
         Left            =   3600
         TabIndex        =   49
         Top             =   240
         Width           =   5085
      End
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   8160
      Top             =   5880
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   "select * from cancellation order by pnr"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Height          =   1455
      Left            =   360
      TabIndex        =   30
      Top             =   6960
      Width           =   11055
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
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Left            =   1680
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
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
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   22
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
         TabIndex        =   23
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   5520
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
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   27
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
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame PASSINFO 
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
      Left            =   360
      TabIndex        =   29
      Top             =   4560
      Width           =   11055
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   375
         Left            =   7800
         Top             =   1560
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
         RecordSource    =   ""
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   7800
         Top             =   960
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
      Begin VB.ComboBox PASS_STATUS 
         Height          =   315
         Left            =   2880
         TabIndex        =   16
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox PASS_NM 
         Height          =   375
         Left            =   2880
         MaxLength       =   20
         TabIndex        =   13
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox PASS_ADD 
         Height          =   615
         Left            =   2880
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox PASS_NO 
         Height          =   375
         Left            =   2880
         MaxLength       =   8
         TabIndex        =   15
         Top             =   1440
         Width           =   4095
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   7800
         Top             =   600
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
         RecordSource    =   ""
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
      Begin VB.Label Label12 
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
         TabIndex        =   39
         Top             =   360
         Width           =   1845
      End
      Begin VB.Label Label13 
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
         TabIndex        =   38
         Top             =   960
         Width           =   2205
      End
      Begin VB.Label Label14 
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
      Begin VB.Label Label15 
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
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   2055
      End
   End
   Begin VB.Frame FRMCANCEL 
      BackColor       =   &H0080C0FF&
      Caption         =   "CANCELLATION DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   360
      TabIndex        =   25
      Top             =   960
      Width           =   11055
      Begin VB.TextBox CANCEL_DATE 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   405
         Left            =   2520
         TabIndex        =   1
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox can_amt 
         Height          =   405
         Left            =   7920
         TabIndex        =   12
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox TOT_AMT 
         Height          =   405
         Left            =   7920
         TabIndex        =   11
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox PNRNO 
         Height          =   375
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1200
         Width           =   2055
      End
      Begin VB.ComboBox FLIGHT_TYPE 
         Height          =   315
         Left            =   7920
         TabIndex        =   10
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox SS_CODE 
         Height          =   375
         Left            =   7920
         TabIndex        =   9
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox BRANCH_CODE 
         Height          =   375
         Left            =   7920
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox TOT_FARE 
         Height          =   375
         Left            =   7920
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox ORIDEST 
         Height          =   375
         Left            =   2520
         MaxLength       =   7
         TabIndex        =   5
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox FLIGHT_NO 
         Height          =   375
         Left            =   2520
         MaxLength       =   7
         TabIndex        =   4
         Top             =   2160
         Width           =   2055
      End
      Begin VB.ComboBox CLASS 
         Height          =   315
         ItemData        =   "FRMCANCELLATION.frx":0815
         Left            =   2520
         List            =   "FRMCANCELLATION.frx":0817
         TabIndex        =   6
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox FLIGHT_DATE 
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox RESERV_DATE 
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
         Left            =   2520
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANCELLATION DATE :"
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
         TabIndex        =   47
         Top             =   840
         Width           =   2085
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CANCEL AMOUNT"
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
         Index           =   1
         Left            =   5280
         TabIndex        =   46
         Top             =   2760
         Width           =   1590
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   5280
         TabIndex        =   45
         Top             =   2280
         Width           =   1275
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
         TabIndex        =   44
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label6 
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
         Left            =   5280
         TabIndex        =   43
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label7 
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
         Left            =   5280
         TabIndex        =   42
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label Label10 
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
         Left            =   5280
         TabIndex        =   41
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label11 
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
         Left            =   5280
         TabIndex        =   40
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ORIGIN DEST. :"
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
         Top             =   2760
         Width           =   1440
      End
      Begin VB.Label Label4 
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
         Top             =   2280
         Width           =   1320
      End
      Begin VB.Label Label5 
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
         TabIndex        =   33
         Top             =   3135
         Width           =   960
      End
      Begin VB.Label Label8 
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
         TabIndex        =   32
         Top             =   1800
         Width           =   1440
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "RESERV. DATE :"
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
         Width           =   1680
      End
   End
End
Attribute VB_Name = "FRMCANC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim chk As Integer
Dim CAMT As Integer
Dim dat, dat1 As Date

Public Sub CLEAR()
PNRNO.Text = " "
FLIGHT_NO.Text = " "
CLASS.Text = " "
FLIGHT_TYPE.Text = " "
SS_CODE.Text = " "
FLIGHT_DATE = " "
RESERV_DATE.Text = " "
CANCEL_DATE.Text = ""
BRANCH_CODE.Text = " "
TOT_FARE.Text = " "
PASS_NM.Text = " "
PASS_ADD.Text = " "
PASS_NO.Text = " "
PASS_STATUS.Text = " "
End Sub

Public Sub TXT_LOCK(NO As Integer)
If NO = 0 Then
  PNRNO.Locked = True
  FLIGHT_NO.Locked = True
  CLASS.Locked = True
  FLIGHT_TYPE.Locked = True
  SS_CODE.Locked = True
  FLIGHT_DATE.Locked = True
  CANCEL_DATE.Locked = True
  RESERV_DATE.Locked = True
  BRANCH_CODE.Locked = True
  TOT_FARE.Locked = True
  PASS_NM.Locked = True
  PASS_ADD.Locked = True
  PASS_NO.Locked = True
  PASS_STATUS.Locked = True
Else
  PNRNO.Locked = False
  FLIGHT_NO.Locked = False
  CLASS.Locked = False
  FLIGHT_TYPE.Locked = False
  SS_CODE.Locked = False
  FLIGHT_DATE.Locked = False
  CANCEL_DATE.Locked = False
  RESERV_DATE.Locked = False
  BRANCH_CODE.Locked = False
  TOT_FARE.Locked = False
  PASS_NM.Locked = False
  PASS_ADD.Locked = False
  PASS_NO.Locked = False
  PASS_STATUS.Locked = False
End If
End Sub

Private Sub CMDADD_Click()
TXT_LOCK (1)
input_pnr
If chk = 1 Then Exit Sub
dat = Date
CANCEL_DATE.Text = Day(Date) & "/" & Month(Date) & "/" & Year(Date)
CANAMT
'can_amt.Text = Val(TOT_AMT.Text) - Val(CAMT)
PNRNO.SetFocus
CMDCANCEL.Enabled = True
CMDADD.Enabled = False
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
Adodc1.Recordset.CANCEL
Adodc1.RecordSource = "SELECT * FROM CANCELLATION"

Adodc1.Refresh
Adodc1.Recordset.MoveFirst
Text_val
CMDCANCEL.Enabled = False
CMDDELETE.Enabled = True
CMDADD.Enabled = True
CMDSAVE.Enabled = False
CMDEXIT.Enabled = True
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
A = MsgBox("DO YOU WANT TO DELETE", vbYesNo, "MESSAGE")
If A = 6 Then
  With Adodc1.Recordset
      .Delete
      .MoveNext
      If .EOF = True Then .MoveLast
  End With
End If
Text_val
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
n = 0
For I = 1 To Adodc1.Recordset.RecordCount
If Val(A) = Adodc1.Recordset.Fields("pnr") Then
  Text_val
  n = 1
  Exit Sub
  Else
  n = 0
End If
Adodc1.Recordset.MoveNext
Next
If n = 0 Then
  MsgBox ("NO SUCH NUMBER")
End If

End Sub

Private Sub CMDFIRST_Click()
If Adodc1.Recordset.RecordCount = 0 Then
  MsgBox ("THERE IS NO RECORD")
  Exit Sub
End If
Adodc1.Recordset.MoveFirst
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
   MsgBox ("PONTER IS ON THE FIRST RECORD")
End If
Text_val
End Sub

Private Sub CMDPRINT_Click()
CANCEL.Show
End Sub

Private Sub CMDSAVE_Click()
If PNRNO.Text = "" Or FLIGHT_NO.Text = "" Or CANCEL_DATE.Text = "" Or CLASS.Text = "" Or FLIGHT_TYPE.Text = "" Or SS_CODE.Text = "" Or FLIGHT_DATE.Text = "" Or RESERV_DATE.Text = "" Or BRANCH_CODE.Text = "" Or TOT_FARE.Text = "" Or PASS_NM.Text = "" Or PASS_ADD.Text = "" Or PASS_NO.Text = "" Or PASS_STATUS.Text = "" Then
   MsgBox ("INCOMLETE RECORD")
Exit Sub
End If

CMDDELETE.Enabled = True
CMDADD.Enabled = True
CMDSAVE.Enabled = False
CMDCANCEL.Enabled = False
CMDEXIT.Enabled = True
CMDFIRST.Enabled = True
CMDPREVIOUS.Enabled = True
CMDLAST.Enabled = True
CMDFIND.Enabled = True
CMDNEXT.Enabled = True
Adodc1.RecordSource = "SELECT * FROM CANCELLATION"
Adodc1.Refresh
A = MsgBox("DO YOU WANT TO SAVE RECORD", vbYesNo, "MESSAGE")
If A = 6 Then
   FLIGHT_UPDATE
   Adodc1.Recordset.AddNew
   FIELD_VAL
   Adodc1.Recordset.Update
    del_rec_resv
Else
  CLEAR
End If
TXT_LOCK (0)
Form_Load
End Sub
Public Sub FIELD_VAL()
Adodc1.Recordset.Fields("PNR") = PNRNO.Text
Adodc1.Recordset.Fields("FLIGHT_DATE") = FLIGHT_DATE.Text
Adodc1.Recordset.Fields("FLIGHT_NO") = FLIGHT_NO.Text
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
Adodc1.Recordset.Fields("reserv_date") = RESERV_DATE.Text
Adodc1.Recordset.Fields("CANCEL_DATE") = CANCEL_DATE.Text
Adodc1.Recordset.Fields("route_code") = ORIDEST.Text
Adodc1.Recordset.Fields("pass_name") = PASS_NM.Text
Adodc1.Recordset.Fields("pass_add") = PASS_ADD.Text
Adodc1.Recordset.Fields("passport_no") = PASS_NO.Text
Adodc1.Recordset.Fields("ss_code") = SS_CODE.Text
Adodc1.Recordset.Fields("pass_status") = PASS_STATUS.Text
Adodc1.Recordset.Fields("total_fare") = TOT_FARE.Text
Adodc1.Recordset.Fields("branch_code") = Left(BRANCH_CODE.Text, 4)
Adodc1.Recordset.Fields("tot_fare") = TOT_AMT.Text
Adodc1.Recordset.Fields("CANC_AMT") = can_amt.Text
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
dat1 = Adodc1.Recordset.Fields("reserv_date")
RESERV_DATE.Text = Format(Adodc1.Recordset.Fields("reserv_date"), "dd/MM/yyyy")
ORIDEST.Text = Adodc1.Recordset.Fields("route_code")
PASS_NM.Text = Adodc1.Recordset.Fields("pass_name")
PASS_ADD.Text = Adodc1.Recordset.Fields("pass_add")
PASS_NO.Text = Adodc1.Recordset.Fields("passport_no")
SS_CODE.Text = Adodc1.Recordset.Fields("ss_code")
PASS_STATUS.Text = Adodc1.Recordset.Fields("pass_status")
TOT_FARE.Text = Adodc1.Recordset.Fields("total_fare")

For I = 1 To Adodc2.Recordset.RecordCount
     If Adodc2.Recordset.Fields("branch_code") = Adodc1.Recordset.Fields("branch_code") Then
    BRANCH_CODE.Text = Adodc2.Recordset.Fields(0) + "  " + Adodc2.Recordset.Fields(3)
     Exit For
  End If
Adodc2.Recordset.MoveNext
Next
TOT_AMT.Text = Adodc1.Recordset.Fields("tot_fare")
End Sub

Public Sub input_pnr()
chk = 0
Adodc1.RecordSource = "select * from reservation"
Adodc1.Refresh
    If Adodc1.Recordset.RecordCount = 0 Then
    MsgBox ("There is no record in Table")
    Exit Sub
    End If
    
A = InputBox("Enter PNR NUMBER:-")
Adodc1.Refresh
n = 0
For I = 1 To Adodc1.Recordset.RecordCount
If Val(A) = Adodc1.Recordset.Fields("pnr") Then
n = 0
Exit For
Else
n = 1
End If
Adodc1.Recordset.MoveNext
Next
If n = 1 Then
MsgBox ("NO SUCH PNR NUMBER")
Exit Sub
End If
Adodc1.Refresh
For I = 1 To Adodc1.Recordset.RecordCount
    If Val(A) = Adodc1.Recordset.Fields("pnr") Then
        If IsNull(Adodc1.Recordset.Fields("canc_flag")) Or Adodc1.Recordset.Fields("canc_flag") <> "Y" Then
           Text_val
        End If
     Exit For
    End If
    Adodc1.Recordset.MoveNext
Next
If Adodc1.Recordset.BOF = True Then Adodc1.Recordset.MoveLast
If Val(A) = Adodc1.Recordset.Fields("pnr") And Adodc1.Recordset.Fields("canc_flag") = "Y" Then
    MsgBox ("ALREADY CANCELLED")
    chk = 1
   Exit Sub
End If
If Adodc1.Recordset.EOF = True Then
    MsgBox ("THERE IS NO SUCH PNR NUMBER")
    chk = 1
End If
End Sub

Public Sub del_rec_resv()
A = PNRNO.Text
Adodc1.RecordSource = "select * from reservation"
Adodc1.Refresh
For I = 1 To Adodc1.Recordset.RecordCount
If Val(A) = Adodc1.Recordset.Fields("pnr") Then
Adodc1.Recordset.Fields("canc_flag") = "Y"
Adodc1.Recordset.Update
Exit Sub
End If
Adodc1.Recordset.MoveNext
Next
End Sub

Private Sub Form_Load()
Adodc1.RecordSource = "SELECT * FROM CANCELLATION"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
Adodc1.Recordset.MoveFirst
CANCEL_DATE.Text = Adodc1.Recordset.Fields("CANCEL_DATE")
can_amt = Adodc1.Recordset.Fields("canc_amt")
CMDSAVE.Enabled = False
CMDCANCEL.Enabled = False
Text_val
End Sub

Public Sub CANAMT()
I = dat - dat1
Adodc1.RecordSource = "SELECT * FROM CONTROL"
Adodc1.Refresh
If I <= 3 Then
can_amt.Text = Val(TOT_AMT.Text) - Adodc1.Recordset.Fields("CANC_DEDUC_3")
Else
If I > 3 And I <= 6 Then
can_amt.Text = Val(TOT_AMT.Text) - Adodc1.Recordset.Fields("CANC_DEDUC_6")
Else
If I > 6 And I <= 12 Then
can_amt.Text = Val(TOT_AMT.Text) - Adodc1.Recordset.Fields("CANC_DEDUC_12")
Else
If I > 12 Then
can_amt.Text = Val(TOT_AMT.Text) - Val(TOT_AMT.Text) * 0.25
End If
End If
End If
End If

End Sub

Public Sub FLIGHT_UPDATE()
If PASS_STATUS = "WAITING" Then Exit Sub
Adodc4.RecordSource = "SELECT * FROM flight"
Adodc4.Refresh
dt = DateValue(Adodc4.Recordset.Fields("FLIGHT_DATE"))
dt1 = DateValue(FLIGHT_DATE.Text)
For I = 1 To Adodc4.Recordset.RecordCount

b = 0
  If UCase(FLIGHT_NO.Text) = UCase(Adodc4.Recordset.Fields("FLIGHT_NO")) And dt = dt1 Then
     
     b = 1
     Else
     b = 0
  End If
      
    If b = 1 Then
        If UCase(CLASS.Text = "FIRST CLASS") Then Adodc4.Recordset.Fields("FIRST_SEATS_BK") = Adodc4.Recordset.Fields("FIRST_SEATS_BK") - 1
        If UCase(CLASS.Text = "BUSINESS CLASS") Then Adodc4.Recordset.Fields("BUS_SEATS_BK") = Adodc4.Recordset.Fields("BUS_SEATS_BK") - 1
        If UCase(CLASS.Text = "ECONOMIC CLASS") Then Adodc4.Recordset.Fields("ECO_SEATS_BK   ") = Adodc4.Recordset.Fields("ECO_SEATS_BK ") - 1
        Adodc4.Recordset.Update
        Exit Sub
    End If
    Adodc4.Recordset.MoveNext
    
    Next
Y:
    If Z = 1 Then
       Adodc4.Refresh
      'Adodc4.Recordset.AddNew
       Adodc4.Recordset.Fields("FLIGHT_DATE") = FLIGHT_DATE.Text
       Adodc4.Recordset.Fields("FLIGHT_NO") = FLIGHT_NO.Text
       If UCase(CLASS.Text = "FIRST CLASS") Then
           Adodc4.Recordset.Fields("FIRST_SEATS_BK") = 1
           Adodc4.Recordset.Fields("BUS_SEATS_BK") = 0
           Adodc4.Recordset.Fields("ECO_SEATS_BK") = 0
       End If
       If UCase(CLASS.Text = "BUSINESS CLASS") Then
           Adodc4.Recordset.Fields("FIRST_SEATS_BK") = 0
           Adodc4.Recordset.Fields("BUS_SEATS_BK") = 1
           Adodc4.Recordset.Fields("ECO_SEATS_BK") = 0
       End If
       If UCase(CLASS.Text = "ECONOMIC CLASS") Then
           Adodc4.Recordset.Fields("FIRST_SEATS_BK") = 0
           Adodc4.Recordset.Fields("BUS_SEATS_BK") = 0
           Adodc4.Recordset.Fields("ECO_SEATS_BK") = 1
       End If
       Adodc4.Recordset.Update
    End If
End Sub

