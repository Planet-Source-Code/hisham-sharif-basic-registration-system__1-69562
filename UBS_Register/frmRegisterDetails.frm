VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegisterDetails 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  UBS Registration Details"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   15690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Caption         =   "Other Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5280
      TabIndex        =   106
      Top             =   8760
      Width           =   10215
      Begin VB.Label lblTotalModuleP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   112
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Payroll                       :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   33
         Left            =   120
         TabIndex        =   111
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblTotalModuleS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   110
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblTotalModuleA 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   109
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Stock Control             :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   32
         Left            =   120
         TabIndex        =   108
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Accounting Module    :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   31
         Left            =   120
         TabIndex        =   107
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Payment Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5280
      TabIndex        =   95
      Top             =   6480
      Width           =   5055
      Begin VB.TextBox txtSumOfTotalPaid 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   96
         Text            =   "0"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Sum Of Examination Fee    :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   30
         Left            =   120
         TabIndex        =   105
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Sum Of Modules                 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   29
         Left            =   120
         TabIndex        =   104
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   " Sum Of Total to be paid      :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   28
         Left            =   120
         TabIndex        =   103
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Sum Of Total paid               :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   27
         Left            =   120
         TabIndex        =   102
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Sum Of Balance                 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   26
         Left            =   120
         TabIndex        =   101
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblSumOfExamFee 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   100
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblSumOfModuleCharge 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   99
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblSumOfTotalToBePaid 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   98
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblSumOfBalance 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   97
         Top             =   1800
         Width           =   1815
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Class Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   10440
      TabIndex        =   75
      Top             =   6480
      Width           =   5055
      Begin VB.Label Label19 
         Caption         =   "Place available"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   93
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Accounting-Friday         :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   25
         Left            =   120
         TabIndex        =   92
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Accounting-Saturday    :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   24
         Left            =   120
         TabIndex        =   91
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Stock-Saturday             :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   23
         Left            =   120
         TabIndex        =   90
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Stock-Friday                  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   22
         Left            =   120
         TabIndex        =   89
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Payroll-Saturday           :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   21
         Left            =   120
         TabIndex        =   88
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Payroll-Friday                :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   20
         Left            =   120
         TabIndex        =   87
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblAccF2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   86
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblAccS2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   85
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblStkF2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   84
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStkS2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   83
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblPayF2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   82
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblPayS2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   81
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Place available"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   80
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Place available"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   79
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Place available"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   78
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Place available"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   77
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Place available"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   76
         Top             =   1680
         Width           =   1815
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "STEP3. Select Class Date (*Optional)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   10440
      TabIndex        =   66
      Top             =   3480
      Width           =   5055
      Begin VB.ComboBox cmbClass 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmRegisterDetails.frx":0000
         Left            =   120
         List            =   "frmRegisterDetails.frx":0002
         TabIndex        =   72
         Text            =   "Select Class"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.ComboBox cmbClassDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmRegisterDetails.frx":0004
         Left            =   2640
         List            =   "frmRegisterDetails.frx":0006
         TabIndex        =   71
         Text            =   "Select Date"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddClass 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         Picture         =   "frmRegisterDetails.frx":0008
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdDeleteClass 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         Picture         =   "frmRegisterDetails.frx":0592
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         Picture         =   "frmRegisterDetails.frx":0B1C
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         Picture         =   "frmRegisterDetails.frx":10A6
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   720
         Width           =   255
      End
      Begin MSComctlLib.ListView listClass 
         Height          =   1335
         Left            =   120
         TabIndex        =   73
         Top             =   1560
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   2355
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Class"
            Object.Width           =   4128
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   1834
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   $"frmRegisterDetails.frx":1630
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   19
         Left            =   240
         TabIndex        =   74
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Examination Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   47
      Top             =   6480
      Width           =   5055
      Begin VB.Label Label13 
         Caption         =   "Place available"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   65
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Place available"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   64
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Place available"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   63
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Place available"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   62
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Place available"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   61
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblPayS 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   60
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblPayF 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   59
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblStkS 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   58
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblStkF 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   57
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblAccS 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   56
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblAccF 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   55
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Payroll-Friday                :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   18
         Left            =   120
         TabIndex        =   54
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Payroll-Saturday           :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   17
         Left            =   120
         TabIndex        =   53
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Stock-Friday                  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   16
         Left            =   120
         TabIndex        =   52
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Stock-Saturday             :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   15
         Left            =   120
         TabIndex        =   51
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Accounting-Saturday    :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   14
         Left            =   120
         TabIndex        =   50
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Accounting-Friday         :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Place available"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   48
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frameSearch 
      Caption         =   "Search Info-Only For Administrator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5280
      TabIndex        =   29
      Top             =   1200
      Width           =   5055
      Begin VB.OptionButton Option3 
         Caption         =   "By Name"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "By IC"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By Matric"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   840
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txtCari 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   30
         Top             =   360
         Width           =   3375
      End
      Begin MSComctlLib.ListView listSearch 
         Height          =   1335
         Left            =   1440
         TabIndex        =   35
         Top             =   720
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   2355
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Search"
            Object.Width           =   4586
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Search Info:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "STEP 2. Select Modules To Buy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   5280
      TabIndex        =   24
      Top             =   3480
      Width           =   5055
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         Picture         =   "frmRegisterDetails.frx":16D3
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         Picture         =   "frmRegisterDetails.frx":1C5D
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   480
         Width           =   255
      End
      Begin VB.ComboBox cmbModules 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmRegisterDetails.frx":21E7
         Left            =   120
         List            =   "frmRegisterDetails.frx":21F4
         TabIndex        =   27
         Text            =   "Select Modules"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton cmdAddModules 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         Picture         =   "frmRegisterDetails.frx":221C
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdDeleteModules 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Picture         =   "frmRegisterDetails.frx":27A6
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1080
         Width           =   375
      End
      Begin MSComctlLib.ListView listModules 
         Height          =   1335
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   2355
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Modules"
            Object.Width           =   4128
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   $"frmRegisterDetails.frx":2D30
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   6
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "STEP 1. Select Exam To Take"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   5055
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         Picture         =   "frmRegisterDetails.frx":2DC3
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         Picture         =   "frmRegisterDetails.frx":334D
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmdDeleteExam 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         Picture         =   "frmRegisterDetails.frx":38D7
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdAddExam 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         Picture         =   "frmRegisterDetails.frx":3E61
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1080
         Width           =   375
      End
      Begin VB.ComboBox cmbDay 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmRegisterDetails.frx":43EB
         Left            =   2640
         List            =   "frmRegisterDetails.frx":43ED
         TabIndex        =   21
         Text            =   "Select Date"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cmbExam 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmRegisterDetails.frx":43EF
         Left            =   120
         List            =   "frmRegisterDetails.frx":43F1
         TabIndex        =   20
         Text            =   "Select Exam Paper"
         Top             =   1080
         Width           =   2415
      End
      Begin MSComctlLib.ListView listExam 
         Height          =   1335
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   2355
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Modules"
            Object.Width           =   4128
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   1834
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   $"frmRegisterDetails.frx":43F3
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   4
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Payment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   10440
      TabIndex        =   12
      Top             =   1200
      Width           =   5055
      Begin VB.TextBox txtPaid 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   46
         Text            =   "0"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblBalance 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   45
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblToBePaid 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   44
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblModules 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   43
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblExaminationFee 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   42
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Balance                 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Total paid               :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Total to be paid      :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Modules                 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Examination Fee    :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   5055
      Begin VB.TextBox txtNama 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   600
         Width           =   3500
      End
      Begin VB.TextBox txtNoKP 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "######-##-####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   3500
      End
      Begin VB.TextBox txtAlamat1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   1320
         Width           =   3500
      End
      Begin VB.TextBox txtNoHP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   1680
         Width           =   3500
      End
      Begin VB.TextBox txtMatric 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   3500
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "IC. No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "College Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Handphone No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Matric No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   1020
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   15690
      _ExtentX        =   27675
      _ExtentY        =   1799
      ButtonWidth     =   2196
      ButtonHeight    =   1799
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageListCNI"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Update"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print Receipt     "
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            ImageIndex      =   4
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "frmRegisterDetails.frx":4494
   End
   Begin MSComctlLib.ImageList ImageListCNI 
      Left            =   720
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483624
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":45F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":4CE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":555D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":5D9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":65B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":6F07
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":77DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":806F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":8937
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":91AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":9AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":A383
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":AA0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":B2B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":BCEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":C6D3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegisterDetails.frx":CF3A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      Caption         =   "Place available"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   94
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frmRegisterDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim duplicate As Boolean
Dim MModules As Currency
Dim EFee As Currency
Dim TotalPaid As Currency
Dim ToBePaid As Currency
Dim Balance As Currency

Dim totalClassAccF As Integer
Dim totalClassStkF As Integer
Dim totalClassPayF As Integer
Dim totalClassAccS As Integer
Dim totalClassStkS As Integer
Dim totalClassPayS As Integer

Dim currentClassAccF As Integer
Dim currentClassStkF As Integer
Dim currentClassPayF As Integer
Dim currentClassAccS As Integer
Dim currentClassStkS As Integer
Dim currentClassPayS As Integer

Dim totalExamAccF As Integer
Dim totalExamStkF As Integer
Dim totalExamPayF As Integer
Dim totalExamAccS As Integer
Dim totalExamStkS As Integer
Dim totalExamPayS As Integer

Dim currentExamAccF As Integer
Dim currentExamStkF As Integer
Dim currentExamPayF As Integer
Dim currentExamAccS As Integer
Dim currentExamStkS As Integer
Dim currentExamPayS As Integer


Dim datDate As Date
Dim strDay As String

Private Sub cmbClass_Click()
    cmbClassDate.Clear
    cmbClassDate.Text = "Select Date"
    checkConnection
    strSql = "SELECT ClassDate from ddlClass Where UpperCode='" & cmbClass.Text & "'"
    rec1.Open strSql, Con, adOpenStatic
    
    While Not rec1.EOF
        cmbClassDate.AddItem (rec1(0))
        rec1.MoveNext
    Wend
End Sub

Private Sub cmbClassDate_Click()
    checkConnection
    strSql = "SELECT ClassDay from ddlClass Where UpperCode='" & cmbClass.Text & "' AND ClassDate=cdate('" & cmbClassDate.Text & "')"
    rec1.Open strSql, Con, adOpenStatic
    
    strDay = rec1!ClassDay

End Sub

Private Sub cmbDay_Click()
    checkConnection
    strSql = "SELECT ExamDay from ddlClass Where UpperCode='" & cmbExam.Text & "' AND ExamDate=cdate('" & cmbDay.Text & "')"
    rec1.Open strSql, Con, adOpenStatic
    
    strDay = rec1!ExamDay

End Sub

Private Sub cmbExam_Click()
    cmbDay.Clear
    cmbDay.Text = "Select Date"
    checkConnection
    strSql = "SELECT ExamDate from ddlClass Where UpperCode='" & cmbExam.Text & "'"
    rec1.Open strSql, Con, adOpenStatic
    
    While Not rec1.EOF
        cmbDay.AddItem (rec1(0))
        rec1.MoveNext
    Wend

End Sub

Private Sub cmdAddClass_Click()
    If cmbClass.Text = "Select Class" Then
        MsgBox "Please select class!"
        Exit Sub
    End If
    If cmbClassDate.Text = "Select Date" Then
        MsgBox "Please select class date!"
        Exit Sub
    End If
    duplicate = False
    checkClassData
    If duplicate = True Then
        MsgBox "You already register for this class!"
        Exit Sub
    Else
        If cmbClass.Text = "Accounting" And strDay = "Friday" Then
            If lblAccF < 0 Then
                MsgBox "Accounting Friday Class Exceed Limit! Insert cannot be done"
                Exit Sub
            Else
                insertClass
            End If
        ElseIf cmbClass.Text = "Accounting" And strDay = "Saturday" Then
            If lblAccS < 0 Then
                MsgBox "Accounting Saturday Class Exceed Limit! Insert cannot be done"
                Exit Sub
            Else
                insertClass
            End If
        ElseIf cmbClass.Text = "Stock Control" And strDay = "Friday" Then
            If lblStkF < 0 Then
                MsgBox "Stock Control Friday Class Exceed Limit! Insert cannot be done"
                Exit Sub
            Else
                insertClass
            End If
        ElseIf cmbClass.Text = "Stock Control" And strDay = "Saturday" Then
            If lblStkS < 0 Then
                MsgBox "Stock Control Friday Class Exceed Limit! Insert cannot be done"
                Exit Sub
            Else
                insertClass
            End If
        ElseIf cmbClass.Text = "Payroll" And strDay = "Friday" Then
            If lblPayF < 0 Then
                MsgBox "Payroll Friday Class Exceed Limit! Insert cannot be done"
                Exit Sub
            Else
                insertClass
            End If
        ElseIf cmbClass.Text = "Payroll" And strDay = "Saturday" Then
            If lblPayS < 0 Then
                MsgBox "Payroll Saturday Class Exceed Limit! Insert cannot be done"
                Exit Sub
            Else
                insertClass
            End If
        End If
    End If
End Sub
Private Sub insertClass()
        strSql = "INSERT INTO tblClass(MatricNo,Class,CDay,CDate)VALUES('" & txtMatric.Text & "','" & cmbClass.Text & "','" & strDay & "','" & cmbClassDate.Text & "')"
        Con.Execute strSql
        loadClass2
        loadExam
        loadPayment
        ClassInfo
        ExamInfo
End Sub
Private Sub cmdAddExam_Click()
    If cmbExam.Text = "Select Exam Paper" Then
        MsgBox "Please select exam paper!"
        Exit Sub
    End If
    If cmbDay.Text = "Select Date" Then
        MsgBox "Please select exam day!"
        Exit Sub
    End If
    duplicate = False
    checkExamData
    If duplicate = True Then
        MsgBox "You already register for this paper!"
        Exit Sub
    Else
        If cmbExam.Text = "Accounting" And strDay = "Friday" Then
            If lblAccF < 0 Then
                MsgBox "Accounting Friday Exam Exceed Limit! Insert cannot be done"
                Exit Sub
            Else
                InsertExam
            End If
        ElseIf cmbExam.Text = "Accounting" And strDay = "Saturday" Then
            If lblAccS < 0 Then
                MsgBox "Accounting Saturday Exam Exceed Limit! Insert cannot be done"
                Exit Sub
            Else
                InsertExam
            End If
        ElseIf cmbExam.Text = "Stock Control" And strDay = "Friday" Then
            If lblStkF < 0 Then
                MsgBox "Stock Control Friday Exam Exceed Limit! Insert cannot be done"
                Exit Sub
            Else
                InsertExam
            End If
        ElseIf cmbExam.Text = "Stock Control" And strDay = "Saturday" Then
            If lblStkS < 0 Then
                MsgBox "Stock Control Friday Exam Exceed Limit! Insert cannot be done"
                Exit Sub
            Else
                InsertExam
            End If
        ElseIf cmbExam.Text = "Payroll" And strDay = "Friday" Then
            If lblPayF < 0 Then
                MsgBox "Payroll Friday Exam Exceed Limit! Insert cannot be done"
                Exit Sub
            Else
                InsertExam
            End If
        ElseIf cmbExam.Text = "Payroll" And strDay = "Saturday" Then
            If lblPayS < 0 Then
                MsgBox "Payroll Saturday Exam Exceed Limit! Insert cannot be done"
                Exit Sub
            Else
                InsertExam
            End If
        End If
    End If
End Sub
Private Sub InsertExam()
        strSql = "INSERT INTO tblExam(MatricNo,Exam,EDay,EDate,EFee)VALUES('" & txtMatric.Text & "','" & cmbExam.Text & "','" & strDay & "','" & cmbDay.Text & "',50)"
        Con.Execute strSql
        loadClass2
        loadExam
        loadPayment
        ClassInfo
        ExamInfo
End Sub
Private Sub loadExam()
    checkConnection
    strSql = "SELECT Exam,EDate from tblExam Where MatricNo='" & txtMatric.Text & "' order by Exam Asc"
    rec1.Open strSql, Con, adOpenStatic
    
    listExam.ListItems.Clear
    
    While Not rec1.EOF
     Set lst = listExam.ListItems.Add(, , rec1(0), , 1)
       For X = 1 To 1
        lst.SubItems(X) = rec1(X)
       Next X
       rec1.MoveNext
    Wend
End Sub
Private Sub loadClass2()
    checkConnection
    strSql = "SELECT Class,CDate from tblClass Where MatricNo='" & txtMatric.Text & "' order by Class Asc"
    rec1.Open strSql, Con, adOpenStatic
    
    listClass.ListItems.Clear
    
    While Not rec1.EOF
     Set lst = listClass.ListItems.Add(, , rec1(0), , 1)
       For X = 1 To 1
        lst.SubItems(X) = rec1(X)
       Next X
       rec1.MoveNext
    Wend
End Sub

Private Sub loadModules()
    checkConnection
    strSql = "SELECT MModules from tblModules Where MatricNo='" & txtMatric.Text & "' order by MModules Asc"
    rec1.Open strSql, Con, adOpenStatic
    
    Dim X As Integer
    
    listModules.ListItems.Clear
    
    While Not rec1.EOF
     Set lst = listModules.ListItems.Add(, , rec1(0), , 1)
       rec1.MoveNext
    Wend
End Sub
Private Sub loadPayment()
On Error Resume Next
    MModules = 0
    EFee = 0
    Balance = 0
    ToBePaid = 0
    TotalPaid = 0
    checkConnection
    strSql = "SELECT SUM(MCharge) as sumOfMCharge from tblModules Where MatricNo='" & txtMatric.Text & "'"
    rec1.Open strSql, Con, adOpenStatic
    
    lblModules.Caption = FormatCurrency(rec1!sumOfMCharge)
    MModules = FormatCurrency(rec1!sumOfMCharge)
    
    checkConnection
    strSql = "SELECT SUM(EFee) as sumOfEFee from tblExam Where MatricNo='" & txtMatric.Text & "'"
    rec1.Open strSql, Con, adOpenStatic
    
    lblExaminationFee.Caption = FormatCurrency(rec1!sumOfEFee)
    EFee = FormatCurrency(rec1!sumOfEFee)

    lblToBePaid = FormatCurrency(MModules + EFee)
    ToBePaid = lblToBePaid
    
    checkConnection
    strSql = "SELECT TotalPaid from tblPayment Where MatricNo='" & txtMatric.Text & "'"
    rec1.Open strSql, Con, adOpenStatic
    
    txtPaid.Text = FormatCurrency(rec1!TotalPaid)
    TotalPaid = FormatCurrency(rec1!TotalPaid)
    
    Balance = ToBePaid - TotalPaid
    lblBalance = FormatCurrency(Balance)
    
    checkConnection
    strSql = "SELECT SUM(TotalPaid) as sumOfTotalPaid from tblPayment"
    rec1.Open strSql, Con, adOpenStatic
    
    txtSumOfTotalPaid.Text = FormatCurrency(rec1!sumOfTotalPaid)

    checkConnection
    strSql = "SELECT SUM(TotalToPaid) as sumOfTotalToPaid from tblPayment"
    rec1.Open strSql, Con, adOpenStatic
    
    lblSumOfTotalToBePaid.Caption = FormatCurrency(rec1!sumOfTotalToPaid)

    checkConnection
    strSql = "SELECT SUM(Balance) as sumOfBalance from tblPayment"
    rec1.Open strSql, Con, adOpenStatic
    
    lblSumOfBalance.Caption = FormatCurrency(rec1!sumOfBalance)
    
    checkConnection
    strSql = "SELECT SUM(EFee) as sumOfEfee from tblExam"
    rec1.Open strSql, Con, adOpenStatic
    
    lblSumOfExamFee.Caption = FormatCurrency(rec1!sumOfEFee)
    
    checkConnection
    strSql = "SELECT SUM(MCharge) as sumOfMCharge from tblModules"
    rec1.Open strSql, Con, adOpenStatic
    
    lblSumOfModuleCharge.Caption = FormatCurrency(rec1!sumOfMCharge)
    
    checkConnection
    strSql = "SELECT COUNT(MModules) as TotalModules from tblModules Where MModules='Accounting'"
    rec1.Open strSql, Con, adOpenStatic
    
    lblTotalModuleA.Caption = rec1!TotalModules
    
    checkConnection
    strSql = "SELECT COUNT(MModules) as TotalModules from tblModules Where MModules='Stock Control'"
    rec1.Open strSql, Con, adOpenStatic
    
    lblTotalModuleS.Caption = rec1!TotalModules

    checkConnection
    strSql = "SELECT COUNT(MModules) as TotalModules from tblModules Where MModules='Payroll'"
    rec1.Open strSql, Con, adOpenStatic
    
    lblTotalModuleP.Caption = rec1!TotalModules

End Sub
Private Sub cmdAddModules_Click()
    If cmbModules.Text = "Select Modules" Then
        MsgBox "Please select modules!"
        Exit Sub
    End If
    duplicate = False
    checkModulesData
    If duplicate = True Then
        MsgBox "You already decided to buy this module!"
        Exit Sub
    Else
        strSql = "INSERT INTO tblModules(MatricNo,MModules,MCharge)VALUES('" & txtMatric.Text & "','" & cmbModules.Text & "',40)"
        Con.Execute strSql
        loadModules
        loadPayment
    End If
End Sub

Private Sub cmdDeleteClass_Click()
    strSql = "DELETE FROM tblClass WHERE MatricNo='" & txtMatric.Text & "' AND Class='" & listClass.SelectedItem.Text & "'"
    Con.Execute strSql
    loadClass2
    loadPayment
    ClassInfo
    ExamInfo
End Sub

Private Sub cmdDeleteExam_Click()
    strSql = "DELETE FROM tblExam WHERE MatricNo='" & txtMatric.Text & "' AND Exam='" & listExam.SelectedItem.Text & "'"
    Con.Execute strSql
    loadExam
    loadPayment
    ClassInfo
    ExamInfo
End Sub
Private Sub cmdDeleteModules_Click()
    strSql = "DELETE FROM tblModules WHERE MatricNo='" & txtMatric.Text & "' AND MModules='" & listModules.SelectedItem.Text & "'"
    Con.Execute strSql
    
    loadModules
    loadPayment
End Sub
Private Sub Form_Load()
 loadInfo
 loadExam
 loadClass
 loadClass2
 loadModules
 loadPayment
 ClassInfo
 ExamInfo
 determineUser
End Sub
Private Sub determineUser()
 If frmLogin.strPeringkatPenggunaan = 1 Then
    txtPaid.Enabled = True
    frameSearch.Enabled = True
 ElseIf frmLogin.strPeringkatPenggunaan = 2 Then
    txtPaid.Enabled = False
    frameSearch.Enabled = False
    Toolbar.Buttons(2).Enabled = False
 End If
End Sub
Private Sub loadClass()
On Error Resume Next
    cmbExam.Clear
    cmbExam.Text = "Select Exam Paper"
    checkConnection
    strSql = "SELECT UpperCode from ddlClass Group By UpperCode"
    rec1.Open strSql, Con, adOpenStatic
    
    While Not rec1.EOF
        cmbExam.AddItem (rec1(0))
        rec1.MoveNext
    Wend
    
    cmbClass.Clear
    cmbClass.Text = "Select Class"
    checkConnection
    strSql = "SELECT UpperCode from ddlClass Group By UpperCode"
    rec1.Open strSql, Con, adOpenStatic
    
    While Not rec1.EOF
        cmbClass.AddItem (rec1(0))
        rec1.MoveNext
    Wend

End Sub

Public Sub loadInfo()
On Error Resume Next

    checkConnection
    strSql = "SELECT * from tblRegistrationInfo where MatricNo='" & strNamaPengguna & "'"
    rec1.Open strSql, Con, adOpenStatic
    
    txtMatric.Text = rec1!MatricNo
    txtNama.Text = rec1!Name
    txtNoKP.Text = rec1!ICNo
    txtAlamat1.Text = rec1!Address
    txtNoHP.Text = rec1!TelNo
    
End Sub
Private Sub ExamInfo()
On Error Resume Next

    checkConnection
    strSql = "SELECT * from ddlClass where ClassCode='ACC-FRI'"
    rec1.Open strSql, Con, adOpenStatic
    
    totalExamAccF = rec1!Limit
    
    checkConnection
    strSql = "SELECT * from ddlClass where ClassCode='ACC-SAT'"
    rec1.Open strSql, Con, adOpenStatic
    
    totalExamAccS = rec1!Limit

    checkConnection
    strSql = "SELECT * from ddlClass where ClassCode='STK-FRI'"
    rec1.Open strSql, Con, adOpenStatic
    
    totalExamStkF = rec1!Limit
    
    checkConnection
    strSql = "SELECT * from ddlClass where ClassCode='STK-SAT'"
    rec1.Open strSql, Con, adOpenStatic
    
    totalExamStkS = rec1!Limit
    
    checkConnection
    strSql = "SELECT * from ddlClass where ClassCode='PAY-FRI'"
    rec1.Open strSql, Con, adOpenStatic
    
    totalExamPayF = rec1!Limit

    checkConnection
    strSql = "SELECT * from ddlClass where ClassCode='PAY-SAT'"
    rec1.Open strSql, Con, adOpenStatic
    
    totalExamPayS = rec1!Limit

    checkConnection
    strSql = "SELECT COUNT(Exam) as AccExam from tblExam where Exam='Accounting' AND EDay='Friday'"
    rec1.Open strSql, Con, adOpenStatic
    
    currentExamAccF = rec1!AccExam
    
    checkConnection
    strSql = "SELECT COUNT(Exam) as AccExam from tblExam where Exam='Accounting' AND EDay='Saturday'"
    rec1.Open strSql, Con, adOpenStatic
    
    currentExamAccS = rec1!AccExam

    checkConnection
    strSql = "SELECT COUNT(Exam) as AccExam from tblExam where Exam='Stock Control' AND EDay='Friday'"
    rec1.Open strSql, Con, adOpenStatic
    
    currentExamStkF = rec1!AccExam

    checkConnection
    strSql = "SELECT COUNT(Exam) as AccExam from tblExam where Exam='Stock Control' AND EDay='Saturday'"
    rec1.Open strSql, Con, adOpenStatic
    
    currentExamStkS = rec1!AccExam

    checkConnection
    strSql = "SELECT COUNT(Exam) as AccExam from tblExam where Exam='Payroll' AND EDay='Friday'"
    rec1.Open strSql, Con, adOpenStatic
    
    currentExamPayF = rec1!AccExam

    checkConnection
    strSql = "SELECT COUNT(Exam) as AccExam from tblExam where Exam='Payroll' AND EDay='Saturday'"
    rec1.Open strSql, Con, adOpenStatic
    
    currentExamPayS = rec1!AccExam
    
    lblAccF = totalExamAccF - currentExamAccF
    lblAccS = totalExamAccS - currentExamAccS
    lblStkF = totalExamStkF - currentExamStkF
    lblStkS = totalExamStkS - currentExamStkS
    lblPayF = totalExamPayF - currentExamPayF
    lblPayS = totalExamPayS - currentExamPayS
    
    
End Sub
Private Sub ClassInfo()
On Error Resume Next

    checkConnection
    strSql = "SELECT * from ddlClass where ClassCode='ACC-FRI'"
    rec1.Open strSql, Con, adOpenStatic
    
    totalClassAccF = rec1!Limit
    
    checkConnection
    strSql = "SELECT * from ddlClass where ClassCode='ACC-SAT'"
    rec1.Open strSql, Con, adOpenStatic
    
    totalClassAccS = rec1!Limit

    checkConnection
    strSql = "SELECT * from ddlClass where ClassCode='STK-FRI'"
    rec1.Open strSql, Con, adOpenStatic
    
    totalClassStkF = rec1!Limit
    
    checkConnection
    strSql = "SELECT * from ddlClass where ClassCode='STK-SAT'"
    rec1.Open strSql, Con, adOpenStatic
    
    totalClassStkS = rec1!Limit
    
    checkConnection
    strSql = "SELECT * from ddlClass where ClassCode='PAY-FRI'"
    rec1.Open strSql, Con, adOpenStatic
    
    totalClassPayF = rec1!Limit

    checkConnection
    strSql = "SELECT * from ddlClass where ClassCode='PAY-SAT'"
    rec1.Open strSql, Con, adOpenStatic
    
    totalClassPayS = rec1!Limit

    checkConnection
    strSql = "SELECT COUNT(Class) as AccClass from tblClass where Class='Accounting' AND CDay='Friday'"
    rec1.Open strSql, Con, adOpenStatic
    
    currentClassAccF = rec1!AccClass
    
    checkConnection
    strSql = "SELECT COUNT(Class) as AccClass from tblClass where Class='Accounting' AND CDay='Saturday'"
    rec1.Open strSql, Con, adOpenStatic
    
    currentClassAccS = rec1!AccClass

    checkConnection
    strSql = "SELECT COUNT(Class) as AccClass from tblClass where Class='Stock Control' AND CDay='Friday'"
    rec1.Open strSql, Con, adOpenStatic
    
    currentClassStkF = rec1!AccClass

    checkConnection
    strSql = "SELECT COUNT(Class) as AccClass from tblClass where Class='Stock Control' AND CDay='Saturday'"
    rec1.Open strSql, Con, adOpenStatic
    
    currentClassStkS = rec1!AccClass

    checkConnection
    strSql = "SELECT COUNT(Class) as AccClass from tblClass where Class='Payroll' AND CDay='Friday'"
    rec1.Open strSql, Con, adOpenStatic
    
    currentClassPayF = rec1!AccClass

    checkConnection
    strSql = "SELECT COUNT(Class) as AccClass from tblClass where Class='Payroll' AND CDay='Saturday'"
    rec1.Open strSql, Con, adOpenStatic
    
    currentClassPayS = rec1!AccClass
    
    lblAccF2 = totalClassAccF - currentClassAccF
    lblAccS2 = totalClassAccS - currentClassAccS
    lblStkF2 = totalClassStkF - currentClassStkF
    lblStkS2 = totalClassStkS - currentClassStkS
    lblPayF2 = totalClassPayF - currentClassPayF
    lblPayS2 = totalClassPayS - currentClassPayS


End Sub

Private Sub listExam_DblClick()
    checkConnection
    strSql = "SELECT * from tblExam where MatricNo='" & txtMatric.Text & "' AND Exam='" & listExam.SelectedItem.Text & "'"
    rec1.Open strSql, Con, adOpenStatic
    
    cmbExam.Text = rec1!Exam
    cmbDay.Text = rec1!EDay
End Sub

Private Sub listModules_BeforeLabelEdit(Cancel As Integer)
    checkConnection
    strSql = "SELECT * from tblModules where MatricNo='" & txtMatric.Text & "' AND MModules='" & listModules.SelectedItem.Text & "'"
    rec1.Open strSql, Con, adOpenStatic
    
    cmbModules.Text = rec1!MModules

End Sub



Private Sub listSearch_DblClick()
    checkConnection
    strSql = "SELECT * from tblRegistrationInfo where MatricNo='" & listSearch.SelectedItem.Text & "' or ICNo='" & listSearch.SelectedItem.Text & "' or Name='" & listSearch.SelectedItem.Text & "'"
    rec1.Open strSql, Con, adOpenStatic
    
    txtMatric.Text = rec1!MatricNo
    txtNama.Text = rec1!Name
    txtNoKP.Text = rec1!ICNo
    txtAlamat1.Text = rec1!Address
    txtNoHP.Text = rec1!TelNo
    loadClass2
    loadExam
    loadModules
    loadPayment
    ClassInfo
    ExamInfo
    determineUser
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        'Update register info and payment
        strSql = "Update tblPayment set TotalToPaid='" & lblToBePaid & "',TotalPaid='" & txtPaid.Text & "' Where MatricNo='" & txtMatric.Text & "'"
        Con.Execute strSql
        
        strSql = "Update tblRegistrationInfo set Name='" & txtNama.Text & "',ICNo='" & txtNoKP.Text & "',Address='" & txtAlamat1.Text & "',TelNo='" & txtNoHP.Text & "' Where MatricNo='" & txtMatric.Text & "'"
        Con.Execute strSql
        
        MsgBox "Information Updated"
        loadPayment
    ElseIf Button.Index = 2 Then
        strSql = "Update tblPayment set TotalToPaid='" & lblToBePaid & "',TotalPaid='" & txtPaid.Text & "',Balance='" & lblBalance & "' Where MatricNo='" & txtMatric.Text & "'"
        Con.Execute strSql
        
        strSql = "Update tblRegistrationInfo set Name='" & txtNama.Text & "',ICNo='" & txtNoKP.Text & "',Address='" & txtAlamat1.Text & "',TelNo='" & txtNoHP.Text & "' Where MatricNo='" & txtMatric.Text & "'"
        Con.Execute strSql

        Set AppAccess = New Access.Application
        
        AppAccess.Visible = True
        
        With AppAccess
            AppAccess.OpenCurrentDatabase strConnection, False, "brietling1884"
            .RunCommand acCmdAppMaximize
            .DoCmd.OpenReport "rptReceipt", acViewPreview, , "MatricNo = '" & txtMatric.Text & "'"
        End With
        loadPayment
    ElseIf Button.Index = 3 Then

        Unload Me
    End If
End Sub
 Private Sub checkExamData()
    checkConnection
    strSql = "SELECT * from tblExam"
    rec1.Open strSql, Con, adOpenStatic
    
    While Not rec1.EOF
        If rec1!MatricNo = txtMatric.Text And rec1!Exam = cmbExam.Text Then
         duplicate = True
         Exit Sub
        End If
        rec1.MoveNext
    Wend
End Sub
  Private Sub checkClassData()
    checkConnection
    strSql = "SELECT * from tblClass"
    rec1.Open strSql, Con, adOpenStatic
    
    While Not rec1.EOF
        If rec1!MatricNo = txtMatric.Text And rec1!Class = cmbClass.Text Then
         duplicate = True
         Exit Sub
        End If
        rec1.MoveNext
    Wend
End Sub

 Private Sub checkModulesData()
    checkConnection
    strSql = "SELECT * from tblModules"
    rec1.Open strSql, Con, adOpenStatic
    
    While Not rec1.EOF
        If rec1!MatricNo = txtMatric.Text And rec1!MModules = cmbModules.Text Then
         duplicate = True
         Exit Sub
        End If
        rec1.MoveNext
    Wend
End Sub
Private Sub txtCari_Change()
'On Error Resume Next
    Dim X As Integer
    Dim strCariNama As String
    strCariNama = txtCari.Text
    
    If Option1.Value = True Then
        checkConnection
        strSql = "SELECT MatricNo from tblRegistrationInfo Where Left(MatricNo,0) = '" & strCariNama & "' or Left(MatricNo,1) = '" & strCariNama & "' or Left(MatricNo,2) = '" & strCariNama & "' or Left(MatricNo,3) = '" & strCariNama & "' or Left(MatricNo,4) = '" & strCariNama & "' or Left(MatricNo,5) = '" & strCariNama & "' or Left(MatricNo,6) = '" & strCariNama & "' OR Left(MatricNo,7) = '" & strCariNama & "' OR Left(MatricNo,8) = '" & strCariNama & "' OR Left(MatricNo,9) = '" & strCariNama & "' OR Left(MatricNo,10) = '" & strCariNama & "' order by MatricNo asc"
        rec1.Open strSql, Con, adOpenStatic
        
        listSearch.ListItems.Clear
    
        While Not rec1.EOF
         Set lst = listSearch.ListItems.Add(, , rec1(0), , 1)
           rec1.MoveNext
        Wend
    ElseIf Option2.Value = True Then
        checkConnection
        strSql = "SELECT ICNo from tblRegistrationInfo Where Left(ICNo,0) = '" & strCariNama & "' or Left(ICNo,1) = '" & strCariNama & "' or Left(ICNo,2) = '" & strCariNama & "' or Left(ICNo,3) = '" & strCariNama & "' or Left(ICNo,4) = '" & strCariNama & "' or Left(ICNo,5) = '" & strCariNama & "' or Left(ICNo,6) = '" & strCariNama & "' OR Left(ICNo,7) = '" & strCariNama & "' OR Left(ICNo,8) = '" & strCariNama & "' OR Left(ICNo,9) = '" & strCariNama & "' OR Left(ICNo,10) = '" & strCariNama & "' or Left(ICNo,11) = '" & strCariNama & "' or Left(ICNo,12) = '" & strCariNama & "' or Left(ICNo,13) = '" & strCariNama & "' or Left(ICNo,14) = '" & strCariNama & "' order by ICNo asc"
        rec1.Open strSql, Con, adOpenStatic
        listSearch.ListItems.Clear
    
        While Not rec1.EOF
         Set lst = listSearch.ListItems.Add(, , rec1(0), , 1)
           rec1.MoveNext
        Wend

    ElseIf Option3.Value = True Then
        checkConnection
        strSql = "SELECT Name tblRegistrationInfo Where Left(Name,0) = '" & strCariNama & "' or Left(Name,1) = '" & strCariNama & "' or Left(Name,2) = '" & strCariNama & "' or Left(Name,3) = '" & strCariNama & "' or Left(Name,4) = '" & strCariNama & "' or Left(Name,5) = '" & strCariNama & "' or Left(Name,6) = '" & strCariNama & "' OR Left(Name,7) = '" & strCariNama & "' OR Left(Name,8) = '" & strCariNama & "' OR Left(Name,9) = '" & strCariNama & "' OR Left(Name,10) = '" & strCariNama & "' or Left(Name,11) = '" & strCariNama & "' or Left(Name,12) = '" & strCariNama & "' or Left(Name,13) = '" & strCariNama & "' or Left(Name,14) = '" & strCariNama & "' order by Name asc"
        rec1.Open strSql, Con, adOpenStatic
        listSearch.ListItems.Clear
    
        While Not rec1.EOF
         Set lst = listSearch.ListItems.Add(, , rec1(0), , 1)
           rec1.MoveNext
        Wend

    End If
        

End Sub
