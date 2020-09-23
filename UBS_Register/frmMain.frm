VERSION 5.00
Object = "Word.Document.8"; "WINWORD.EXE"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9315
   ClientLeft      =   390
   ClientTop       =   1365
   ClientWidth     =   19005
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   19005
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Instructions && Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   18735
      Begin WordCtl.Document Document2 
         Height          =   4500
         Left            =   14040
         OleObjectBlob   =   "frmMain.frx":0000
         TabIndex        =   2
         Top             =   240
         Width           =   4545
      End
      Begin WordCtl.Document Document1 
         Height          =   8535
         Left            =   240
         OleObjectBlob   =   "frmMain.frx":1E818
         TabIndex        =   1
         Top             =   360
         Width           =   3795
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
