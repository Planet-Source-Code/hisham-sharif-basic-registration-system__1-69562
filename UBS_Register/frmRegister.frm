VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegister 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Please Fill In Required Details"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbSemester 
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
      ItemData        =   "frmRegister.frx":0000
      Left            =   1440
      List            =   "frmRegister.frx":001F
      TabIndex        =   14
      Text            =   "Select Course"
      Top             =   3360
      Width           =   7695
   End
   Begin VB.ComboBox cmbCourse 
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
      Left            =   1440
      TabIndex        =   13
      Text            =   "Select Course"
      Top             =   3000
      Width           =   7695
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
      TabIndex        =   0
      Top             =   1200
      Width           =   7695
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
      TabIndex        =   5
      Top             =   2640
      Width           =   7695
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
      Top             =   2280
      Width           =   7695
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
      TabIndex        =   2
      Top             =   1920
      Width           =   7695
   End
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
      TabIndex        =   1
      Top             =   1560
      Width           =   7695
   End
   Begin MSComctlLib.ImageList ImageListCNI 
      Left            =   8640
      Top             =   960
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
            Picture         =   "frmRegister.frx":0042
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegister.frx":0734
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegister.frx":0FA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegister.frx":17E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegister.frx":1FFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegister.frx":2953
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegister.frx":322B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegister.frx":3ABB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegister.frx":4383
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegister.frx":4BF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegister.frx":54FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegister.frx":5DCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegister.frx":6459
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegister.frx":6CFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegister.frx":773A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegister.frx":811F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8040
      Top             =   960
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
            Picture         =   "frmRegister.frx":8986
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   1020
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   1799
      ButtonWidth     =   1535
      ButtonHeight    =   1799
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageListCNI"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "      Exit     "
            ImageIndex      =   4
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "frmRegister.frx":8E10
   End
   Begin VB.Label Label1 
      Caption         =   "Semester :"
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
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Course :"
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
      TabIndex        =   11
      Top             =   3000
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
      TabIndex        =   10
      Top             =   1200
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
      TabIndex        =   8
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Home Address:"
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
      TabIndex        =   7
      Top             =   2280
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
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
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
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim duplicate As Boolean

Private Sub Form_Load()
    loadCourse
End Sub
Private Sub loadCourse()
    cmbCourse.Clear
    cmbCourse.Text = "Select Course"
    checkConnection
    strSql = "SELECT CourseCode from ddlCourse"
    rec1.Open strSql, Con, adOpenStatic
    
    While Not rec1.EOF
        cmbCourse.AddItem (rec1(0))
        rec1.MoveNext
    Wend

End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        If txtMatric.Text = "" Or txtNoKP.Text = "" Or txtNama.Text = "" Or txtAlamat1.Text = "" Or txtNoHP.Text = "" Then
            MsgBox "Incomplete Information!,Please Fill In Required Details"
            Exit Sub
        End If
        duplicate = False
        checkData
        If duplicate = True Then
            MsgBox "Matric No. Already Exist!"
            txtMatric.Text = ""
            Exit Sub
        Else
            strSql = "INSERT INTO tblRegistrationInfo(MatricNo,Name,ICNo,Address,TelNo,Course,Semester,LevelOfUser)VALUES('" & txtMatric.Text & "','" & txtNama.Text & "','" & txtNoKP.Text & "','" & txtAlamat1.Text & "','" & txtNoHP.Text & "','" & cmbCourse.Text & "','" & cmbSemester.Text & "','2')"
            Con.Execute strSql
            
            strSql = "INSERT INTO tblPayment(MatricNo)VALUES('" & txtMatric.Text & "')"
            Con.Execute strSql

            MsgBox "Information Saved"
        End If
    ElseIf Button.Index = 2 Then
        'Kemaskini
    ElseIf Button.Index = 3 Then
        Unload Me
    End If
End Sub
 Private Sub checkData()
    checkConnection
    strSql = "SELECT MatricNo from tblRegistrationInfo"
    rec1.Open strSql, Con, adOpenStatic
    
    While Not rec1.EOF
        If rec1!MatricNo = txtMatric.Text Then
         duplicate = True
         Exit Sub
        End If
        rec1.MoveNext
    Wend
End Sub

