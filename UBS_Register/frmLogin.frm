VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGoRegister 
      Caption         =   "Click To Register"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   2505
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5805
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   2535
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   5550
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1935
            Left            =   120
            Picture         =   "frmLogin.frx":0000
            ScaleHeight     =   1935
            ScaleWidth      =   1905
            TabIndex        =   11
            Top             =   240
            Width           =   1905
         End
         Begin VB.TextBox txtNamaPengguna 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3375
            TabIndex        =   1
            Top             =   480
            Width           =   2040
         End
         Begin VB.TextBox txtKataLaluan 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   3375
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   960
            Width           =   2040
         End
         Begin VB.CommandButton cmdLogin 
            Caption         =   "Log-In"
            BeginProperty Font 
               Name            =   "Lucida Sans"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   4050
            Picture         =   "frmLogin.frx":16F3
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1440
            Width           =   1305
         End
         Begin VB.CommandButton cmdOut 
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Lucida Sans"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   2160
            Picture         =   "frmLogin.frx":1CAF
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "IC Number"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   13
            Top             =   960
            Width           =   1290
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Matrix Number"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   12
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   5550
         Begin VB.Label lblCompanyProduct 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "UBS Online Registration System"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   5280
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Faculty Of Accountancy Universiti Utara Malaysia"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   5220
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   0
         TabIndex        =   6
         Top             =   4560
         Width           =   255
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "First time user,Please Register"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   3180
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strPeringkatPenggunaan As String
Public strDomain As String

Private Sub cmdGoRegister_Click()
    frmRegister.Show 1
End Sub

Private Sub cmdLogin_Click()
    checkConnection
    strNamaPengguna = txtNamaPengguna.Text
    strKataLaluan = txtKataLaluan.Text
    
    strSql = "SELECT * from tblRegistrationInfo where MatricNo='" & strNamaPengguna & "'"
    rec1.Open strSql, Con, adOpenStatic
    
        'Adakah user memasukkan maklumat?
        If txtKataLaluan.Text = "" Or txtNamaPengguna.Text = "" Then
           MsgBox "Sila penuhkan semua maklumat terlebih dahulu!!", 0 + 48, "Mesej"
           txtKataLaluan.Text = ""
           txtNamaPengguna.Text = ""
           txtNamaPengguna.SetFocus
           Exit Sub
          
        ElseIf strKataLaluan = rec1!ICNo Then
          DBConnection.strNamaPengguna = strNamaPengguna
          strNamaKakitangan = rec1!Name
          strPeringkatPenggunaan = rec1!LevelOfUser
            If strPeringkatPenggunaan = 1 Then
                Unload Me
           ElseIf strPeringkatPenggunaan = 2 Then
                Unload Me
           End If

        Else
            MsgBox "Harap Maaf, maklumat Log-In Salah, Sila Cuba Semula", 0 + 48, "Mesej"
            txtKataLaluan.Text = ""
            txtNamaPengguna.Text = ""
            txtNamaPengguna.SetFocus
        End If
        'mdiMain.determineUser
End Sub

Private Sub cmdOut_Click()
 End
End Sub
Private Sub Form_Load()
            txtKataLaluan.Text = ""
            txtNamaPengguna.Text = ""
End Sub
Private Sub txtKataLaluan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    checkConnection
    strNamaPengguna = txtNamaPengguna.Text
    strKataLaluan = txtKataLaluan.Text
    
    strSql = "SELECT * from tblRegistrationInfo where MatricNo='" & strNamaPengguna & "'"
    rec1.Open strSql, Con, adOpenStatic
    
        'Adakah user memasukkan maklumat?
        If txtKataLaluan.Text = "" Or txtNamaPengguna.Text = "" Then
           MsgBox "Sila penuhkan semua maklumat terlebih dahulu!!", 0 + 48, "Mesej"
           txtKataLaluan.Text = ""
           txtNamaPengguna.Text = ""
           txtNamaPengguna.SetFocus
           Exit Sub
          
        ElseIf strKataLaluan = rec1!ICNo Then
          DBConnection.strNamaPengguna = strNamaPengguna
          strNamaKakitangan = rec1!Name
          strPeringkatPenggunaan = rec1!LevelOfUser
            If strPeringkatPenggunaan = 1 Then
                Unload Me
           ElseIf strPeringkatPenggunaan = 2 Then
                Unload Me
           End If

        Else
            MsgBox "Harap Maaf, maklumat Log-In Salah, Sila Cuba Semula", 0 + 48, "Mesej"
            txtKataLaluan.Text = ""
            txtNamaPengguna.Text = ""
            txtNamaPengguna.SetFocus
        End If
        'mdiMain.determineUser
        ElseIf KeyAscii = 27 Then
        End
    End If
    'mdiMain.determineUser
        
End Sub

