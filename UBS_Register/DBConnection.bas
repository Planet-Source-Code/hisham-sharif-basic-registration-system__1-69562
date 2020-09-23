Attribute VB_Name = "DBConnection"
Public Con As ADODB.Connection
Public rec1 As ADODB.Recordset
Public rec2 As ADODB.Recordset
Public rec3 As ADODB.Recordset
Public strConnection As String
Public strNamaPengguna As String
Public strKataLaluan As String
Public strSql As String
Public strSql2 As String
Public strSql3 As String
Dim AppAccess As Access.Application
Public Sub conn()
    Set Con = New ADODB.Connection
    Con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\db\dbUBS_Register.mdb;Jet OLEDB:Database Password=brietling1884;"
    Set rec1 = New ADODB.Recordset
    Set rec2 = New ADODB.Recordset
    Set rec3 = New ADODB.Recordset
    strConnection = "" & App.Path & "\db\dbUBS_Register.mdb"
End Sub
Sub Main()
 conn
 mdiMain.Show
 frmMain.Show
 frmLogin.Show vbModal
End Sub
Public Sub checkConnection()
 If rec1.State = 1 Then rec1.Close
End Sub
Public Sub checkConnection2()
 If rec2.State = 1 Then rec2.Close
End Sub
Public Sub checkConnection3()
 If rec3.State = 1 Then rec3.Close
End Sub



