' / --------------------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gsoft.com
' /
' / Microsoft Visual Basic .NET (2010)
' /
' / This is open source code under @Copyleft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------------------
Imports System.Data.OleDb
Imports Microsoft.VisualBasic

Module modDataBase
    '// Declare variable one time but use many times.
    Public Conn As OleDbConnection
    Public Cmd As OleDbCommand
    Public DS As DataSet
    Public DR As OleDbDataReader
    Public DA As OleDbDataAdapter
    Public strSQL As String '// Major SQL
    'Public strStmt As String    '// Minor SQL

    '// Data Path 
    Public strPathData As String = MyPath(Application.StartupPath) & "Data\"
    '// Images Path
    Public strPathImages As String = MyPath(Application.StartupPath) & "Images\"

    Public Function ConnectDataBase() As Boolean
        Conn = New OleDbConnection(
            "Provider = Microsoft.ACE.OLEDB.12.0;" &
            "Data Source = " & strPathData & "QRCode.accdb"
            )
        Try
            Conn.Open()
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Report Status", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End
        End Try
    End Function

    ' / --------------------------------------------------------------------------------
    ' / Function to find and create the new Primary Key not to duplicate.
    Public Function SetupNewPK(ByVal Sql As String) As Long
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        Cmd = New OleDbCommand(Sql, Conn)
        '/ Check if the information is available. And return it back
        If IsDBNull(Cmd.ExecuteScalar) Then
            '// Start at 1
            SetupNewPK = 1
        Else
            SetupNewPK = Cmd.ExecuteScalar + 1
        End If
    End Function

    ' / --------------------------------------------------------------------------------
    ' / Get my project path
    ' / AppPath = C:\My Project\bin\debug
    ' / Replace "\bin\debug" with "\"
    ' / Return : C:\My Project\
    Function MyPath(AppPath As String) As String
        '/ MessageBox.Show(AppPath);
        AppPath = AppPath.ToLower()
        '/ Return Value
        MyPath = AppPath.Replace("\bin\debug", "\").Replace("\bin\release", "\").Replace("\bin\x86\debug", "\")
        '// If not found folder then put the \ (BackSlash) at the end.
        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    End Function

End Module
