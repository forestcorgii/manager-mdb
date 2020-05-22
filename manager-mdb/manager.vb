Public Class manager
    '    Public Const MDBConfigFileExtension = ".mdb.config.xml"

    Public Connection As OleDb.OleDbConnection
    Public DBPath As String
    Public UserID As String
    Public Password As String

    Sub New(_dbpath As String, Optional _userid As String = "", Optional _password As String = "", Optional openNow As Boolean = False)
        DBPath = _dbpath
        UserID = _userid
        Password = _password
        If openNow Then
            Open()
        End If
    End Sub


    Public Sub Open()
        Open(DBPath, Connection, UserID, Password)
    End Sub
    Public Sub CreateTable(ByVal tbl As String, ByVal flds As String())
        CreateTable(tbl, flds, Connection)
    End Sub
    Public Sub ExecuteQuery(ByVal Qry As String)
        ExecuteQuery(Qry, Connection)
    End Sub
    Public Sub Insert(ByVal tbl As String, ByVal fld As String(), ByVal val As Object())
        Insert(tbl, fld, val, Connection)
    End Sub
    Public Sub Update(ByVal tbl As String, ByVal fld As String(), ByVal val As Object(), ByVal condition As Object())
        Update(tbl, fld, val, condition, Connection)
    End Sub
    Public Function ToDT(ByVal qry As String) As DataTable
        Return ToDT(qry, Connection)
    End Function
    Public Function CheckTable(tbl As String) As Boolean
        Return CheckTable(tbl, Connection)
    End Function

    Public Sub Close()
        Connection.Close()
        Connection.Dispose()
    End Sub

    Public Shared Sub Open(ByVal _dbpath As String, ByRef _con As OleDb.OleDbConnection, Optional _userid As String = "", Optional _password As String = "")
        Try
            _con = New System.Data.OleDb.OleDbConnection(String.Format("Provider=Microsoft.JET.OLEDB.4.0;Data Source={0};User Id={1};Password={2};", _dbpath, _userid, _password))
            _con.Open()
        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Public Sub SaveToDBF(dbfPath As String, flds As String(),values As List(Of clspay)
    '    Dim ExportedDbf = New SocialExplorer.IO.FastDBF.DbfFile(Text.Encoding.GetEncoding(1252))
    '    ExportedDbf.Open(dbfPath, FileMode.Create)


    '    For Each col In flds
    '        ExportedDbf.Header.AddColumn(New SocialExplorer.IO.FastDBF.DbfColumn(col.ToString, SocialExplorer.IO.FastDBF.DbfColumn.DbfColumnType.Character, 50, 0))
    '    Next

    '    Dim Counter As Integer = 0

    '    For Each row In dt.Rows

    '        Dim ColumnCounter As Integer = 0
    '        Dim NewRec = New SocialExplorer.IO.FastDBF.DbfRecord(ExportedDbf.Header)

    '        For Each col In flds

    '            NewRec(ColumnCounter) = dt.Rows(Counter)(col).ToString
    '            ColumnCounter = ColumnCounter + 1
    '        Next

    '        ExportedDbf.Write(NewRec, True)
    '        Counter = Counter + 1

    '    Next

    '    ExportedDbf.Close()
    'End Sub

    Public Shared Sub Create(ByVal _dbpath As String, Optional _userid As String = "", Optional _password As String = "")
        Try
            Dim cat As New ADOX.Catalog
            cat.Create(String.Format("Provider=Microsoft.JET.OLEDB.4.0;Data Source={0};User Id={1};Password={2};", _dbpath, _userid, _password))
        Catch Ex As System.Exception
        End Try
    End Sub

    Public Shared Sub CreateTable(ByVal tbl As String, ByVal flds As String(), ByVal con As OleDb.OleDbConnection)
        Dim qry As String = String.Format("CREATE TABLE {0}(", tbl)
        For i As Integer = 0 To flds.Length - 1
            qry &= IIf(i = 0, flds(i), "," & flds(i))
        Next
        qry &= ")"

        ExecuteQuery(qry, con)
    End Sub
    Public Shared Function ToDT(ByVal qry As String, ByVal con As OleDb.OleDbConnection) As DataTable
        Try
            Dim dt As New DataTable
            Dim da As New OleDb.OleDbDataAdapter(qry, con)
            da.Fill(dt)
            Return dt
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Shared Sub ExecuteQuery(ByVal Qry As String, ByVal con As OleDb.OleDbConnection)
        Try
            Dim com As New OleDb.OleDbCommand(Qry, con)
            com.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Shared Sub Insert(ByVal tbl As String, ByVal fld As String(), ByVal val As Object(), con As OleDb.OleDbConnection)
        Dim qry As String = String.Format("INSERT INTO {0} (", tbl)
        Dim valtype As String = ""

        For i As Integer = 0 To fld.Length - 1
            Dim f As String = fld(i)
            If f = fld(0) Then
                qry &= String.Format("[{0}]", f)
            Else
                qry &= String.Format(",[{0}]", f)
            End If
        Next

        qry &= ") VALUES("

        For i As Integer = 0 To val.Length - 1
            Dim v = val(i)
            valtype = TypeName(v)
            If i = 0 Then
                If valtype = "String" Then
                    qry &= String.Format("'{0}'", v)
                Else
                    qry &= String.Format("{0}", v)
                End If
            Else
                If valtype = "String" Then
                    qry &= String.Format(",'{0}'", v)
                Else
                    qry &= String.Format(",{0}", v)
                End If
            End If
        Next
        qry &= ")"

        ExecuteQuery(qry, con)
    End Sub

    Public Shared Sub Update(ByVal tbl As String, ByVal fld As String(), ByVal val As Object(), ByVal condition As Object(), ByVal con As OleDb.OleDbConnection)
        Dim qry As String = String.Format("UPDATE {0} SET ", tbl)
        Dim valtype As String = ""

        If fld.Length = val.Length Then
            For f As Integer = 0 To fld.GetUpperBound(0)
                valtype = TypeName(val(f))
                If f = 0 Then
                    If valtype = "String" Then
                        qry &= String.Format("[{0}]='{1}'", fld(f), val(f))
                    Else
                        qry &= String.Format("[{0}]={1}", fld(f), val(f))
                    End If
                Else
                    If valtype = "String" Then
                        qry &= String.Format(",[{0}]='{1}'", fld(f), val(f))
                    Else
                        qry &= String.Format(",[{0}]={1}", fld(f), val(f))
                    End If
                End If
            Next
        End If

        If Not condition Is Nothing Then
            If TypeName(condition(1)) = "String" Then
                qry &= String.Format(" WHERE {0} = '{1}'", condition(0), condition(1))
            Else
                qry &= String.Format(" WHERE {0} = {1}", condition(0), condition(1))
            End If
        End If

        ExecuteQuery(qry, con)
    End Sub

    Public Shared Function CheckTable(tbl As String, con As OleDb.OleDbConnection) As Boolean
        Return getTables(con).Contains(tbl)
    End Function

    Public Shared Function getTables(ByVal con As OleDb.OleDbConnection) As List(Of String)
        getTables = New List(Of String)
        Dim restrictions() As String = New String(3) {}
        restrictions(3) = "Table"
        Dim dt As DataTable = con.GetSchema("Tables", restrictions)
        For i As Integer = 0 To dt.Rows.Count - 1
            getTables.Add(dt.Rows(i)(2).ToString)
        Next
        Return getTables
    End Function

End Class
