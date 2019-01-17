Imports System.Configuration
Public Class cls_db_con
    Private Const STANDBSTRING As String = "User Id=GHDaten_LesenSchreiben; password=nww3vz..ke8h!qvdwknyqeh:hfkj??ew-abfc;Connect Timeout=120;Initial Catalog=ghdaten;Data Source=."
    Private con As SqlClient.SqlConnection
    Public LastErrSQL As String = ""

    Public Sub New()
        con = New SqlClient.SqlConnection
        con.ConnectionString = "User Id=GHDaten_LesenSchreiben; password=nww3vz..ke8h!qvdwknyqeh:hfkj??ew-abfc;Connect Timeout=120;Initial Catalog=ghdaten;Data Source=."
    End Sub


    Public Function GetCon(Optional ByVal Verbose As Boolean = True) As SqlClient.SqlConnection
        Try

            If con.State = ConnectionState.Open Or con.State = ConnectionState.Connecting Then Return con



            Try
                If con.ConnectionString = "" Then

                    con.ConnectionString = "User Id=GHDaten_LesenSchreiben; password=nww3vz..ke8h!qvdwknyqeh:hfkj??ew-abfc;Connect Timeout=120;Initial Catalog=ghdaten;Data Source=."

                End If
                'MsgBox(con.ConnectionString)
                If Not (con.State = ConnectionState.Connecting Or con.State = ConnectionState.Open) Then

                    con.Open()

                End If

                Return con

            Catch ex As Exception

            End Try

            If con.State = ConnectionState.Closed Then
                Throw New System.Exception("Verbindung zur Datenbank konnte nicht hergestellt werden.")
                Return con
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally


        End Try


        Return Nothing

    End Function



    Public Function FireSQL(ByVal SQL As String,
                                Optional ByVal trans As SqlClient.SqlTransaction = Nothing,
                                Optional ByVal c As SqlClient.SqlConnection = Nothing) As Integer
        Try


            Dim m_con As SqlClient.SqlConnection
            Dim ret As Integer

            Debug.Write(Date.Now & " " & SQL & Chr(13))

            If trans Is Nothing Then
                If Not (c Is Nothing) Then
                    m_con = c
                Else
                    If con.State = ConnectionState.Closed Then con.Open()
                    m_con = con
                End If
                'Log("Ohne Transaktion: " & SQL)
            Else
                If Not trans.Connection Is Nothing Then
                    m_con = trans.Connection
                Else
                    If con.State = ConnectionState.Closed Then Me.GetCon()
                    m_con = con

                End If

                'Log("Transaktion: " & SQL)


            End If
            'WaitForConnection(m_con)

            Dim SQLCOM As New SqlClient.SqlCommand(SQL, m_con)
            If Not trans Is Nothing Then SQLCOM.Transaction = trans
            SQLCOM.CommandTimeout = con.ConnectionTimeout


            'If oLogger Is Nothing Then oLogger = New LogSQL(STORE)
            'oLogger.LogSQL(SQLCOM)



            ret = SQLCOM.ExecuteNonQuery()

            Return ret

        Catch ex As Exception
            LastErrSQL = SQL
            'Log("SQL-Fehler: " & SQL)
            'Log(ex.Message & Chr(13) & ex.StackTrace)
            Throw New Exception(ex.Message & vbNewLine & "SQL: " & SQL)
        Finally


        End Try
    End Function

    Public Function GetRecordSet(ByVal SQL As String,
                           Optional ByVal trans As SqlClient.SqlTransaction = Nothing,
                           Optional ByVal c As SqlClient.SqlConnection = Nothing,
                           Optional ByVal startRecord As Integer = 0,
                           Optional ByVal maxRecords As Integer = 0) As DataTable
        Try


            Dim tb As New DataTable
            Dim m_con As SqlClient.SqlConnection

            Debug.Write(Date.Now & " " & SQL & Chr(13))

            If trans Is Nothing Then
                If Not (c Is Nothing) Then
                    m_con = c
                Else
                    If Not con.State = ConnectionState.Open Then con.Open()

                    m_con = con
                End If
                'Log("Ohne Transaktion: " & SQL)
            Else
                If Not trans.Connection Is Nothing Then
                    m_con = trans.Connection
                Else
                    If con.State = ConnectionState.Closed Then Me.GetCon()
                    m_con = con

                End If
                'Log("Transaktion: " & SQL)

            End If
            'WaitForConnection(m_con)
            Dim da As New SqlClient.SqlDataAdapter(SQL, m_con)



            Try
                If Not trans Is Nothing Then da.SelectCommand.Transaction = trans
                If startRecord = 0 And maxRecords = 0 Then

                    da.Fill(tb)

                Else
                    Dim ds As New DataSet
                    da.Fill(ds, startRecord, maxRecords, "tblDummy")
                    tb = ds.Tables(0)
                End If
            Catch ex As Exception
                'Log(ex.Message & " SQL: " & SQL)

                'Log(ex.Message & Chr(13) & ex.StackTrace)
                LastErrSQL = SQL
                Throw New Exception(ex.Message)

            End Try


            'MsgBox("Aus:" & Chr(13) & SQL)



            Return tb

        Catch ex As Exception
            LastErrSQL = SQL

            'Log(ex.Message & " SQL: " & SQL)
            'Log(ex.Message & Chr(13) & ex.StackTrace)
            Throw New Exception(ex.Message)
        Finally

        End Try

    End Function
End Class
