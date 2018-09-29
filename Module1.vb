Module Module1
    Public Setting As New Einstellungen
    Public c As New SqlClient.SqlConnection
    Sub Main()
        Dim osync As New sync

        Try
            osync.SyncVacc()

        Catch ex As Exception
            osync.AktionsLog("Fehler: " & ex.Message & vbNewLine & ex.StackTrace, sync.AktionslogKat.Integration_BH_Impfungen)
        End Try


    End Sub

End Module
