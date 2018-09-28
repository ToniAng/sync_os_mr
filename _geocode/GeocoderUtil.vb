Public Class GeocoderUtil
    Private m_Street As String
    Private HNr As Integer

    Private db_con As New cls_db_con

    Public Function Preparestreet(ByVal Str As String, ByVal Gemeindeid As Integer) As String
        Dim tb As DataTable = db_con.GetRecordset("select gemeinden.dbo.preparestreet('" & Str.Replace("'", "´") & "',1," & Gemeindeid & ")")

        If tb.Rows.Count > 0 Then
            Return IIf(IsDBNull(tb.Rows(0)(0)), "", tb.Rows(0)(0))
        Else
            Return Str
        End If
    End Function

    Public Function GetSimIDX(ByVal Str1 As String, ByVal Str2 As String) As Single

        Dim tb As DataTable = db_con.GetRecordset("select gemeinden.dbo.funcsimidx('" & Str1.Trim & "','" & Str2.Trim & "',1,1)")

        If tb.Rows.Count > 0 Then
            Return tb.Rows(0)(0)
        Else
            Return 0.0
        End If
    End Function

    Public Function _GetHNR(ByVal strStreet As String) As Integer
        'Dim sb As New System.Text.StringBuilder
        Dim pos1 As Integer
        Dim pos2 As Integer
        If strStreet.Length = 0 Then Return 0

        For i As Integer = 0 To strStreet.Length - 1
            If IsNumeric(strStreet.Substring(i, 1)) Then
                'sb.Append(strStreet.Substring(i, 1))
                pos1 = i
                Exit For
            End If
        Next

        If pos1 = 0 Then
            If IsNumeric(strStreet) Then
                Return strStreet
            Else
                Return (0)
            End If
        End If


        For i As Integer = pos1 To strStreet.Length - 1
            If Not IsNumeric(strStreet.Substring(i, 1)) Then
                pos2 = i
                Exit For
            End If

        Next

        If pos1 >= pos2 Then
            Return strStreet.Substring(pos1)
        Else
            Return strStreet.Substring(pos1, pos2 - pos1)

        End If


    End Function

    Public Function GetStreet(ByVal Street As String, ByVal Ort As String, ByVal GemeindeId As Integer, ByVal fromOrig As Boolean) As String
        If Not fromOrig And GemeindeId > 0 Then Street = Me.Preparestreet(Street, GemeindeId)

        If Street.Length = 0 Then
            Return ""
        End If
        If Street.Length >= 4 Then
            If Street.Substring(0, 4) = "Nr. " Or Street.Substring(0, 3) = "Nr " Then
                Street = Ort & " " & Street.Replace("Nr. ", "").Replace("Nr ", "")
                'Street = Street.Replace("Nr ", "")
                'Street = Ort & " " & Street
            End If
        End If
        If IsNumeric(Street) Then
            HNr = Street
            Return Ort
        End If

        'Es gibt Strassennamen, die mit Ziffern beginnen 


        HNr = Me._GetHNR(Street)

        If HNr > 0 Then
            'Wenn Hausnummer dann Ort=Strasse
            If HNr.ToString.Length + 5 > Street.Length Then
                Return Ort
            End If
        End If

        For i As Integer = HNr.ToString.Length + 1 To Street.Length - 1
            If IsNumeric(Street.Substring(i, 1)) Then
                HNr = Me._GetHNR(Street.Substring(i))
                Return Street.Substring(0, i - 1)


            End If
        Next
        Return Street
    End Function

    Public Function GetHNR() As Integer
        Return hnr
    End Function

    'Public Function LoadKG() As DataTable
    '    Dim tb As DataTable = db_con.GetRecordset("select distinct left(gemeindeid,5) from ghdaten_gis..at_geometrie where aufloesung=" &
    '        geocoder.AreaLevels.GEMEINDE & " and gemeindeid>99999")
    '    For Each r As DataRow In tb.Rows
    '        geocoder.KG.Add(r(0))
    '    Next
    '    geocoder.KGList = db_con.GetRecordset("select gmf,  left(gemeindeid,5) gemeindeid, gemeindeid kg_id  from ghdaten_gis..at_geometrie where aufloesung=" &
    '        geocoder.AreaLevels.GEMEINDE & " and gemeindeid>99999")

    'End Function


End Class