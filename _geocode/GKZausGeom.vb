Imports System.Drawing

Public Class GKZausGeom

    Private Const GKZ_GRAZ As Integer = 60101
    Private GKZ_GRAZER_BEZIRKE_UNBEK As Integer = 6010199

    Private Geom_Stmk As New List(Of Geom)
    Private Geom_Graz As New List(Of Geom)
    Private m_grazerbezirke As Boolean


    Public Sub New(GrazerBezirke As Boolean)
        LoadGeom(GrazerBezirke)
    End Sub


    Private Sub LoadGeom(GrazerBezirke As Boolean)
        Dim tb As DataTable
        m_grazerbezirke = GrazerBezirke


        Dim db_con As New cls_db_con
        tb = db_con.GetRecordSet("select * from ghdaten..geojson where geo_id=101 ")
        _loadgeom(tb, Geom_Stmk)
        tb = db_con.GetRecordSet("select * from ghdaten..geojson where geo_id=100 ")
        _loadgeom(tb, Geom_Graz)




    End Sub


    Private Sub _loadgeom(tb As DataTable, geoms As List(Of Geom))

        Dim j As Integer = 0

        Dim zeilen As String() = tb.Rows(j)("geo_data").Split(vbLf)

        For i As Integer = 1 To zeilen.Length - 2
            zeilen(i) = zeilen(i).Replace(" ", "")
            Dim pos As Integer = zeilen(i).IndexOf("[[[")
            Dim pos2 As Integer = zeilen(i).IndexOf("]]]")
            If pos >= 0 And pos2 >= 0 Then
                Dim arrstr As String() = zeilen(i).Substring(pos + 3, pos2 - pos - 3).Split("],[")

                Dim poly(arrstr.Length) As System.Drawing.PointF
                For k As Integer = 0 To arrstr.Length - 1
                    Dim strpoint As String() = arrstr(k).Replace(",[", "").Replace("[", "").Replace("]", "").Split(",")
                    strpoint(0) = strpoint(0).Replace(".", ",")
                    strpoint(1) = strpoint(1).Replace(".", ",")

                    Dim p As PointF
                    p.X = strpoint(0)
                    p.Y = strpoint(1)
                    poly(k) = p

                Next


                Dim g As New Geom
                g.Polygon = poly
                g.GKZ = GetProperty(tb.Rows(j)("geo_gkz"), zeilen(i))






                geoms.Add(g)
                'End If
            End If

        Next
    End Sub

    Private Function GetProperty(prop As String, json As String) As String


        Dim pos As Integer = json.IndexOf(prop)
        Dim pos2 As Integer = json.IndexOf(",", pos)
        Dim p As String = json.Substring(pos, pos2 - pos)

        Dim sb As New System.Text.StringBuilder
        Dim DPPassed As Boolean = False

        For i As Integer = 0 To p.Length - 1
            If p.Substring(i, 1) = ":" Then
                DPPassed = True
            End If
            If DPPassed Then
                If IsNumeric(p.Substring(i, 1)) Then sb.Append(p.Substring(i, 1))
            End If

        Next


        Return sb.ToString


    End Function


    ''' <summary>
    ''' P.x=laengengrad/lon, p.y=breitengrad/lat
    ''' </summary>
    ''' <param name="p"></param>
    ''' <returns>GKZ</returns>
    Public Function GetGKZ(p As PointF) As Integer

        Dim gkz As Integer = _getgkz(p, Geom_Stmk)


        If gkz = GKZ_GRAZ And m_grazerbezirke Then


            gkz = _getgkz(p, Geom_Stmk)

            If gkz = 0 Then gkz = GKZ_GRAZER_BEZIRKE_UNBEK






        End If



        Return gkz





    End Function


    Private Function _getgkz(p As PointF, oGeoms As List(Of Geom)) As Integer
        For Each g As Geom In oGeoms
            If IsPointInPolygon4(g.Polygon, p) Then


                Return g.GKZ
            End If


        Next




        Return 0
    End Function

    Private Function IsPointInPolygon4(polygon As PointF(), testPoint As PointF) As Boolean
        Dim result As Boolean = False



        Dim j As Integer = polygon.Length() - 1
        For i As Integer = 0 To polygon.Length() - 1
            If polygon(i).Y < testPoint.Y AndAlso polygon(j).Y >= testPoint.Y OrElse polygon(j).Y < testPoint.Y AndAlso polygon(i).Y >= testPoint.Y Then
                If polygon(i).X + (testPoint.Y - polygon(i).Y) / (polygon(j).Y - polygon(i).Y) * (polygon(j).X - polygon(i).X) < testPoint.X Then
                    result = Not result
                End If
            End If
            j = i
        Next
        Return result
    End Function
End Class
Public Class Geom
    Public Polygon As System.Drawing.PointF()
    Public GKZ As Integer
End Class