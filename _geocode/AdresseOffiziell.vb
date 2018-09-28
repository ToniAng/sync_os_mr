Public Class AdresseOffiziell

    Private Const MIN_SCORE As Integer = 95
    Public Const BEZIRK_GRAZ_UNBEKANNT As Integer = 99

    Private Const FRM_ADR_BEARBEITEN As String = "adressen_bearbeiten"
    Private Const FRM_AERZTE As String = "frmaerzteliste"
    Private Const FRM_APO As String = "frmapotheke"
    Private Const FRM_SCHULEN As String = "schulen"


    Private db_con As cls_db_con


    Public Sub New()
        db_con = New cls_db_con
    End Sub

    

    Public Function GeocodeImport(Strasse As String, PLZ As String, Ort As String) As Adresse
        Dim adr As Adresse
        adr = New Adresse(Strasse, PLZ, Ort)

        Dim tb As DataTable = db_con.GetRecordset("select val from config where param='XT_CITY'")
        Dim GEOM_STMK = New GKZausGeom(IIf(tb.Rows(0)(0) <> 0, True, False))

        tb = db_con.GetRecordset("select val from config where param='GoogleKey'")

        Dim adrfull As Adresse = FullAdress(adr, tb.Rows(0)(0), False, GEOM_STMK)



        Return adr



    End Function


    Public Function FullAdress(Anfrage As Adresse, GoogleKey As String, GrazerBezirke As Boolean, gg As GKZausGeom) As Adresse

        Dim m_adresse As Adresse

        Dim gisstmk As New GISStmkHTTPGeocoder

        m_adresse = gisstmk.GetAdress(Anfrage)


        If m_adresse.GISScore < MIN_SCORE Then

            Dim goog As New GoogleHTTPGeocoder
            Dim msg As New Text.StringBuilder

            Dim res As GoogleGeocodeResult = goog.GetLatLon(Anfrage.Strasse, Anfrage.Hausnummer, Anfrage.PLZ, Anfrage.Ort, GoogleKey, 0)


            If res.Genauigkeit = GoogleHTTPGeocoder.GOOG_ACURRACY_ADDRESS And res.Breitengrad > 0 And res.Längengrad > 0 Then
                m_adresse.Lat = res.Breitengrad
                m_adresse.Lon = res.Längengrad
                m_adresse.StrasseUndHNr = res.StrasseUndHNr

                m_adresse.PLZ = res.PLZ
                m_adresse.Ort = res.Ort

                m_adresse.ID = res.ID
                m_adresse.GKZ = gg.GetGKZ(New Drawing.PointF(m_adresse.Lon, m_adresse.Lat))



                If m_adresse.GKZ = 0 Then m_adresse.GKZ = GetGKZ(Anfrage.Strasse, Anfrage.Ort, Anfrage.PLZ)


                m_adresse.ResultLevel = geocoder.ResultLevel.Address
            Else
                Anfrage.ResultLevel = geocoder.ResultLevel.NoResult
                'Anfrage.GKZ = oGHParam.GetGKZ(Anfrage.Strasse, Anfrage.Ort, Anfrage.PLZ)
                If GrazerBezirke Then
                    If Anfrage.GKZ = 60101 Then
                        Anfrage.GKZ = 6010199
                    End If
                End If
                Return Anfrage
            End If




        Else
            m_adresse.ResultLevel = geocoder.ResultLevel.Address
            If GrazerBezirke Then
                If m_adresse.GKZ = 60101 Then
                    m_adresse.GKZ = 60101 & m_adresse.BezirkGraz.ToString("00")
                End If
            End If



        End If



        UpdateAdressDB(m_adresse)
        Return m_adresse



    End Function


    Public Function GetGKZ(ByVal strasse As String, ByVal ort As String, ByVal plz As Integer, Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Integer
        'On Error GoTo gg_err
        Dim rs As DataTable
        'Dim db_con As New cls_db_con
        rs = db_con.GetRecordset("select gemeinden.dbo.getgkz_neu('" & strasse & "','" & ort & "', " & plz & ")", trans)
        If rs.Rows.Count > 0 Then
            GetGKZ = CInt(CStr(rs.Rows(0)(0)) & CStr(IIf(Me.IsXTCity(CLng(rs.Rows(0)(0))), "99", "")))
        Else
            GetGKZ = 0
        End If

exit_gg:
        On Error Resume Next
        rs = Nothing
        Exit Function

gg_err:
        GetGKZ = 0
        Resume exit_gg

    End Function



    Public Function IsXTCity(ByVal gemeindeid As Long) As Boolean
        'Dim db_con As New cls_db_con

        Dim tb As DataTable = db_con.GetRecordset("select val from config where param='XT_CITY'")

        If tb.Rows.Count = 0 Then
            Return False
        End If



        If tb.Rows(0)(0) = 0 Then
            Return False

        End If
        Dim l As List(Of Integer) = GetXT_Cities()

        For Each i As Integer In l
            If i = gemeindeid Then
                Return True
            End If

        Next
        Return False




    End Function
    Private Function GetXT_Cities() As List(Of Integer)

        Dim rs As DataTable

        'Dim db_con As New cls_db_con
        Dim l As New List(Of Integer)

        rs = db_con.GetRecordset("select gemeindeid from xt_cities")
        For Each r As DataRow In rs.Rows

            l.Add(r("gemeindeid"))


        Next

        Return l

    End Function

    Private Sub UpdateAdressDB(adr As Adresse)
        Return
        Dim ret As Integer = db_con.FireSQL("update ghdaten..wf_adressendb set adr_str_hnr='" & adr.StrasseUndHNr.Replace("'", "''") & "', adr_plz=" & adr.PLZ & ", adr_ort='" & adr.Ort.Replace("'", "''") & "' where adr_id='" & adr.ID & "'")
        If ret = 0 Then

            db_con.FireSQL("insert into ghdaten..wf_adressendb (adr_id, adr_str_hnr,adr_plz,adr_ort) values (" &
                           "'" & adr.ID & "'," &
                           "'" & adr.StrasseUndHNr.Replace("'", "''") & "'," &
                           adr.PLZ & "," &
                           "'" & adr.Ort.Replace("'", "''") & "')")

        End If


    End Sub


End Class








Public Class Adresse
    Public StrasseUndHNr As String

    Public PLZ As String

    Public Ort As String

    Public GKZ As Integer = 0

    Public Lat As Double
    Public Lon As Double

    Public GISScore As Double = 0

    Public ResultLevel As geocoder.ResultLevel

    Public ID As String

    Public BezirkGraz As Integer = AdresseOffiziell.BEZIRK_GRAZ_UNBEKANNT

    Private m_Strassenname As String

    Private m_Hausnummer As String

    Public ReadOnly Property LonKommaAsPoint As String
        Get
            Return Lon.ToString.Replace(",", ".")
        End Get
    End Property
    Public ReadOnly Property LatKommaAsPoint As String
        Get
            Return Lat.ToString.Replace(",", ".")
        End Get
    End Property


    Public Sub New(StrUndHnr As String, iPlz As Integer, sOrt As String)
        StrasseUndHNr = StrUndHnr
        PLZ = iPlz
        Ort = sOrt
    End Sub


    Public ReadOnly Property Strasse As String
        Get
            SetStrasse(StrasseUndHNr)
            Return m_STrassenname
        End Get
    End Property


    Public ReadOnly Property Hausnummer As String
        Get
            SetStrasse(StrasseUndHNr)
            Return m_Hausnummer
        End Get
    End Property





    Private Sub SetStrasse(ByVal txtStr As String)
        Dim i As Integer
        Dim Pos As Integer = 1
        m_STrassenname = txtStr
        m_Hausnummer = ""
        If txtStr = "" Then
            Exit Sub
        End If

        'Falls Strasse mit Ziffer anfängt
        If IsNumeric(Mid$(txtStr, 1, 1)) Then
            For i = 1 To Len(txtStr)
                If Not IsNumeric(Mid$(txtStr, i, 1)) Then Exit For
            Next
            Pos = i
            If Pos >= txtStr.Length Then Return
        End If


        For i = Pos To Len(txtStr)
            If IsNumeric(Mid$(txtStr, i, 1)) Then

                m_STrassenname = Trim(Left(txtStr, i - 1))
                m_Hausnummer = Trim(Right(txtStr, Len(txtStr) - i + 1))

                Return
            End If
        Next

    End Sub




End Class
