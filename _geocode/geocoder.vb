Public Class geocoder


    Private goog_src As New GoogleHTTPGeocoder
    Private m_country As Integer

    'Public m_googlekey As String

    Private Const GOOGLE_WAIT As Integer = 200 'Zeit zwischen Abfragen in ms





    Public Enum ResultLevel
        Address = 0
        Street = 1
        PLZ = 2
        City = 3
        NoResult = 4
        DoNotRecode = 9
    End Enum

    Public Enum AreaLevels
        NUTS0 = 0
        NUTS1 = 1
        NUTS2 = 2
        NUTS3 = 3
        BEZIRK = 4
        GEMEINDE = 5
        KATATRALGEMEINDE = 6
    End Enum

    Public KGList As DataTable
    Public KG As New ArrayList

    Public Sub GetCode(str As String, plz As Integer, ort As String, GoogleID As String)

        Dim strHNR As String
        Dim strStreet As String

        Dim m_objUtil As New GeocoderUtil




        strStreet = m_objUtil.GetStreet(str, ort, 0, False)
        strHNR = m_objUtil.GetHNR()



        If Search(strStreet, strHNR, ort, plz, 0, False, GoogleID) Then
            'Kontingent ausgeschöpft
            Return
        End If


        'If Result <> ResultLevel.NoResult Then
        '    m_lon = m_objCoder.Laengengrad
        '    m_lat = m_objCoder.Breitengrad
        'End If

        'If Save Then SaveCode(False)
        'Return False




    End Sub

    Private Function Search(ByVal strStreet As String,
                        ByVal strHNr As String,
                        ByVal strCity As String,
                        ByVal PLZ As String,
                        ByVal Gemeindeid As Integer,
                        ByVal StreetFromOrig As Boolean,
                        ByVal GoogleID As String) As Boolean





        lat = 0
        lon = 0
        Result = ResultLevel.NoResult



        Try




            Dim res As GoogleGeocodeResult = goog_src.GetLatLon(strStreet, strHNr, PLZ, strCity, GoogleID, 0)

            If res.KontingentAusgeschöpft Then Return True

            Select Case res.Genauigkeit
                Case GoogleHTTPGeocoder.GOOG_ACURRACY_ADDRESS
                    Result = ResultLevel.Address

                Case GoogleHTTPGeocoder.GOOG_ACURRACY_STEET
                    Result = ResultLevel.Street
                Case Else
                    Result = ResultLevel.NoResult

                    Return False



            End Select

            lat = res.Breitengrad
            lon = res.Längengrad



        Catch ex As Exception

            Result = ResultLevel.NoResult

            Return False
        Finally


        End Try


        Return True


    End Function


    Public Property lat As Double 'Y-Achse
    Public Property lon As Double 'X-Achse
    Public Property Result As ResultLevel

End Class
