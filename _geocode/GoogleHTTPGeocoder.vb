Imports System.Text

Public Class GoogleHTTPGeocoder

    Public Längengrad As Double
    Public Breitengrad As Double
    Public Höhe As Double
    Public ReturnCode As Integer
    'Public Genauigkeit As GooogleGeocodeAccuracy

    Public TooManyQueries As Boolean = False
    Public BadKey As Boolean = False

    Public Const MAX_QUERIES As Integer = 2500 'Begrenzung Google Geocoder
    Private Const GOOGLE_WAIT_TOOMANYQUERIES As Integer = 60000 'Zeit zwischen Abfragen in ms




    'Public Enum GooogleGeocodeAccuracy
    '    Unbekannter_Ort = 0
    '    Land = 1
    '    Bundesland = 2
    '    Bezirk = 3
    '    Ortschaft = 4
    '    Postleitzahl = 5
    '    Straße = 6
    '    Kreuzung = 7
    '    Adresse = 8
    '    Grundstück = 9
    'End Enum

    Private Const A_ACCURACY As String = "location_type"
    Private Const T_ADDR_DETAILS As String = "formatted_address"
    Private Const T_CODE As String = "status"
    Private Const T_COORD As String = "location"
    Private Const T_STRASSE_HNR As String = "ThoroughfareName"
    Private Const T_PLZ As String = "PostalCodeNumber"

    Private Const T_LAT As String = "lat"
    Private Const T_LONG As String = "lng"
    Private Const T_ADDR As String = "formatted_address"
    Private Const T_ADRESS_ID As String = "place_id"


    Private Const NO_RES As String = "ZERO_RESULTS"

    Public Const GOOG_ACURRACY_ADDRESS As String = "ROOFTOP"
    Public Const GOOG_ACURRACY_STEET As String = "RANGE_INTERPOLATED"
    Public Const GOOG_ACURRACY_CENTER As String = "GEOMETRIC_CENTER"
    Public Const GOOG_ACURRACY_APPROXIMATE As String = "APPROXIMATE"



    Private Const G_GEO_SUCCESS As String = "OK" 'Keine Fehler aufgetreten; die Adresse wurde erfolgreich analysiert. Der Geocode wurde zurückgegeben. (Seit 2.55) 
    Private Const G_GEO_BAD_REQUEST As String = "REQUEST_DENIED" 'Eine Routenanforderung konnte nicht erfolgreich analysiert werden. (Seit 2.81) 
    'Private Const G_GEO_SERVER_ERROR As Integer = 500 'Eine Geokodierungs- oder Routenanforderung konnte nicht erfolgreich verarbeitet werden, da der genaue Grund für den Fehler nicht bekannt ist. (Seit 2.55) 
    Private Const G_GEO_MISSING_QUERY As String = "INVALID_REQUEST" 'Der HTTP-Parameter q fehlt oder enthält keinen Wert. Für Geokodierungsanforderungen bedeutet dies, dass eine leere Adresse angegeben wurde. Für Routenanforderungen bedeutet dies, dass keine Abfrage angegeben wurde. (Seit 2.81) 
    'Private Const G_GEO_MISSING_ADDRESS As Integer = 601 'Synonym für G_GEO_MISSING_QUERY. (Seit 2.55) 
    Private Const G_GEO_UNKNOWN_ADDRESS As String = "ZERO_RESULTS" 'Es konnte keine entsprechende geografische Position für die angegebene Adresse gefunden werden. Dies kann daran liegen, dass die Adresse relativ neu oder möglicherweise falsch ist. (Seit 2.55) 
    'Private Const G_GEO_UNAVAILABLE_ADDRESS As Integer = 603 'Der Geocode für die angegebene Adresse oder die Route für die angegebene Richtungsanfrage kann aus rechtlichen oder Vertragsgründen nicht zurückgegeben werden. (Seit 2.55) 
    'Private Const G_GEO_UNKNOWN_DIRECTIONS As Integer = 604 'Das GDirections-Objekt konnte keinen Routenplan zwischen den Punkten in der Suchanfrage berechnen. Dies ist üblich, da es keine Route zwischen den beiden Punkten gibt oder keine Daten für die Routenplanung in dieser Region vorhanden sind. (Seit 2.81) 
    'Public Const G_GEO_BAD_KEY As Integer = 610 'Der angegebene Schlüssel ist entweder ungültig oder passt nicht zur Domain, für die er angegeben wurde. (Seit 2.55) 
    Public Const G_GEO_TOO_MANY_QUERIES As String = "OVER_QUERY_LIMIT" 'Der angegebene Schlüssel hat das Anforderungslimit innerhalb der 24-Stunden-Frist überschritten. (Seit 2.55) 



    Public Function GetLatLon(ByVal Str As String, _
                    ByVal HNr As String, _
                    ByVal PLZ As String, _
                    ByVal Ort As String, _
                    ByVal GoogleID As String, _
                    ByVal DS_ID As String) As GoogleGeocodeResult






        Dim Request As System.Net.HttpWebRequest
        Dim Response As System.Net.HttpWebResponse
        'Dim bytes() As Byte
        'Dim RequestStream As System.IO.Stream
        Dim ResponseStream As System.IO.Stream
        Dim ResponseXmlDoc As System.Xml.XmlDocument
        Dim RES As New GoogleGeocodeResult

        Dim geocode_pending = True
        Dim delay = 200

        Dim HRef As String


        'If IsTestEnvironment() Then

        '    RES.Genauigkeit = GoogleHTTPGeocoder.GOOG_ACURRACY_ADDRESS
        '    RES.Längengrad = 15.4520877
        '    RES.Breitengrad = 47.0685389
        '    RES.StrasseUndHNr = "Katzianergasse 10"
        '    RES.PLZ = 8010
        '    RES.Ort = "Graz"
        '    Return RES

        'End If


        'If HNr = "" Or Str = "" Or PLZ = "" Or Ort = "" Then
        '    Throw New Exception("Adressangabe ist unvollständig.")
        'End If

        If HNr = "" Or Str = "" Or PLZ = "" Or Ort = "" Or PLZ = 9999 Or Str = "unbekannt" Or Ort = "unbekannt" Then
            RES.Genauigkeit = geocoder.ResultLevel.NoResult
            Return RES
        End If


        HRef = "https://maps.googleapis.com/maps/api/geocode/xml?address=" & Str & "+" & HNr & ",+" & PLZ & "+" & Ort & "+,Austria&sensor=false"

        If Not String.IsNullOrEmpty(GoogleID) Then
            HRef = HRef & "&key=" & GoogleID
        End If

        Dim enc As Encoding = Encoding.GetEncoding("utf-8")

        'Do While geocode_pending



        Request = CType(System.Net.HttpWebRequest.Create(HRef),
                            System.Net.HttpWebRequest)
        Request.ContentType = "text/xml"

        Response = CType(Request.GetResponse(), System.Net.HttpWebResponse)
        ResponseStream = Response.GetResponseStream()
        Dim reader As IO.StreamReader = New IO.StreamReader(ResponseStream, enc)
        Dim responseFromServer As String = reader.ReadToEnd()



        'If responseFromServer.ToLower.IndexOf("klagenf") >= 0 Then
        '    Dim xx As Int16
        '    xx = 1

        'End If

        reader.Close()
        Response.Close()

        If responseFromServer.IndexOf("OVER_QUERY_LIMIT") >= 0 Then

            Throw New Exception("Google Account erlaubt keine weiteren Abfragen.")

        End If


        'If responseFromServer.IndexOf(NO_RES, 0, StringComparison.CurrentCultureIgnoreCase) >= 0 Then

        'End If

        ResponseXmlDoc = New System.Xml.XmlDocument

        ResponseXmlDoc.LoadXml(responseFromServer)

        Dim NL As System.Xml.XmlNodeList

        NL = ResponseXmlDoc.GetElementsByTagName(T_CODE)

        RES.Breitengrad = 0
        RES.Längengrad = 0
        Try
            RES.Genauigkeit = ResponseXmlDoc.GetElementsByTagName(A_ACCURACY)(0).InnerText

        Catch ex As Exception

        End Try
        RES.Höhe = 0
        RES.ReturnCode = NL(0).InnerText




        Select Case NL(0).InnerText
            Case G_GEO_SUCCESS



                RES.ID = ResponseXmlDoc.GetElementsByTagName(T_ADRESS_ID)(0).InnerText

                RES.Längengrad = ResponseXmlDoc.GetElementsByTagName(T_LONG)(0).InnerText.Replace(".", ",")

                RES.Breitengrad = ResponseXmlDoc.GetElementsByTagName(T_LAT)(0).InnerText.Replace(".", ",")

                Dim stradr As String = ResponseXmlDoc.GetElementsByTagName(T_ADDR)(0).InnerText.Replace(", Austria", "").Replace(", Österreich", "")
                Dim l As Integer = stradr.LastIndexOf(",")

                Dim retPLZ As Integer = 0
                Dim retOrt As String = ""
                Dim StrHnr As String = ""
                If l >= 0 Then
                    StrHnr = stradr.Substring(0, l)
                    Try

                        retPLZ = stradr.Replace(StrHnr, "").Replace(",", "").Trim.Split(" ")(0)
                    Catch ex As Exception
                        retPLZ = 9999
                    End Try
                    Try
                        retOrt = stradr.Replace(StrHnr, "").Replace(",", "").Replace(retPLZ, "").Trim

                    Catch ex As Exception
                        retOrt = "unbekannt"
                    End Try
                    If retOrt = "" Then

                        Dim tmp_a As New Adresse(StrHnr, retPLZ, "")

                        retOrt = tmp_a.Strasse


                    End If

                Else
                    If IsNumeric(stradr.Substring(0, 4)) Then
                        retPLZ = stradr.Trim.Split(" ")(0)
                    End If

                    retOrt = stradr.Replace(retPLZ, "").Trim

                End If





                'Dim PLZOrt As String() = adr(1).Trim.Split(" ")

                RES.StrasseUndHNr = StrHnr
                RES.PLZ = retPLZ
                RES.Ort = retOrt

                geocode_pending = False



            Case G_GEO_TOO_MANY_QUERIES
                'Log("Code 620 - Zu viele Abfragen")

                RES.KontingentAusgeschöpft = True
                Throw New Exception("Lokalisierung nicht möglich - Google-Kontingent ist ausgeschöpft.")





            Case Else


                RES.Genauigkeit = geocoder.ResultLevel.NoResult

        End Select

        'Loop


        Return RES



    End Function






End Class
Public Class GoogleGeocodeResult
    Public Längengrad As Double
    Public Breitengrad As Double
    Public Höhe As Double
    Public ReturnCode As String
    Public Genauigkeit As String
    Public KontingentAusgeschöpft As Boolean = False
    Public StrasseUndHNr As String
    Public PLZ As Integer
    Public Ort As String
    Public ID As String

End Class



'<?xml version="1.0" encoding="UTF-8" ?> 
'- <kml xmlns="http://earth.google.com/kml/2.0">
'- <Response>
'  <name>Katzianergasse 10,8010 Graz,Austria</name> 
'- <Status>
'  <code>200</code> 
'  <request>geocode</request> 
'  </Status>
'- <Placemark id="p1">
'  <address>Katzianergasse 10, 8010 Graz, Österreich</address> 
'- <AddressDetails Accuracy="8" xmlns="urn:oasis:names:tc:ciq:xsdschema:xAL:2.0">
'- <Country>
'  <CountryNameCode>AT</CountryNameCode> 
'  <CountryName>Österreich</CountryName> 
'- <AdministrativeArea>
'  <AdministrativeAreaName>Steiermark</AdministrativeAreaName> 
'- <SubAdministrativeArea>
'  <SubAdministrativeAreaName>Graz (Stadt)</SubAdministrativeAreaName> 
'- <Locality>
'  <LocalityName>Graz</LocalityName> 
'- <DependentLocality>
'  <DependentLocalityName>Graz,02.Bez.:Sankt Leonhard</DependentLocalityName> 
'- <Thoroughfare>
'  <ThoroughfareName>Katzianergasse 10</ThoroughfareName> 
'  </Thoroughfare>
'- <PostalCode>
'  <PostalCodeNumber>8010</PostalCodeNumber> 
'  </PostalCode>
'  </DependentLocality>
'  </Locality>
'  </SubAdministrativeArea>
'  </AdministrativeArea>
'  </Country>
'  </AddressDetails>
'- <Point>
'  <coordinates>15.4542732,47.0686218,0</coordinates> 
'  </Point>
'  </Placemark>
'  </Response>
'  </kml>