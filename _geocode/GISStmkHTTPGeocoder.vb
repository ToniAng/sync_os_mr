Imports System.Web.HttpUtility
Imports System.Text
Imports System.Web.Script.Serialization

Public Class GISStmkHTTPGeocoder




    Public Function GetAdress(Anfrage As Adresse) As Adresse

        'GIS-Stmk funktioniert nicht

        Anfrage.GISScore = 0
        Return Anfrage



        If Anfrage.Hausnummer = "" Or Anfrage.Strasse = "" Or Anfrage.PLZ = "" Or Anfrage.Ort = "" Or Anfrage.PLZ = 9999 Or Anfrage.Strasse = "unbekannt" Or Anfrage.Ort = "unbekannt" Then
            Anfrage.GISScore = 0
            Return Anfrage
        End If





        Dim HRef As String = "https://gis.stmk.gv.at/solr/gisstmkadr/search/adressenWgs84/?q=" &
            HtmlEncode(Anfrage.Strasse) &
            "+" & HtmlEncode(Anfrage.Hausnummer) &
            "+" & Anfrage.PLZ &
            "+" & HtmlEncode(Anfrage.Ort)




        Dim Request As System.Net.HttpWebRequest
        Dim Response As System.Net.HttpWebResponse

        Dim ResponseStream As System.IO.Stream

        Dim enc As Encoding = Encoding.GetEncoding("utf-8")





        Request = CType(System.Net.HttpWebRequest.Create(HRef),
                            System.Net.HttpWebRequest)
        Request.ContentType = "text/xml"

        response = CType(Request.GetResponse(), System.Net.HttpWebResponse)
        ResponseStream = response.GetResponseStream()
        Dim reader As IO.StreamReader = New IO.StreamReader(ResponseStream, enc)
        Dim responseFromServer As String = reader.ReadToEnd()

        reader.Close()
        response.Close()

        Dim js As New JavaScriptSerializer()
        Dim ret As GisStmkErgebnis = js.Deserialize(Of GisStmkErgebnis)(responseFromServer)

        If ret.response.docs.Count = 0 Then
            Anfrage.GISScore = 0
            Return Anfrage
        End If

        Dim treffer As Doc = ret.response.docs(0)
        Dim l As Integer = treffer.title(0).LastIndexOf(",")

        Dim StrHnr As String = treffer.title(0).Substring(0, l)
        Dim PLZ = treffer.title(0).Replace(StrHnr, "").Replace(",", "").Trim.Split(" ")(0)
        Dim Ort As String = treffer.title(0).Replace(StrHnr, "").Replace(",", "").Replace(PLZ, "").Trim

        'Dim Adr As String() = treffer.title(0).Split(",")




        'Dim PLZOrt As String() = treffer.title(0).Replace(StrHnr, "").Replace(",", "").Trim.Split(" ")

        Dim AdressErg As New Adresse(StrHnr, PLZ, Ort)


        AdressErg.GISScore = treffer.score

        AdressErg.GKZ = treffer.subtext.Split(":")(1).Trim


        AdressErg.ID = treffer.id

        Dim LatLon As String() = treffer.geo(0).Replace("Point", "").Replace("(", "").Replace(")", "").Trim.Split(" ")

        AdressErg.Lat = LatLon(0).Replace(".", ",")
        AdressErg.Lon = LatLon(1).Replace(".", ",")


        If AdressErg.GKZ = 60101 Then AdressErg.BezirkGraz = treffer.extrasearch.Substring(5, 2)











        Return AdressErg








    End Function



End Class



Public Class Doc
    Public subtext As String
    Public maxy As Integer
    Public maxx As Integer
    Public title As String()
    Public type As String
    Public thumbnail_url As String
    Public geo As String()
    Public textsuggest As String
    Public extradisplay As String
    Public miny As Integer
    Public minx As Integer
    Public popularity As Integer
    Public action As String
    Public id As Int64
    Public extrasearch As String
    Public _version_ As String
    Public timestamp As String
    Public score As Double
End Class


Public Class responseHeader

    Public status As Integer
    Public QTime As Integer



End Class


Public Class params
    Public q As String
    Public pt As String
    Public wt As String

End Class


Public Class response
    Public numFound As Integer
    Public start As Integer
    Public maxScore As Double
    Public docs As List(Of Doc)
End Class


Public Class GisStmkErgebnis
    Public responseHeader As responseHeader
    Public response As response

End Class



'{
'responseHeader: {
'status: 0,
'QTime: 2,
'params: {
'q: "katzianergasse 10 8010 graz"
'}
'},
'response: {
'   numFound: 180385,
'   start: 0,
'   maxScore: 13.805061,
'   docs: [
'       {
'           subtext: "Adresse in der Gemeinde Graz GKZ: 60101",
'           maxy: 5212881,
'           maxx: 534491,
'           title: [
'               "Katzianergasse 10, 8010 Graz"
'           ],
'           type: "AdressenWGS84",
'           thumbnail_url: "http://gis.stmk.gv.at/content/dokumente/img/search/adressen.png",
'           geo: [
'               "Point (47.06854100 15.45427100)"
'           ],
'           textsuggest: "Katzianergasse 10 8010 Graz",
'           extradisplay: "Lat: 47.06854100N Long: 15.45427100E",
'           miny: 5212881,
'           minx: 534491,
'           popularity: 0,
'           action: "http://gis2.stmk.gv.at/atlas2/route.html?s=Katzianergasse 10, 8010 Graz",
'           id: "5737683001",
'           extrasearch: "Graz,02.Bez.:Sankt Leonhard Katzianergasse 10 8010 Graz 5737683 Graz,02.Bez.:St.Leonhard Katzianerg.",
'           _version_: 1597895832532680700,
'           timestamp: "2018-04-16T09:50:15.779Z",
'           score: 13.805061
'       },
'       {
'           subtext: "Adresse in der Gemeinde Graz GKZ: 60101",
'           maxy: 5212861,
'maxx: 534507,
'title: [
'"Katzianergasse 12, 8010 Graz"
'],
'type: "AdressenWGS84",
'thumbnail_url: "http://gis.stmk.gv.at/content/dokumente/img/search/adressen.png",
'geo: [
'"Point (47.06836700 15.45446900)"
'],
'textsuggest: "Katzianergasse 12 8010 Graz",
'extradisplay: "Lat: 47.06836700N Long: 15.45446900E",
'miny: 5212861,
'minx: 534507,
'popularity: 0,
'action: "http://gis2.stmk.gv.at/atlas2/route.html?s=Katzianergasse 12, 8010 Graz",
'id: "5737685001",
'extrasearch: "Graz,02.Bez.:Sankt Leonhard Katzianergasse 12 8010 Graz 5737685 Graz,02.Bez.:St.Leonhard Katzianerg.",
'_version_: 1597895832531632000,
'timestamp: "2018-04-16T09:50:15.778Z",
'score: 9.692781
'},
'{
'subtext: "Adresse in der Gemeinde Graz GKZ: 60101",
'maxy: 5212883,
'maxx: 534523,
'title: [
'"Katzianergasse 11, 8010 Graz"
'],
'type: "AdressenWGS84",
'thumbnail_url: "http://gis.stmk.gv.at/content/dokumente/img/search/adressen.png",
'geo: [
'"Point (47.06856000 15.45469100)"
'],
'textsuggest: "Katzianergasse 11 8010 Graz",
'extradisplay: "Lat: 47.06856000N Long: 15.45469100E",
'miny: 5212883,
'minx: 534523,
'popularity: 0,
'action: "http://gis2.stmk.gv.at/atlas2/route.html?s=Katzianergasse 11, 8010 Graz",
'id: "5737684001",
'extrasearch: "Graz,02.Bez.:Sankt Leonhard Katzianergasse 11 8010 Graz 5737684 Graz,02.Bez.:St.Leonhard Katzianerg.",
'_version_: 1597895832532680700,
'timestamp: "2018-04-16T09:50:15.779Z",
'score: 9.692781
'},
'{
'subtext: "Adresse in der Gemeinde Graz GKZ: 60101",
'maxy: 5212898,
'maxx: 534508,
'title: [
'"Katzianergasse 9, 8010 Graz"
'],
'type: "AdressenWGS84",
'thumbnail_url: "http://gis.stmk.gv.at/content/dokumente/img/search/adressen.png",
'geo: [
'"Point (47.06870200 15.45448500)"
'],
'textsuggest: "Katzianergasse 9 8010 Graz",
'extradisplay: "Lat: 47.06870200N Long: 15.45448500E",
'miny: 5212898,
'minx: 534508,
'popularity: 0,
'action: "http://gis2.stmk.gv.at/atlas2/route.html?s=Katzianergasse 9, 8010 Graz",
'id: "5737682001",
'extrasearch: "Graz,02.Bez.:Sankt Leonhard Katzianergasse 9 8010 Graz 5737682 Graz,02.Bez.:St.Leonhard Katzianerg.",
'_version_: 1597895832532680700,
'timestamp: "2018-04-16T09:50:15.779Z",
'score: 9.692781
'},
'{
'subtext: "Adresse in der Gemeinde Graz GKZ: 60101",
'maxy: 5212893,
'maxx: 534480,
'title: [
'"Katzianergasse 8, 8010 Graz"
'],
'type: "AdressenWGS84",
'thumbnail_url: "http://gis.stmk.gv.at/content/dokumente/img/search/adressen.png",
'geo: [
'"Point (47.06865300 15.45411700)"
'],
'textsuggest: "Katzianergasse 8 8010 Graz",
'extradisplay: "Lat: 47.06865300N Long: 15.45411700E",
'miny: 5212893,
'minx: 534480,
'popularity: 0,
'action: "http://gis2.stmk.gv.at/atlas2/route.html?s=Katzianergasse 8, 8010 Graz",
'id: "5737681001",
'extrasearch: "Graz,02.Bez.:Sankt Leonhard Katzianergasse 8 8010 Graz 5737681 Graz,02.Bez.:St.Leonhard Katzianerg.",
'_version_: 1597895832532680700,
'timestamp: "2018-04-16T09:50:15.779Z",
'score: 9.692781
'}
']
'}
'}