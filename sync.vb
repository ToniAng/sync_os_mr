Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Collections.Specialized

Imports System.ComponentModel
Imports System.Web.Script.Serialization
Imports VB = Microsoft.VisualBasic

Public Class sync
    Private Const TEST_SERVER As String = "proximacentauri"
    Private Const TEST_SERVER2 As String = "70-Ophiuchi"


    Public Const NO_PLZ As Integer = 9999
    Public Const HEFTBASIS As Long = 90110000
    Public Const NO_ADDRESS As String = "unbekannt"
    Public Const NO_VAL As Integer = -1
    Public Const FA_ID As Integer = 9000093


    Private PARAM_LAST_VACC_SYNC As String = "LastVaccSyncOS"




    Private Enum DS_Typ
        Impfung = 0
        Impfling = 1
    End Enum


    Private Enum Programme
        Kein_Impfprogramm = 0
        MKP = 1
        Impfnetzwerk = 2
        Schulimpfungen = 3
        Nachholbon = 4
        PNC_Impfungen_Land = 5
        MMR_Nachholbon = 6
        Mammographiescreening = 7
        Magistratsimpfungen = 9
        Adipositas = 10
        ErwachsenenImpfung = 11
        FA_Impfungen = 12

    End Enum
    Public Enum AktionslogKat
        Allgemein = 0
        TeilnehmerInnen_löschen = 1
        Kinderdatensätze_löschen = 2
        Impfbestellungen_löschen = 3
        Impfdoku_löschen = 4
        Untersuchungen_bearbeiten_löschen = 5
        Heftausgaben_löschen = 6
        Ersatzheft = 7
        Heftrückforderung_Doppelausgabe = 8
        Datenübermittlung = 9
        Einstellungsänderungen = 10
        Druck_Gutscheinheft = 11
        Import_Kennwerte = 12
        Erstellung_Datawarehouse = 13
        Änderungslog = 14
        Scanvorgang_ändern = 15
        Dokumentmanipulationen = 16
        Dokumentlog = 17
        Integration_BH_Impfungen = 18

    End Enum

    Public Enum ConfigParam
        Benachrichtigung_AbrechnungArzte = 0
        Benachrichtigung_BHAufstellung = 1
        WO_AbrechnungErstellenLäuft = 2
        LetzteAccountPflegeMeldung = 3
        Allgemeine_EMail_Benachrichtigung_Online_Service = 4
        Allgemeine_EMail_Benachrichtigung_Online_Service_Texte = 5
        Anmerkungen_Impfststaus_Online = 6
        LastTask = 7
        DeleteTask = 8
        AbrechnungErstellen = 9
        SchwelleEinzugsgebiet = 10
        GoogleKey = 11
        AccouintInfo = 12
        StandGHDaten = 13
        Abrechnungerneuern = 14
        UploadOnlineRecherche = 15
        DeleteOnlineRecherche = 16
        OnlineRecherchen = 17
        RealDeleteOnlineRecherche = 18
        RechercheDownloaded = 19
        DownloadRechercheData = 20
        ResetDownloadMarkerOnlineRecherchen = 21
        SendUrgenzmailOnlinerecherchen = 22
        SyncDoc = 23
        ValidateDocSync = 24
        SyncVaccData = 25


    End Enum

    Private Sub AddSynclog(dstyp As DS_Typ, dsid As Integer, Optional dsid_tmp As Integer = 0, Optional trans As SqlClient.SqlTransaction = Nothing, Optional boncode As String = "")

        Dim db_con As New cls_db_con
        Dim tmp As String = IIf(dsid_tmp = 0, "NULL", dsid_tmp)
        Dim bc As String = IIf(boncode = "", "NULL", boncode)
        db_con.FireSQL("insert into os_synclog (os_dsid,os_dsid_tmp, os_dstyp, os_boncode) values (" & dsid & "," & tmp & "," & CInt(dstyp) & "," & bc & ")", trans)


    End Sub

    Private Function GetPersistentID(tmp_dsid As Integer, dstyp As DS_Typ, Optional trans As SqlClient.SqlTransaction = Nothing) As Integer

        Dim db_con As New cls_db_con
        Dim tb As DataTable = db_con.GetRecordSet("select os_dsid from os_synclog where os_dsid_tmp=" & tmp_dsid & " and os_dstyp=" & CInt(DS_Typ.Impfling), trans)
        If tb.Rows.Count > 0 Then
            Return tb.Rows(0)("os_dsid")
        Else
            Return tmp_dsid
        End If


    End Function
    Private Sub SendAdminmail(ByVal Msg As String, ByVal Subject As String)


        Console.WriteLine(Msg)

        Dim x As New System.Net.Mail.SmtpClient
        'x.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials



        'Dim basicAuthenticationInfo As New System.Net.NetworkCredential("administrator", "huekuvlof")
        'Put your own, or your ISPs, mail server name onthis next line
        'x.UseDefaultCredentials = False
        'x.Credentials = basicAuthenticationInfo


        x.Host = Setting.SMTP_SERVER_ALTERNATIV





        x.Send("sync_os-mr@vorsorgemedizin.st",
                Setting.EMail,
               Subject,
               Msg)

    End Sub

    Public Sub SyncVacc()

        If SyncLog() Then Return



        Dim ret As String = ""
        Dim ds As New DataSet
        Dim db_con As New cls_db_con
        Dim trans As SqlClient.SqlTransaction = Nothing
        Dim con As SqlClient.SqlConnection = db_con.GetCon(False)




        Try
            Dim s As String = GetRemoteParam(URL_Online_Service & "/ConfigEx/GetOSVaccData")

            If s = "false" Then
                AktionsLog("Daten für den Abgleich der Impfungen konnten nicht abgerufen werden.", AktionslogKat.Integration_BH_Impfungen)
                Return
            ElseIf s = "" Then
                AktionsLog("Keine Daten vom Online-Service (Down?)- bitte prüfen.", AktionslogKat.Integration_BH_Impfungen)
                Return
            End If
            Dim js As New JavaScriptSerializer()
            Dim osv As OSVaccData = js.Deserialize(Of OSVaccData)(s)





            trans = con.BeginTransaction







            Dim curPat As OSImpfling
            For Each curPat In osv.pers
                curPat.heftnr_os = curPat.heftnr
                If curPat.heftnr < 0 Then
                    curPat.heftnr = GetPersistentID(curPat.heftnr, DS_Typ.Impfling, trans)
                End If
                If curPat.heftnr < 0 Then

                    AddSynclog(DS_Typ.Impfling, SetPatData(curPat, trans), curPat.heftnr_os, trans)
                Else
                    SetPatData(curPat, trans)
                    AddSynclog(DS_Typ.Impfling, curPat.heftnr,, trans)
                End If
            Next



            For Each vacc As Idata In osv.Vacc

                curPat = GetCurPat(vacc, trans)
                AddSynclog(DS_Typ.Impfung, vacc.dsid, , trans, savevacc(vacc, curPat, trans))

            Next
            If osv.Vacc.Count > 0 Then
                AktionsLog(osv.Vacc.Count & " Impfeinträge wurden synchronisiert.", AktionslogKat.Integration_BH_Impfungen, trans)
            End If

            trans.Commit()
        Catch ex As Exception
            Try
                trans.Rollback()
            Catch ex2 As Exception

            End Try
            'AktionsLog("Fehler: " & ex.Message & vbNewLine & ex.StackTrace, AktionslogKat.Integration_BH_Impfungen)
            LogErr(ex, db_con.LastErrSQL)
        Finally

        End Try





        If SyncLog() Then Return


    End Sub
    'Private Function GetNewPatData(plist As List(Of OSImpfling)) As List(Of OSImpfling)
    '    Dim ret As New List(Of OSImpfling)
    '    For Each i As OSImpfling In plist
    '        If i.newpat <> 0 Then ret.Add(i)
    '    Next

    '    Return ret
    'End Function

    Private Function SyncLog() As Boolean
        Dim db_con As New cls_db_con
        Try


            Dim pers As DataTable = db_con.GetRecordSet("select * from os_synclog where os_commit=0 and os_dstyp=" & CInt(DS_Typ.Impfling))
            Dim vacc As DataTable = db_con.GetRecordSet("select * from os_synclog where os_commit=0 and os_dstyp=" & CInt(DS_Typ.Impfung))
            pers.TableName = "pers"
            vacc.TableName = "vacc"


            If pers.Rows.Count = 0 And vacc.Rows.Count = 0 Then Return False


            Dim ds As New DataSet
            ds.Tables.Add(pers)
            ds.Tables.Add(vacc)



            Dim ret As Boolean = SetRemoteParam(URL_Online_Service & "/configex/SetParam", ConfigParam.SyncVaccData, ToBase64(ds.GetXml), False)
            If ret Then
                Try

                    AktionsLog("Fehler beim Loggen der bereits integrierten Impfdaten im Online-Service.", AktionslogKat.Integration_BH_Impfungen)
                Catch ex As Exception

                End Try
                Return True
            End If
            db_con.FireSQL("update os_synclog set os_commit=1 where os_commit=0")
            SetLastSynTime()

        Catch ex As Exception
            Try

                'AktionsLog("Fehler: " & ex.Message & vbNewLine & ex.StackTrace, AktionslogKat.Integration_BH_Impfungen)
                LogErr(ex, db_con.lasterrsql)
            Catch ex2 As Exception

            End Try

            Return True
        End Try

        Return False
    End Function


    Private Sub SetLastSynTime()
        Dim db_con As New cls_db_con
        Dim ret As Integer = db_con.FireSQL("update ghdaten..config set val='" & Date.Now & "' where param='" & PARAM_LAST_VACC_SYNC & "'")

        If ret = 0 Then
            db_con.FireSQL("insert into ghdaten..config (param, val) values ('" & PARAM_LAST_VACC_SYNC & "','" & Date.Now & "')")
        End If

    End Sub

    'Public Function ToBase64(ByVal sText As String) As String
    '    Return System.Convert.ToBase64String(System.Text.Encoding.Default.GetBytes(sText))
    'End Function
    Private Function savevacc(vacc As Idata, pat As OSImpfling, trans As SqlClient.SqlTransaction) As String
        Dim db_con As New cls_db_con
        Dim strSQL As String
        Dim Charge = "NULL"
        If Not String.IsNullOrEmpty(vacc.charge) Then Charge = "'" & vacc.charge & "'"



        Dim standort = "NULL"
        Dim standortid = "NULL"
        If Not String.IsNullOrEmpty(vacc.standort) Then standort = "'" & vacc.standort & "'"
        If Not String.IsNullOrEmpty(vacc.standortid) Then standortid = "'" & vacc.standortid & "'"

        If vacc.heftnr < 0 Then

            vacc.heftnr = GetPersistentID(vacc.heftnr, DS_Typ.Impfling, trans)
        End If

        If vacc.heftnr <= 0 Then
            Throw New Exception("Kann persistente Heftnr für " & pat.nn & " " & pat.vn & " (ID tmp.: " & vacc.heftnr & ") nicht eruieren.")
            Return ""
        End If


        Dim bc As String = GetBoncode(vacc.heftnr, vacc.serum, vacc.impfung)

        If String.IsNullOrEmpty(vacc.plid) Then





            db_con.FireSQL("insert into ghdaten..impfdoku  (" &
                     "datum,nname,vname,gebdat,chargenr," &
                     "arztnr,serum,bis6,eingang,geandert," &
                     "geandertam,satz,mwsthon,aposatz,apomwst," &
                     "heftnr,boncode,impfung,NoBilling,RsnNoBilling," &
                     "Prog,bhnr,standort,standortid) values (" &
                     "'" & CDate(vacc.datum).ToShortDateString & "'," &
                     "'" & pat.nn & "'," &
                     "'" & pat.vn & "'," &
                     "'" & pat.gebdat & "'," &
                     Charge & "," &
                     vacc.arztnr & "," &
                     vacc.serum & "," &
                     IIf(bis6(pat.gebdat, CDate(vacc.datum)), -1, 0) & "," &
                     "'" & Date.Today & "'," &
                     "'OS_" & vacc.geandert & "'," &
                     "'" & vacc.geandertam & "'," &
                     "0,0,0,0," &
                     vacc.heftnr & ", " &
                     bc & "," &
                     vacc.impfung & "," &
                     "-1,'BH-Impfung, Online'," &
                     vacc.prog & "," &
                     vacc.bhnr & "," &
                     standort & "," &
                     standortid &
                     ")", trans)

        Else
            Dim s As String() = vacc.plid.Split("|")

            If vacc.del <> 0 Then
                strSQL = "delete ghdaten..impfdoku where datum='" & s(0) & "' and boncode=" & s(1)
                db_con.FireSQL(strSQL, trans)
                AktionsLog("Impfung wurde gelöscht: " & strSQL, AktionslogKat.Integration_BH_Impfungen, trans)
            Else
                strSQL = "update ghdaten..impfdoku set " &
                    "datum='" & CDate(vacc.datum).ToShortDateString & "'," &
                    "nname='" & pat.nn & "'," &
                    "vname='" & pat.vn & "'," &
                    "gebdat='" & pat.gebdat & "'," &
                    "chargenr=" & Charge & "," &
                    "arztnr=" & vacc.arztnr & "," &
                    "serum=" & vacc.serum & "," &
                    "bis6=" & IIf(bis6(pat.gebdat, CDate(vacc.datum)), -1, 0) & "," &
                    "eingang='" & Date.Today & "'," &
                    "geandert='OS_" & vacc.geandert & "'," &
                    "geandertam='" & vacc.geandertam & "'," &
                    "satz=0,mwsthon=0,aposatz=0,apomwst=0," &
                    "heftnr=" & vacc.heftnr & ", " &
                    "boncode=" & bc & "," &
                    "impfung=" & vacc.impfung & "," &
                    "NoBilling=-1,RsnNoBilling='BH-Impfung, Online'," &
                    "Prog=" & vacc.prog & "," &
                    "bhnr=" & vacc.bhnr & "," &
                    "standort=" & standort & "," &
                    "standortid=" & standortid & " " &
                    "where datum='" & s(0) & "' and boncode=" & bc

                db_con.FireSQL(strSQL, trans)


                AktionsLog("Impfung wurde verändert: " & strSQL, AktionslogKat.Integration_BH_Impfungen, trans)


            End If




        End If

        Return bc

    End Function
    Private Function GetBoncode(Heftnr As Integer, Impfstoff As Integer, Impfung As Integer) As String
        Return Heftnr & Format(CInt(Impfstoff), "00") & Impfung
    End Function

    Function bis6(Geburtsdatum As Date, Impfdatum As Date) As Boolean
        'Validate auf true, wenn gecancelt werden soll
        If Date.Compare(Geburtsdatum, DateAdd(Microsoft.VisualBasic.DateInterval.Year, -7, Impfdatum)) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Function GetCurPat(vacc As Idata, Optional trans As SqlClient.SqlTransaction = Nothing) As OSImpfling



        Dim db_con As New cls_db_con

        If vacc.heftnr < 0 Then
            vacc.heftnr = GetPersistentID(vacc.heftnr, DS_Typ.Impfung, trans)
        End If
        If vacc.heftnr <= 0 Then
            Throw New Exception("Kann persistente Heftnr für ID tmp. " & vacc.heftnr & " nicht eruieren.")
            Return Nothing
        End If


        Dim tb As DataTable = db_con.GetRecordSet("select * from frauen where heftnrf=" & vacc.heftnr, trans)
        If tb.Rows.Count = 0 Then
            tb = db_con.GetRecordSet("select anrede, nname nachname, vname vorname, gb_tm gebdat, heftnr heftnrf, strasse,plz,ort, svnrkind svnrdata, k_gbdatum gbdatum from frauen, eeinh where frauen.svnr_id=eeinh.svnr_id and heftnr=" & vacc.heftnr, trans)
        End If


        Dim r As DataRow = tb.Rows(0)

        Dim p As New OSImpfling

        p.nn = r("nachname")
        p.vn = r("vorname")
        p.adresse = r("strasse") & ", " & r("plz") & " " & r("ort")
        p.gebdat = r("gebdat")
        p.heftnr = r("heftnrf")
        p.sex = IIf(r("anrede") <> 0, "w", "m")
        If Not (IsDBNull(r("svnrdata"))) And Not (IsDBNull(r("gbdatum"))) Then

            p.svn = r("svnrdata") & r("gbdatum")
        End If


        Return p

    End Function

    Private Function SetPatData(pat As OSImpfling, Optional trans As SqlClient.SqlTransaction = Nothing) As Integer
        Dim db_con As New cls_db_con
        Dim NewHeftnr As Integer
        Dim tb As DataTable

        If KindUpdateMöglich(pat) Then
            UpdateKI(pat, trans)
            Return pat.heftnr
        End If

        ''Kind+Elternteil anlegen
        'If Not String.IsNullOrEmpty(pat.refpers_gebdat) And pat.heftnr_os < 0 Then
        '    Dim SVNRID As Integer = InsertEltern(pat, trans)
        '    Return InsertKind(pat, SVNRID, trans)
        'End If

        'Pat mit schon vorhandener Heftnr
        If pat.heftnr > 0 Then
            tb = db_con.GetRecordSet("select svnr_id  from frauen where heftnrf=" & pat.heftnr, trans)
            If tb.Rows.Count > 0 Then
                UpdateTN(pat, pat.heftnr, trans)
            Else
                InsertTN(pat, pat.heftnr, trans)
            End If
            Return pat.heftnr
        End If

        'Heftnr muss erst gefunden/kreiert werden
        If Not String.IsNullOrEmpty(pat.svn) Then
            tb = db_con.GetRecordset("select svnr_id, heftnrf from frauen where svnrdata='" & pat.svn.Substring(0, 4) & "' and gbdatum='" & pat.svn.Substring(4, 6) & "'", trans)
            If tb.Rows.Count = 0 Then

                tb = Search_SVN_Kind(pat.svn, trans)  'db_con.GetRecordset("select id, svnr_id, heftnr from eeinh where SVNRKIND='" & pat.svn.Substring(0, 4) & "' and k_gbdatum='" & pat.svn.Substring(4, 6) & "'", trans)
                If tb.Rows.Count > 0 Then
                    If tb.Rows(0)("heftnr") > 0 Then

                        NewHeftnr = tb.Rows(0)("heftnr")
                    Else
                        NewHeftnr = NO_VAL

                    End If

                    Return InsertTN(pat, NewHeftnr)


                End If


            Else
                If Not IsDBNull(tb.Rows(0)("heftnrf")) Then
                    If tb.Rows(0)("heftnrf") > 0 Then

                        NewHeftnr = tb.Rows(0)("heftnrf")

                    Else
                        NewHeftnr = HEFTBASIS + tb.Rows(0)("svnr_id")
                        db_con.FireSQL("update frauen set heftnrf=" & NewHeftnr & " where svnr_id=" & tb.Rows(0)("svnr_id"), trans)

                    End If

                    UpdateTN(pat, NewHeftnr, trans)
                    Return NewHeftnr
                Else
                    NewHeftnr = InsertHeftnr(pat, tb.Rows(0)("svnr_id"), trans)
                    'pat.heftnr = NewHeftnr
                    UpdateTN(pat, NewHeftnr, trans)
                    Return NewHeftnr
                End If
            End If


        End If


        'Jetzt über NN+VN+Gebdat - und ggf Adresse
        tb = db_con.GetRecordset("select svnr_id, heftnrf, plz, strasse, ort from frauen where " &
                                                  "nachname='" & pat.nn & "' " &
                                                  "and vorname='" & pat.nn & "' " &
                                                  "and gebdat='" & pat.gebdat & "' ", trans)


        If tb.Rows.Count > 0 Then
            Dim d As New DataView(tb)

            If d.Count > 1 Then
                d.RowFilter = "strasse='" & pat.strasse & "'"
            End If

            If d.Count = 1 Then
                If IsDBNull(d(0)("heftnrf")) Then
                    NewHeftnr = InsertHeftnr(pat, d(0)("svnr_id"), trans)
                Else
                    If d(0)("heftnrf") = 0 Then
                        NewHeftnr = InsertHeftnr(pat, d(0)("svnr_id"), trans)
                    Else
                        NewHeftnr = d(0)("heftnrf")
                    End If
                End If
                Return NewHeftnr
            End If



        End If

        tb = db_con.GetRecordSet("select frauen.svnr_id, heftnr, plz, strasse, ort,SVNRKIND, k_gbdatum from eeinh, frauen where " &
                                "frauen.svnr_id=eeinh.svnr_id " &
                                "and nname='" & pat.nn & "' " &
                                "and vname='" & pat.nn & "' " &
                                "and gb_tm='" & pat.gebdat & "' ", trans)


        If tb.Rows.Count > 0 Then
            Dim d As New DataView(tb)
            If d.Count > 1 Then
                d.RowFilter = "strasse='" & pat.strasse & "'"
            End If

            If d.Count = 1 Then
                If Not IsDBNull(d(0)("SVNRKIND") And Not IsDBNull(d(0)("k_gbdatum"))) Then
                    pat.svn = d(0)("SVNRKIND") & d(0)("k_gbdatum")
                End If

                If IsDBNull(d(0)("heftnr")) Then
                    NewHeftnr = InsertTN(pat, NO_VAL, trans)
                Else
                    If d(0)("heftnrf") = 0 Then
                        NewHeftnr = InsertTN(pat, NO_VAL, trans)
                    Else

                        NewHeftnr = d(0)("heftnrf")
                        InsertTN(pat, NewHeftnr, trans)
                    End If
                End If
                Return NewHeftnr



            End If
        End If


        Return InsertTN(pat, NO_VAL, trans)


    End Function

    Private Function Search_SVN_Kind(svn As String, Optional trans As SqlClient.SqlTransaction = Nothing) As DataTable
        Dim db_con As New cls_db_con
        Return db_con.GetRecordSet("select id, svnr_id, heftnr from eeinh where SVNRKIND='" & svn.Substring(0, 4) & "' and k_gbdatum='" & svn.Substring(4, 6) & "'", trans)
    End Function
    Private Function Search_SVN_TN(svn As String, Optional trans As SqlClient.SqlTransaction = Nothing) As DataTable
        Dim db_con As New cls_db_con
        Return db_con.GetRecordSet("select svnr_id, heftnrf from frauen where svnrdata='" & svn.Substring(0, 4) & "' and gbdatum='" & svn.Substring(4, 6) & "'", trans)

    End Function

    Private Function Search_Kombi_TN(pat As OSImpfling, Optional trans As SqlClient.SqlTransaction = Nothing) As DataTable
        Dim db_con As New cls_db_con
        Return db_con.GetRecordSet("select svnr_id from frauen where nachname='" & pat.nn & "' and vorname='" & pat.vn & "' " &
                                     "and gebdat='" & pat.gebdat & "'", trans)

    End Function

    Private Function Search_Kombi_Kind(pat As OSImpfling, Optional trans As SqlClient.SqlTransaction = Nothing) As DataTable
        Dim db_con As New cls_db_con
        Return db_con.GetRecordSet("select heftnr from eeinh where nname='" & pat.nn & "' and vname='" & pat.vn & "' " &
                                     "and gb_tm='" & pat.gebdat & "'", trans)

    End Function


    'Private Function InsertKind(pat As OSImpfling, SVNRID As Integer, trans As SqlClient.SqlTransaction) As Integer


    '    Dim newheftnr As Integer = InsertTN(pat, NO_VAL, trans)
    '    AktionsLog("Neuer Impfling - Kind muß wg HeftNr-Konflikten als Teilnehmer angelegt werden - Heftnr " & newheftnr, AktionslogKat.Integration_BH_Impfungen, trans)
    '    Return newheftnr



    '    Dim db_con As New cls_db_con
    '    Dim hasSVN As Boolean = IIf(String.IsNullOrEmpty(pat.svn), False, True)
    '    Dim svnrdata As String = "NULL"
    '    Dim gebdat As String = "NULL"

    '    If hasSVN Then
    '        svnrdata = "'" & pat.svn.Substring(0, 4) & "'"
    '        gebdat = "'" & pat.svn.Substring(4, 6) & "'"
    '    End If


    '    Dim heftnr As Integer = HEFTBASIS + SVNRID

    '    If HeftnrVergeben(heftnr) Then


    '        Dim newheftnr As Integer = InsertTN(pat, NO_VAL, trans)
    '        AktionsLog("Neuer Impfling - Kind muß wg HeftNr-Konflikten als Teilnehmer angelegt werden - Heftnr " & newheftnr, AktionslogKat.Integration_BH_Impfungen, trans)
    '        Return newheftnr
    '    Else
    '        Dim sql As String = "insert into eeinh (svnr_id,gb_tm, vname,nname,svnrkind,k_gbdatum," &
    '       "kreiert,geandert, kreiertam,geandertam,heftnr) values (" &
    '       SVNRID & "," &
    '       "'" & pat.gebdat & "'," &
    '       "'" & pat.vn & "'," &
    '       "'" & pat.nn & "'," &
    '       svnrdata & "," &
    '       gebdat & "," &
    '       "'OS_" & pat.username & "'," &
    '       "'OS_" & pat.username & "'," &
    '       "'" & Date.Now & "'," &
    '       "'" & Date.Now & "'," &
    '       "0)"



    '        db_con.FireSQL(sql, trans)

    '        AktionsLog("Neuer Impfling: " & sql, AktionslogKat.Integration_BH_Impfungen, trans)

    '        Return InsertHeftnrKind(pat, SVNRID, trans)
    '    End If













    'End Function

    Private Sub UpdateTN(pat As OSImpfling, heftnr As Integer, Optional trans As SqlClient.SqlTransaction = Nothing)

        Dim db_con As New cls_db_con
        Dim hasSVN As Boolean = IIf(String.IsNullOrEmpty(pat.svn), False, True)
        If Not hasSVN Then

            Dim tb As DataTable = db_con.GetRecordset("select svnrdata, gbdatum from frauen where heftnrf=" & heftnr, trans)

            If Not IsDBNull(tb.Rows(0)("svnrdata")) And Not IsDBNull(tb.Rows(0)("gbdatum")) Then
                pat.svn = tb.Rows(0)("svnrdata") & tb.Rows(0)("gbdatum")
                hasSVN = True
            End If
        End If

        Dim ic As New ImportGeocodierung(pat.strasse, pat.plz, pat.ort)


        Dim svnrdata As String = "NULL"
        Dim gebdat As String = "NULL"

        If hasSVN Then
            svnrdata = "'" & pat.svn.Substring(0, 4) & "'"
            gebdat = "'" & pat.svn.Substring(4, 6) & "'"
        Else
            svnrdata = "'" & GetGBShort(pat.gebdat) & "'"
        End If

        Dim SQLRefpers As String = ""
        If Not String.IsNullOrEmpty(pat.refpers_nn) Then
            SQLRefpers = "nneltern=" & GetFldVal(pat.refpers_nn) & "," &
                       "vneltern=" & GetFldVal(pat.refpers_vn) & "," &
                       "titeleltern=" & GetFldVal(pat.refpers_titel) & "," &
                       "titel_suffix_eltern=" & GetFldVal(pat.refpers_ngtitel) & ","

        End If

        Dim sql As String = "update frauen set " &
                       "nachname='" & pat.nn & "'," &
                       "vorname='" & pat.vn & "'," &
                       "anrede=" & IIf(pat.sex = "w", -1, 0) & "," &
                       "gebdat='" & pat.gebdat & "'," &
                       "gbjahr=" & CDate(pat.gebdat).Year & "," &
                       "strasse='" & pat.strasse & " " & pat.hnr & "'," &
                       "plz=" & pat.plz & "," &
                       "ort='" & pat.ort & "'," &
                       "svnrdata=" & svnrdata & "," &
                       "gbdatum=" & gebdat & "," &
                       "geandert='OS_" & pat.username & "'," &
                       "geandertam='" & Date.Now & "', " &
                       "AlterAddress='" & Date.Now & "', " &
                        SQLRefpers &
                        ic.UpdateSQLmitGKZ &
                        "where heftnrf=" & heftnr

        db_con.FireSQL(sql, trans)


        AktionsLog("Update Impfling: " & sql, AktionslogKat.Integration_BH_Impfungen, trans)



    End Sub

    Private Sub UpdateKI(pat As OSImpfling, Optional trans As SqlClient.SqlTransaction = Nothing)

        Dim db_con As New cls_db_con
        Dim hasSVN As Boolean = IIf(String.IsNullOrEmpty(pat.svn), False, True)


        UpdateEltern(pat, trans)

        Dim svnrdata As String = "NULL"
        Dim gebdat As String = "NULL"

        If hasSVN Then
            svnrdata = "'" & pat.svn.Substring(0, 4) & "'"
            gebdat = "'" & pat.svn.Substring(4, 6) & "'"
        End If







        Dim sql As String = "update eeinh set " &
                       "nname='" & pat.nn & "'," &
                       "vname='" & pat.vn & "'," &
                       "gb_tm='" & pat.gebdat & "'," &
                       "SVNRKIND=" & svnrdata & "," &
                       "k_gbdatum=" & gebdat & "," &
                       "KSVA=8," &
                       "sex=" & IIf(pat.sex = "w", 1, 0) & "," &
                       "geandert='OS_" & pat.username & "'," &
                       "geandertam='" & Date.Now & "' " &
                        "where heftnr=" & pat.heftnr

        db_con.FireSQL(sql, trans)


        AktionsLog("Update Impfling: " & sql, AktionslogKat.Integration_BH_Impfungen, trans)



    End Sub


    'Private Sub InsertKI(pat As OSImpfling, heftnr As Integer, SVNRID As Integer, Optional trans As SqlClient.SqlTransaction = Nothing)

    '    Dim db_con As New cls_db_con
    '    Dim hasSVN As Boolean = IIf(String.IsNullOrEmpty(pat.svn), False, True)





    '    UpdateMutter(pat, SVNRID, trans)

    '    Dim svnrdata As String = "NULL"
    '    Dim gebdat As String = "NULL"

    '    If hasSVN Then
    '        svnrdata = "'" & pat.svn.Substring(0, 4) & "'"
    '        gebdat = "'" & pat.svn.Substring(4, 6) & "'"
    '    End If







    '    Dim sql As String = "insert into eeinh  (svnr_id,nname,vname,gb_tm,SVNRKIND,k_gbdatum,KSVA,sex,geandert,geandertam,kreiert,kreiertam,heftnr) values(" &
    '                    SVNRID & "," &
    '                   "'" & pat.nn & "'," &
    '                   "'" & pat.vn & "," &
    '                   "'" & pat.gebdat & "'," &
    '                   svnrdata & "," &
    '                    gebdat & "," &
    '                   "8," &
    '                    IIf(pat.sex = "w", 1, 0) & "," &
    '                   "'OS_" & pat.username & "'," &
    '                   "'" & Date.Now & "', " &
    '                   "'OS_" & pat.username & "'," &
    '                   "'" & Date.Now & "') " &
    '                    "where heftnr=" & heftnr

    '    db_con.FireSQL(sql, trans)





    '    AktionsLog("Neuer Impfling: " & sql, AktionslogKat.Integration_BH_Impfungen, trans)
    '    If heftnr = NO_VAL Then





    '    End If


    'End Sub



    Private Sub UpdateEltern(pat As OSImpfling, Optional trans As SqlClient.SqlTransaction = Nothing)



        Dim db_con As New cls_db_con
        Dim tb As DataTable = db_con.GetRecordSet("select svnr_id from eeinh where heftnr=" & pat.heftnr)


        Dim ic As New ImportGeocodierung(pat.strasse, pat.plz, pat.ort)
        Dim sql As String = "update frauen set " &
                       "strasse='" & pat.strasse & " " & pat.hnr & "'," &
                       "plz=" & pat.plz & "," &
                       "ort='" & pat.ort & "'," &
                       "geandert='OS_" & pat.username & "'," &
                       "geandertam='" & Date.Now & "', " &
                       "AlterAddress='" & Date.Now & "', " &
                       ic.UpdateSQLmitGKZ &
                        "where svnr_id=" & tb.Rows(0)("svnr_id")

        db_con.FireSQL(sql, trans)
        AktionsLog("Update Eltern Impfling: " & sql, AktionslogKat.Integration_BH_Impfungen, trans)



    End Sub

    Private Function HeftnrVergeben(heftnr As Integer, Optional trans As SqlClient.SqlTransaction = Nothing) As Boolean
        Dim db_con As New cls_db_con
        Dim tb As DataTable = db_con.GetRecordSet("select svnr_id from frauen where heftnrf=" & heftnr, trans)

        If tb.Rows.Count > 0 Then Return True

        tb = db_con.GetRecordSet("select id from eeinh where heftnr=" & heftnr)


        If tb.Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If


    End Function


    Private Function KindUpdateMöglich(pat As OSImpfling, Optional trans As SqlClient.SqlTransaction = Nothing) As Boolean

        Dim tb As DataTable
        Dim db_con As New cls_db_con
        If Date.Compare(pat.gebdat, DateAdd(Microsoft.VisualBasic.DateInterval.Year, -7, Date.Today)) > 0 Then


            If pat.heftnr_os < 0 Then
                If Not String.IsNullOrEmpty(pat.svn) Then
                    tb = Search_SVN_Kind(pat.svn)
                    If tb.Rows.Count = 1 Then

                        pat.heftnr = tb.Rows(0)(0)
                        Return True

                    Else
                        Return False

                    End If

                Else
                    Return False

                End If



            End If






            'Existiert heftNr schon als TN?

            tb = db_con.GetRecordSet("Select heftnrf from frauen where heftnrf=" & pat.heftnr_os, trans)
            If tb.Rows.Count > 0 Then Return False






            'Exisitiert Kind-DS?
            'Wir können aus einem Kind-DS nicht den Erwachsenen ableiten!
            tb = db_con.GetRecordSet("Select heftnr from eeinh where heftnr=" & pat.heftnr_os, trans)
            If tb.Rows.Count = 0 Then
                Return False

            Else
                Return True
            End If


        Else
            'Ältere Personen
            Return False
        End If
    End Function

    Private Function InsertTN(pat As OSImpfling, heftnr As Integer, Optional trans As SqlClient.SqlTransaction = Nothing) As Integer

        Dim db_con As New cls_db_con
        Dim tb As DataTable
        Dim hasSVN As Boolean = IIf(String.IsNullOrEmpty(pat.svn), False, True)

        If Not hasSVN And heftnr > 0 Then
            'Kind mit SVN?

            tb = db_con.GetRecordSet("Select SVNRKIND, k_gbdatum from eeinh where heftnr=" & heftnr, trans)
            If tb.Rows.Count > 0 Then
                If Not IsDBNull(tb.Rows(0)("SVNRKIND") And Not IsDBNull(tb.Rows(0)("k_gbdatum"))) Then
                    pat.svn = tb.Rows(0)("SVNRKIND") & tb.Rows(0)("k_gbdatum")
                    hasSVN = True
                End If
            End If
        End If

        Dim ic As New ImportGeocodierung(pat.strasse & " " & pat.hnr, pat.plz, pat.ort)

        Dim svnrdata As String = "NULL"
        Dim gebdat As String = "NULL"

        If hasSVN Then
            svnrdata = "'" & pat.svn.Substring(0, 4) & "'"
            gebdat = "'" & pat.svn.Substring(4, 6) & "'"
        Else
            gebdat = "'" & GetGBShort(pat.gebdat) & "'"
        End If



        Dim sql As String = "insert into frauen (heftnrf,anrede,nachname,vorname,gebdat,gbjahr,strasse,plz,ort,svnrdata," &
                        "gbdatum,sva,kreiertam,geandertam,AlterAddress,kreiert,geandert, " &
                        "nneltern,vneltern,titeleltern,titel_suffix_eltern," &
                        "breitengrad,laengengrad," &
                        "GeocodeLevel,gemeindeid) values ( " &
                       IIf(heftnr = NO_VAL, 0, heftnr) & "," &
                       IIf(pat.sex = "w", -1, 0) & "," &
                       " '" & pat.nn & "'," &
                       "'" & pat.vn & "'," &
                       "'" & pat.gebdat & "'," &
                       CDate(pat.gebdat).Year & "," &
                       "'" & pat.strasse & " " & pat.hnr & "'," &
                       pat.plz & "," &
                       "'" & pat.ort & "'," &
                       svnrdata & "," &
                       gebdat & ", " &
                       "8," &
                       "'" & Date.Now & "'," &
                       "'" & Date.Now & "'," &
                       "'" & Date.Now & "'," &
                       "'OS_" & pat.username & "'," &
                       "'OS_" & pat.username & "'," &
                       GetFldVal(pat.refpers_nn) & "," &
                       GetFldVal(pat.refpers_vn) & "," &
                       GetFldVal(pat.refpers_titel) & "," &
                       GetFldVal(pat.refpers_ngtitel) & "," &
                       ic.InsertSQLMitGKZ & ")"




        db_con.FireSQL(sql, trans)

        AktionsLog("Neuer Impfling: " & sql, AktionslogKat.Integration_BH_Impfungen, trans)


        If heftnr = NO_VAL Then


            heftnr = InsertHeftnr(pat,, trans)


        End If


        Return heftnr

    End Function

    Public Function GetGBShort(Gebdat As DateTime) As String
        Return Gebdat.Day.ToString.PadLeft(2, "0") &
            Gebdat.Month.ToString.PadLeft(2, "0") &
            Gebdat.Year.ToString.Substring(2, 2)

    End Function
    'Private Function InsertEltern(pat As OSImpfling, Optional trans As SqlClient.SqlTransaction = Nothing) As Integer

    '    Dim db_con As New cls_db_con
    '    Dim tb As DataTable
    '    Dim hasSVN As Boolean = IIf(String.IsNullOrEmpty(pat.refpers_vnr1) Or String.IsNullOrEmpty(pat.refpers_vnr2), False, True)

    '    Dim svnrdata As String = "NULL"
    '    Dim gebdat As String = "NULL"

    '    Dim SVNRID As Integer = 0


    '    If hasSVN Then
    '        svnrdata = "'" & pat.refpers_vnr1 & "'"
    '        gebdat = "'" & pat.refpers_vnr2 & "'"


    '        tb = db_con.GetRecordSet("select svnrid from frauen where svnrdata='" & pat.refpers_vnr1 & "' and gbdatum='" & pat.refpers_vnr2 & "'", trans)
    '        If tb.Rows.Count > 0 Then SVNRID = tb.Rows(0)("svnr_id")



    '    End If

    '    Dim Titel As String = "NULL"
    '    Dim NGTitel As String = "NULL"
    '    If Not String.IsNullOrEmpty(pat.refpers_titel) Then Titel = "'" & pat.refpers_titel & "'"
    '    If Not String.IsNullOrEmpty(pat.refpers_ngtitel) Then NGTitel = "'" & pat.refpers_ngtitel & "'"


    '    If SVNRID = 0 Then

    '        tb = db_con.GetRecordSet("select svnr_id from frauen where nachname='" & pat.refpers_nn & "' and vorname='" & pat.refpers_vn & "' " &
    '                                 "and gebdat='" & pat.refpers_gebdat & "'", trans)


    '        If tb.Rows.Count = 1 Then
    '            SVNRID = tb.Rows(0)(0)
    '        End If

    '    End If

    '    Dim ic As New ImportGeocodierung(pat.strasse & " " & pat.hnr, pat.plz, pat.ort)

    '    Dim Jetzt As Date = Date.Now
    '    If SVNRID > 0 Then



    '        Dim sql As String = "update frauen set " &
    '                   "anrede=" & IIf(pat.refpers_sex = "w", -1, 0) & "," &
    '                   "nachname='" & pat.refpers_nn & "'," &
    '                   "vorname='" & pat.refpers_vn & "'," &
    '                   "gebdat='" & pat.refpers_gebdat & "'," &
    '                   "strasse='" & pat.strasse & " " & pat.hnr & "'," &
    '                   "plz=" & pat.plz & "," &
    '                   "ort='" & pat.ort & "'," &
    '                   "svnrdata=" & svnrdata & "," &
    '                   "gbdatum=" & gebdat & ", " &
    '                   "sva=8," &
    '                   "kreiertam='" & Jetzt & "'," &
    '                   "geandertam='" & Jetzt & "'," &
    '                   "AlterAddress='" & Jetzt & "'," &
    '                   "kreiert='OS_" & pat.username & "'," &
    '                   "geandert='OS_" & pat.username & "'," &
    '                   "titel=" & Titel & "," &
    '                   "titel_suffix=" & NGTitel & "," &
    '                   ic.UpdateSQLmitGKZ & " " &
    '                   "where svnr_id=" & SVNRID

    '        db_con.FireSQL(sql, trans)

    '        AktionsLog("Update Eltern-DS: " & sql, AktionslogKat.Integration_BH_Impfungen, trans)

    '    Else
    '        Dim sql As String = "insert into frauen (heftnrf,anrede,nachname,vorname,gebdat,strasse,plz,ort,svnrdata,gbdatum,sva,kreiertam,geandertam,AlterAddress,kreiert,geandert,titel, titel_suffix, breitengrad,laengengrad,GeocodeLevel,gemeindeid) values ( " &
    '                   "0," &
    '                   IIf(pat.refpers_sex = "w", -1, 0) & "," &
    '                   " '" & pat.refpers_nn & "'," &
    '                   "'" & pat.refpers_vn & "'," &
    '                   "'" & pat.refpers_gebdat & "'," &
    '                   "'" & pat.strasse & " " & pat.hnr & "'," &
    '                   pat.plz & "," &
    '                   "'" & pat.ort & "'," &
    '                   svnrdata & "," &
    '                   gebdat & ", " &
    '                   "8," &
    '                   "'" & Jetzt & "'," &
    '                   "'" & Jetzt & "'," &
    '                   "'" & Jetzt & "'," &
    '                   "'OS_" & pat.username & "'," &
    '                   "'OS_" & pat.username & "'," &
    '                   Titel & "," &
    '                   NGTitel & "," &
    '                   ic.InsertSQLMitGKZ & ")"

    '        db_con.FireSQL(sql, trans)

    '        AktionsLog("Neuer Eltern-DS: " & sql, AktionslogKat.Integration_BH_Impfungen, trans)

    '        tb = db_con.GetRecordSet("select svnr_id from frauen where " &
    '                    "kreiertam='" & Jetzt & "' " &
    '                    "and nachname='" & pat.refpers_nn & "' " &
    '                    "and kreiert='OS_" & pat.username & "'", trans)


    '        If tb.Rows.Count = 0 Then
    '            Throw New Exception("Eingefügten Elternteil nicht gefunden.")
    '        Else
    '            SVNRID = tb.Rows(0)("svnr_id")
    '        End If

    '    End If

    '    Return SVNRID



    'End Function

    Private Function InsertHeftnr(pat As OSImpfling, Optional SVNRID As Integer = NO_VAL, Optional trans As SqlClient.SqlTransaction = Nothing) As Integer
        Dim db_con As New cls_db_con
        If SVNRID = NO_VAL Then SVNRID = getSVNRID(pat, trans)
        Dim heftnr As Integer = HEFTBASIS + SVNRID

        Dim sql As String = "update frauen set heftnrf=" & heftnr & " where svnr_id=" & SVNRID

        db_con.FireSQL(sql, trans)

        AktionsLog("Neue Impf-ID: " & sql, AktionslogKat.Integration_BH_Impfungen, trans)

        Return heftnr
    End Function

    'Private Function InsertHeftnrKind(pat As OSImpfling, SVNRID As Integer, Optional trans As SqlClient.SqlTransaction = Nothing) As Integer
    '    Dim db_con As New cls_db_con
    '    Dim heftnr As Integer = HEFTBASIS + SVNRID





    '    Dim tb As DataTable = db_con.GetRecordSet("select id from eeinh where " &
    '                                             "nname='" & pat.nn & "' " &
    '                                             "and vname='" & pat.vn & "' " &
    '                                             "and gb_tm='" & pat.gebdat & "' " &
    '                                             "and svnr_id=" & SVNRID, trans)


    '    If tb.Rows.Count = 0 Then
    '        Throw New Exception("Kinderdatensatz wurde nicht gefunden.")
    '    End If

    '    Dim KindID As Integer = tb.Rows(0)(0)


    '    Dim sql As String = "update eeinh set heftnr=" & heftnr & " where id=" & KindID

    '    db_con.FireSQL(sql, trans)

    '    AktionsLog("Neue Impf-ID: " & sql, AktionslogKat.Integration_BH_Impfungen, trans)

    '    Return heftnr
    'End Function




    Private Function getSVNRID(pat As OSImpfling, trans As SqlClient.SqlTransaction) As Integer
        Dim db_con As New cls_db_con
        Dim tb As DataTable = db_con.GetRecordSet("select svnr_id from frauen where nachname='" & pat.nn & "' and vorname='" & pat.vn & "' and gebdat='" & pat.gebdat & "' and strasse='" & pat.strasse & " " & pat.hnr & "' and plz=" & pat.plz, trans)
        Return tb.Rows(0)("svnr_id")

    End Function


    Private Sub UpdateKontodaten(iban As String, bic As Integer, arztnr As Integer)

        Dim db_con As New cls_db_con

        Dim blz As Integer = BLZfromIBAN(iban)
        Dim Konto As String = KontoNrfromIBAN(iban)
        Dim institut As String = GetInstitut(blz)

        If String.IsNullOrEmpty(institut) Then
            institut = "NULL"
        Else
            institut = "'" & institut & "'"
        End If


        Dim ret As Integer = db_con.FireSQL("update aerzteliste set iban='" & iban & "', bic='" & bic & "', instutitut=" & institut & ", blz=" & blz & ", konto=" & Konto & ", billingdata=-1 where arztnr=" & arztnr)




        Dim TargetAddr As String = URL_Online_Service & "/ConfigEx/SetParam?Param="
        SetRemoteParam(TargetAddr, ConfigParam.ValidateDocSync, arztnr)










    End Sub


    Public ToDoList As New System.Text.StringBuilder
    Private Function GetInstitut(blz As Integer) As String
        Try
            Dim db_con As New cls_db_con
            Dim rstmp As DataTable
            rstmp = db_con.GetRecordset("select institut from blz with (nolock) where blz=" & blz)
            If rstmp.Rows.Count > 0 Then

                Return rstmp.Rows(0)("Institut")
            End If
        Catch ex As Exception

        End Try


        Return ""

    End Function

    Private Function BLZfromIBAN(strIBAN As String) As String
        Return strIBAN.Substring(4, 5)
    End Function

    Private Function CleanText(txt As String) As String
        Return txt.Replace("'", "''")
    End Function


    Private Function KontoNrfromIBAN(strIBAN As String) As String
        Return strIBAN.Substring(9)

    End Function
    Private Function GetRemoteParam(TargetAddr As String) As String
        Dim WebRequest As HttpWebRequest
        Dim WebResponse As HttpWebResponse = Nothing
        Dim responseString As String

        Try


            WebRequest = CType(Net.WebRequest.Create(TargetAddr), HttpWebRequest) 'CType(WebRequest.Create(TargetAddr), HttpWebRequest)
            WebRequest.Timeout = 10000
            WebRequest.ReadWriteTimeout = 10000

            WebResponse = CType(WebRequest.GetResponse(), HttpWebResponse)
            'receiveStream = WebResponse.GetResponseStream()
            'totlen = WebResponse.ContentLength


            Using stream As Stream = WebResponse.GetResponseStream()
                Dim reader As New StreamReader(stream, Encoding.UTF8)
                responseString = reader.ReadToEnd()
            End Using

            Return responseString

        Catch ex As Exception
            'AktionsLog("Fehler: " & ex.Message & vbNewLine & ex.StackTrace, AktionslogKat.Integration_BH_Impfungen)
            LogErr(ex, "")
            If Not WebResponse Is Nothing Then WebResponse.Close()

        End Try

        Return ""

    End Function


    Private Sub LogErr(ex As Exception, LastErrSQL As String)
        AktionsLog("Fehler: " & ex.Message & vbNewLine & ex.StackTrace, AktionslogKat.Integration_BH_Impfungen)
        SendAdminmail("Fehler: " & ex.Message & vbNewLine & ex.StackTrace & vbNewLine & vbNewLine & LastErrSQL, "Fehler beim Synchronisieren")
    End Sub


    Public Sub AktionsLog(ByVal aktion As String, ByVal VProgramm As AktionslogKat, Optional trans As SqlClient.SqlTransaction = Nothing)
        Try
            Dim db_con As New cls_db_con
            db_con.FireSQL("aktionslogreg @aktion='" & GetSQLText(aktion) & "', @VProgramm=" & GetSQLText(VProgramm) & ", @CurUsr='System'")
            Console.WriteLine(aktion)
        Catch ex As Exception
        Finally
        End Try


    End Sub

    Private Function GetSQLText(txt As String) As String
        Return txt.Replace("'", "''")
    End Function


    Private Function GetFldVal(txt As String) As String
        If String.IsNullOrEmpty(txt) Then
            Return "NULL"
        Else
            Return "'" & txt & "'"
        End If

    End Function

    Private Function SetRemoteParam(TargetAddr As String, Param As ConfigParam, Val As String, Optional verbose As Boolean = False) As Boolean
        Dim uri As New Uri(TargetAddr)
        Dim servicePoint As Net.ServicePoint = ServicePointManager.FindServicePoint(uri)
        servicePoint.Expect100Continue = False

        'System.Net.ServicePointManager.Expect100Continue = False
        Dim response As WebResponse = Nothing

        Dim outgoingQueryString As NameValueCollection = Web.HttpUtility.ParseQueryString([String].Empty)
        outgoingQueryString.Add("Param", CInt(Param))
        outgoingQueryString.Add("val", Val)
        outgoingQueryString.Add("U", "toni")
        outgoingQueryString.Add("P", "KeinHund!")

        Dim postdata As String = outgoingQueryString.ToString()


        Try
            ' Create a request using a URL that can receive a post. 
            Dim request As WebRequest = WebRequest.Create(TargetAddr)
            request.Method = WebRequestMethods.Http.Post

            request.Timeout = 30 * 60 * 1000
            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postdata)
            request.ContentType = "application/x-www-form-urlencoded"
            request.ContentLength = byteArray.Length
            Dim dataStream As Stream = request.GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close()

            response = request.GetResponse()
            'Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)

            dataStream = response.GetResponseStream()
            Dim reader As New StreamReader(dataStream)
            Dim responseFromServer As String = reader.ReadToEnd()
            reader.Close()
            dataStream.Close()
            response.Close()

            If verbose Then
                'MsgBox(responseFromServer, MsgBoxStyle.Information, APP_NAME)
            Else
                If responseFromServer.ToLower <> "false" Then
                    If responseFromServer.ToLower = "true" Then
                        'log("Parameter " & CType(Param, ConfigParam).ToString & " konnte nicht gesetzt werden.")

                        AktionsLog("Fehler: Parameter " & CType(Param, ConfigParam).ToString & " konnte nicht gesetzt werden.", AktionslogKat.Integration_BH_Impfungen)

                    Else
                        'log(responseFromServer)
                        AktionsLog(responseFromServer, AktionslogKat.Integration_BH_Impfungen)

                    End If

                    Return True

                End If


            End If

            Return False


        Catch ex As Exception

            If Not response Is Nothing Then response.Close()
            'AktionsLog("Fehler: " & ex.Message & vbNewLine & ex.StackTrace, AktionslogKat.Integration_BH_Impfungen)
            LogErr(ex, "")
            Return True
        End Try




    End Function

    'Private Sub log(txt)
    '    Dim db_con As New cls_db_con
    '    If String.IsNullOrEmpty(txt) Then Return

    '    If txt.Length > 3000 Then txt = txt.Substring(0, 3000)


    '    db_con.FireSQL("insert into wf_log (log_ma,log_datum,log_txt) values ('Syncdienst OS-MR','" & Date.Now & "','" & txt.Replace("'", "''") & "')")
    'End Sub
    Private ReadOnly Property URL_Online_Service() As String
        Get

            If IsTestEnvironment() Then
                Return "http://localhost:17221"
            Else
                Return "https://www.ghdbservices.at/wavm"

            End If





        End Get

    End Property

    Private Function IsTestEnvironment() As Boolean

        If Environment.MachineName.ToLower = TEST_SERVER.ToLower Or Environment.MachineName.ToLower = TEST_SERVER2.ToLower Then
            Return True
        Else
            Return False
        End If

    End Function


    Private Function ToBase64(ByVal sText As String) As String
        Return System.Convert.ToBase64String(System.Text.Encoding.Default.GetBytes(sText))
    End Function
    Private Function FromBase64(ByVal sText As String) As String
        Return System.Text.Encoding.Default.GetString(System.Convert.FromBase64String(sText))
    End Function

End Class


Public Class OSVaccData
    Public Vacc As List(Of Idata)
    Public pers As List(Of OSImpfling)
End Class
Public Class OSImpfling
    Private m_hnr As String
    Private m_str As String

    Public Property tmp As Boolean
    Public Property username As String


    Public Property heftnr As Integer
    Public Property heftnr_os As Integer
    Public Property nn As String
    Public Property vn As String
    Public Property sex As String

    Public Property adresse As String

    Public ReadOnly Property strasse As String
        Get
            If String.IsNullOrEmpty(m_str) Then Setstr()
            Return m_str
        End Get
    End Property


    Public ReadOnly Property hnr As String
        Get
            If String.IsNullOrEmpty(m_str) Then Setstr()
            Return m_hnr
        End Get
    End Property


    Public ReadOnly Property plz As Integer
        Get
            If String.IsNullOrEmpty(adresse) Then Return 0
            Return adresse.Split(",")(1).Trim.Split(" ")(0).Trim
        End Get
    End Property

    Public ReadOnly Property ort As String
        Get
            If String.IsNullOrEmpty(adresse) Then Return ""
            Return adresse.Split(",")(1).Trim.Split(" ")(1).Trim
        End Get
    End Property


    Public Property gebdat As String

    Public Property svn As String

    Private Sub Setstr()
        Dim i As Integer
        Dim Pos As Integer = 1


        If String.IsNullOrEmpty(adresse) Then
            m_str = ""
            m_hnr = ""
            Return

        End If

        Dim str = adresse.Split(",")(0).Trim
        If String.IsNullOrEmpty(str) Then Return

        If str.IndexOf(" ") < 0 Then
            m_str = str
            m_hnr = ""
            Return
        End If




        'Falls str mit Ziffer anfängt
        If IsNumeric(Mid$(str, 1, 1)) Then
            For i = 1 To Len(str)
                If Not IsNumeric(Mid$(str, i, 1)) Then Exit For
            Next
            Pos = i
            If Pos >= str.Length Then Return

        End If


        For i = Pos To Len(str)
            If IsNumeric(Mid$(str, i, 1)) Then
                m_str = Trim(Left(str, i - 1))
                m_hnr = Trim(Right(str, Len(str) - i + 1))
                Return

            End If
        Next

    End Sub
    Public Property refpers_nn As String
    Public Property refpers_vn As String
    Public Property refpers_sex As String
    Public Property refpers_titel As String
    Public Property refpers_ngtitel As String
    Public Property refpers_vnr1 As String
    Public Property refpers_vnr2 As String
    Public Property refpers_gebdat As String

End Class

Public Class Idata
    Public Property heftnr As Integer
    Public Property serum As Integer
    Public Property impfung As Integer
    Public Property arztnr As Integer
    Public Property datum As String
    Public Property charge As String = ""
    Public Property standortid As String = ""
    Public Property standort As String = ""
    Public Property dsid As Integer

    Public Property bhnr As Integer

    Public Property geandert As String

    Public Property geandertam As Date

    Public Property prog As Integer
    Public Property plid As String = ""

    Public Property del As Integer = 0


End Class



Public Class ImportGeocodierung
    Public Const NO_STADTBEZIRK As String = "99"

    Public InsertSQL As String
    Public InsertSQLMitGKZ As String
    Public UpdateSQL As String
    Public UpdateSQLmitGKZ As String
    Public GKZ As Integer



    Public Sub New(Strasse As String, PLZ As String, Ort As String)
        Dim ao As New AdresseOffiziell
        Dim adr As Adresse = ao.GeocodeImport(Strasse, PLZ, Ort)

        Dim Lat As String = "NULL"
        Dim Lon As String = "NULL"

        'If adr.ResultLevel <> geocoder.ResultLevel.Address Then
        '    oCorrParent.AddToGeoCodeLog(NN & " " & VN & " - " & Strasse & ", " & PLZ & " " & Ort)
        'End If
        If adr.ResultLevel = geocoder.ResultLevel.Address Then
            GKZ = adr.GKZ
            If adr.Lat <> 0 Then Lat = adr.LatKommaAsPoint
            If adr.Lon <> 0 Then Lon = adr.LonKommaAsPoint

        Else
            GKZ = GetGKZ(Strasse, Ort, PLZ)

        End If

        InsertSQL = Lat & "," & Lon & "," & CInt(adr.ResultLevel) & " "
        InsertSQLMitGKZ = InsertSQL & "," & GKZ & " "
        UpdateSQL = "breitengrad=" & Lat & ", laengengrad=" & Lon & ", geocodelevel=" & CInt(adr.ResultLevel) & " "
        UpdateSQLmitGKZ = UpdateSQL & ", gemeindeid=" & GKZ & " "
    End Sub


    Public Function GetGKZ(ByVal strasse As String, ByVal ort As String, ByVal plz As Integer, Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Integer
        'On Error GoTo gg_err
        Dim rs As DataTable
        Dim db_con As New cls_db_con
        rs = db_con.GetRecordset("select gemeinden.dbo.getgkz_neu('" & strasse & "','" & ort & "', " & plz & ")", trans)
        If rs.Rows.Count > 0 Then
            GetGKZ = CInt(CStr(rs.Rows(0)(0)) & CStr(IIf(Me.IsXTCity(CLng(rs.Rows(0)(0))), NO_STADTBEZIRK, "")))
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
        Dim db_con As New cls_db_con

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

        Dim db_con As New cls_db_con
        Dim l As New List(Of Integer)

        rs = db_con.GetRecordset("select gemeindeid from xt_cities")
        For Each r As DataRow In rs.Rows

            l.Add(r("gemeindeid"))


        Next

        Return l

    End Function
End Class