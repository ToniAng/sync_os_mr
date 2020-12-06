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



    Private Const TAB_PERS As String = "pers"
    Private Const TAB_VACC As String = "vacc"
    Private Const TAB_AA As String = "aa"
    Private Const TAB_AIS As String = "ais"

    Private Const KEINE_CHARGE As String = "XXX-XXX"


    Private PARAM_LAST_VACC_SYNC As String = "LastVaccSyncOS"


    Private Const MD_TRÄGER_ID_MIN As Integer = 1000


    Private Enum DS_Typ
        Impfung = 0
        Impfling = 1
        Amtsarzt = 2
        AIS_Data_log = 3
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

        GrippeImpfung65Plus = 14
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

        Console.WriteLine("Log mit DSID " & dsid & " und dsid_tmp " & dsid_tmp & " hinzugefügt.")


    End Sub

    Private Function GetPersistentID(tmp_dsid As Integer, dstyp As DS_Typ, Optional trans As SqlClient.SqlTransaction = Nothing) As Integer

        Dim db_con As New cls_db_con
        Dim tb As DataTable = db_con.GetRecordSet("Select os_dsid from os_synclog where os_dsid_tmp=" & tmp_dsid & " And os_dstyp=" & CInt(DS_Typ.Impfling), trans)
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

            If s = "False" Then
                AktionsLog("Daten für den Abgleich der Impfungen konnten nicht abgerufen werden.", AktionslogKat.Integration_BH_Impfungen)
                Return
            ElseIf s = "" Then
                AktionsLog("Keine Daten vom Online-Service (Down?)- bitte prüfen.", AktionslogKat.Integration_BH_Impfungen)
                Return
            End If
            Dim js As New JavaScriptSerializer()
            Dim osv As OSVaccData = js.Deserialize(Of OSVaccData)(s)





            trans = con.BeginTransaction()

            If osv.aa IsNot Nothing Then
                Dim curaa As amtsarzt_voll
                For Each curaa In osv.aa

                    Dim aa As New ghdb_amtsarzt
                    aa.Add(curaa, False, trans)

                    AddSynclog(DS_Typ.Amtsarzt, curaa.Arztnr,, trans)
                Next
            End If



            Dim curPat As OSImpfling
            Dim remote_heftnr As Integer = 0

            For Each curPat In osv.pers
                remote_heftnr = curPat.heftnr
                curPat.heftnr_os = curPat.heftnr
                If curPat.heftnr < 0 Then
                    curPat.heftnr = GetPersistentID(curPat.heftnr, DS_Typ.Impfling, trans)
                End If
                If curPat.heftnr < 0 Then

                Else


                    curPat.KindUpdateMöglich = KindUpdateMöglich(curPat, trans)

                End If
                curPat.heftnr = remote_heftnr
            Next

            For Each curPat In osv.pers
                remote_heftnr = 0
                curPat.heftnr_os = curPat.heftnr

                If curPat.heftnr < 0 Then
                    remote_heftnr = curPat.heftnr
                    curPat.heftnr = GetPersistentID(curPat.heftnr, DS_Typ.Impfling, trans)
                End If
                Console.WriteLine("curPat.heftnr=" & curPat.heftnr & ", remote_heftnr=" & remote_heftnr)
                If curPat.heftnr < 0 Then

                    AddSynclog(DS_Typ.Impfling, SetPatData(curPat, trans), curPat.heftnr_os, trans)
                Else
                    SetPatData(curPat, trans)
                    AddSynclog(DS_Typ.Impfling, curPat.heftnr, remote_heftnr, trans)
                End If
            Next



            For Each vacc As Idata In osv.Vacc

                curPat = GetCurPat(vacc, trans)
                AddSynclog(DS_Typ.Impfung, vacc.dsid, , trans, savevacc(vacc, curPat, trans))

            Next
            If osv.Vacc.Count > 0 Then
                AktionsLog(osv.Vacc.Count & " Impfeinträge wurden synchronisiert.", AktionslogKat.Integration_BH_Impfungen, trans)
            End If



            If osv.ais IsNot Nothing Then
                For Each ais In osv.ais

                    sync_ais(ais, osv.mobile_vorwahlen, trans)

                    AddSynclog(DS_Typ.AIS_Data_log, ais.id,, trans)
                Next


            End If


            trans.Commit()

            Console.WriteLine("Synchronisierung beendet.")

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


    Private Function sync_ais(dl As ais_data_log, mobile_vorwahlen As String, trans As SqlClient.SqlTransaction) As Boolean

        Dim js As New JavaScriptSerializer()
        Dim db_con As New cls_db_con



        Dim au As AIS_Update = js.Deserialize(Of AIS_Update)(dl.json)
        Dim sql As String
        Dim from_auseeinh As String = "from eeinh, frauen where frauen.svnr_id=eeinh.svnr_id and eeinh.heftnr="
        Dim from_ausfrauen As String = "where heftnrf="

        'Dim sql_from As String

        'Select Case au.Typ
        '    Case AIS_Update.UpdateTyp.EmailEltern, AIS_Update.UpdateTyp.TelEltern, AIS_Update.UpdateTyp.SVNImpfling
        '        sql_from = from_auseeinh & au.heftnr
        '    Case AIS_Update.UpdateTyp.EmailImpfling, AIS_Update.UpdateTyp.TelImpfling
        '        sql_from = from_ausfrauen & au.heftnr


        '    Case Else
        '        Throw New Exception("ais_data_log mit nicht behandeltem Datenfeld")
        'End Select


        Dim strlog As String
        Dim arzt As String = GetArztNameFromArztNr(dl.arztnr, trans) & " (" & dl.arztnr & ")"

        Select Case au.Typ
            Case AIS_Update.UpdateTyp.EmailEltern, AIS_Update.UpdateTyp.EmailImpfling




                strlog = "'" & Date.Today & " Email durch Arzt " & arzt & "/Ärzteinfoservice auf " & au.text & " aktualisiert' "



                sql = "update  frauen set email='" & au.text & "' " & GetAktenvermerk(strlog)

                If au.Typ = AIS_Update.UpdateTyp.EmailEltern Then
                    sql = sql & from_auseeinh & au.heftnr
                Else
                    sql = sql & from_ausfrauen & au.heftnr
                End If



            Case AIS_Update.UpdateTyp.TelEltern, AIS_Update.UpdateTyp.TelImpfling

                Dim fld As String = GetTelField(au.text, mobile_vorwahlen)

                strlog = "'" & Date.Today & " " & fld & " Telefon durch Arzt " & arzt & "/Ärzteinfoservice auf " & au.text & " aktualisiert' "


                sql = "update frauen Set " & fld & "='" & au.text & "' " & GetAktenvermerk(strlog)

                If au.Typ = AIS_Update.UpdateTyp.TelEltern Then
                    sql = sql & from_auseeinh & au.heftnr
                Else
                    sql = sql & from_ausfrauen & au.heftnr

                End If






            Case AIS_Update.UpdateTyp.SVNImpfling
                strlog = "'" & Date.Today & " SVN Impfling (Heft " & au.heftnr & ") durch Arzt " & arzt & "/Ärzteinfoservice auf " & au.text & " aktualisiert' "


                Dim strupdate_frauen As String = "update frauen Set SVNRDATA='" & au.text.Substring(0, 4) & "', GBDATUM='" & au.text.Substring(4) & "' " & GetAktenvermerk(strlog)
                Dim strupdate_eeinh As String = "update eeinh Set SVNRKIND='" & au.text.Substring(0, 4) & "', k_gbdatum='" & au.text.Substring(4) & "' where heftnr=" & au.heftnr ' & GetAktenvermerk(strlog)
                Dim strupdate_frauen2 As String = "update frauen set " & GetAktenvermerk(strlog, True)

                sql = strupdate_frauen & from_ausfrauen & au.heftnr
                db_con.FireSQL(sql, trans)
                db_con.FireSQL(strupdate_eeinh, trans)


                sql = strupdate_frauen2 & from_auseeinh & au.heftnr


            Case AIS_Update.UpdateTyp.Adresse

                strlog = "'" & Date.Today & " Adresse durch Arzt " & arzt & "/Ärzteinfoservice auf " & au.Strasse & " " & au.PLZ & " " & au.Ort & " aktualisiert' "
                Dim sql_addr = "update frauen set STRASSE='" & au.Strasse & "', plz=" & au.PLZ & ", ort='" & au.Ort & "' "

                sql = sql_addr & GetAktenvermerk(strlog) & from_ausfrauen & au.heftnr
                db_con.FireSQL(sql, trans)


                sql = sql_addr & GetAktenvermerk(strlog) & from_auseeinh & au.heftnr


            Case AIS_Update.UpdateTyp.NameImpfling
                Dim strTit_Prä As String = If(au.Titel_Pra = "", "NULL", "'" & au.Titel_Pra & "' ")
                Dim strTit_Suf As String = If(au.Titel_Suf = "", "NULL", "'" & au.Titel_Suf & "' ")
                strlog = "'" & Date.Today & " Name Impfling (Heft " & au.heftnr & ") durch Arzt " & arzt & "/Ärzteinfoservice auf " & au.Titel_Pra & " " & au.Nachname & " " & au.Vorname & ", " & au.Titel_Suf & " aktualisiert' "
                Dim strupdate_frauen As String = "update frauen set Nachname='" & au.Nachname & "', Vorname='" & au.Vorname & "', TITEL=" & strTit_Prä & ", titel_suffix=" & strTit_Suf & " " & GetAktenvermerk(strlog)
                Dim strupdate_eeinh As String = "update eeinh set nname='" & au.Nachname & "', Vname='" & au.Vorname & "' where heftnr=" & au.heftnr
                Dim strupdate_frauen2 As String = "update frauen set " & GetAktenvermerk(strlog, True)


                sql = strupdate_frauen & from_ausfrauen & au.heftnr
                db_con.FireSQL(sql, trans)

                db_con.FireSQL(strupdate_eeinh)
                sql = strupdate_frauen2 & from_auseeinh & au.heftnr




            Case Else
                Throw New Exception("ais_data_log mit nicht behandeltem Datenfeld (2)")

        End Select


        db_con.FireSQL(sql, trans)

        Return True

    End Function

    'Private Sub AISUpdateEeinh(sql As String, sql_from As String, strlog As String, db_con As cls_db_con, trans As SqlClient.SqlTransaction)
    '    Dim ret As Integer = db_con.FireSQL(sql, trans)
    '    If ret > 0 Then

    '        sql = "update frauen " & GetAktenvermerk(strlog, True) & sql_from
    '        db_con.FireSQL(sql, trans)
    '    End If


    'End Sub

    Private Function GetAktenvermerk(strlog As String, Optional singlefld As Boolean = False) As String
        Return If(singlefld, "", ",") & " aktenvermerke=(case when aktenvermerke IS null then " & strlog & " else aktenvermerke+ CHAR(13)+CHAR(10)+CHAR(13)+CHAR(10)+" & strlog & " end)  "
    End Function

    Public Function GetArztNameFromArztNr(ByVal Arztnr As Integer, ByVal trans As SqlClient.SqlTransaction) As String
        Dim db_con As New cls_db_con
        Dim tb As DataTable = db_con.GetRecordSet("select nname+' '+vname from ghdaten..aerzteliste where arztnr=" & Arztnr, trans)
        If tb.Rows.Count = 0 Then

            Return Arztnr
        End If
        Return tb.Rows(0)(0).ToString

    End Function
    Private Function GetTelField(TelNr As String, mobile_vorwahlen As String) As String

        If IsMobileNumber(GetMSISDNFormat(TelNr), mobile_vorwahlen) Then
            Return "Mobiltel"
        Else
            Return "Telefon"
        End If

    End Function


    Public Function IsMobileNumber(TelNr_MSISDNFormat As String, mobile_vorwahlen As String) As Boolean

        Dim vw_mobil As List(Of String) = mobile_vorwahlen.Split(",").ToList

        Dim vw As String
        Dim pos As Integer = 0

        Dim tmp As String = TelNr_MSISDNFormat.Substring(2)




        If tmp.StartsWith("0") Then
            pos = 1
            If tmp.Length < 4 Then Return False

        Else
            If tmp.Length < 3 Then Return False

        End If
        vw = tmp.Substring(pos, 3)

        If vw_mobil.Contains(vw) Then
            Return True
        Else
            Return False
        End If

    End Function
    Public Function GetMSISDNFormat(TelNr As String) As String
        'Zuerst alle Zeichen außer Ziffern entfernen
        Dim sb As New StringBuilder


        For i = 0 To TelNr.Length - 1
            If IsNumeric(TelNr.Substring(i, 1)) Then sb.Append(TelNr.Substring(i, 1))
        Next


        Dim tmp As String = sb.ToString
        If tmp.StartsWith("0043") Then Return tmp.Substring(2)
        If tmp.StartsWith("043") Then Return tmp.Substring(1)
        If tmp.StartsWith("0") Then
            Return "43" & tmp.Substring(1)
        Else
            If tmp.StartsWith("43") Then
                Return tmp
            Else
                Return tmp & "43" & tmp
            End If
        End If


    End Function
    Private Function SyncLog() As Boolean
        Dim db_con As New cls_db_con
        Try


            Dim pers As DataTable = db_con.GetRecordSet("Select * from os_synclog where os_commit=0 And os_dstyp=" & CInt(DS_Typ.Impfling))
            Dim vacc As DataTable = db_con.GetRecordSet("Select * from os_synclog where os_commit=0 And os_dstyp=" & CInt(DS_Typ.Impfung))
            Dim aa As DataTable = db_con.GetRecordSet("Select * from os_synclog where os_commit=0 And os_dstyp=" & CInt(DS_Typ.Amtsarzt))
            Dim ais As DataTable = db_con.GetRecordSet("Select * from os_synclog where os_commit=0 And os_dstyp=" & CInt(DS_Typ.AIS_Data_log))
            pers.TableName = TAB_PERS
            vacc.TableName = TAB_VACC
            aa.TableName = TAB_AA
            ais.TableName = TAB_AIS


            If pers.Rows.Count = 0 And vacc.Rows.Count = 0 And aa.Rows.Count = 0 And ais.Rows.Count = 0 Then Return False


            Dim ds As New DataSet
            ds.Tables.Add(pers)
            ds.Tables.Add(vacc)
            ds.Tables.Add(aa)
            ds.Tables.Add(ais)


            Dim ret As Boolean = SetRemoteParam(URL_Online_Service & "/configex/SetParam", ConfigParam.SyncVaccData, ToBase64(ds.GetXml), False)
            If ret Then
                Try

                    AktionsLog("Fehler beim Loggen der bereits integrierten Impfdaten im Online-Service.", AktionslogKat.Integration_BH_Impfungen)
                    Console.WriteLine("Fehler beim Loggen der bereits integrierten Impfdaten im Online-Service.")
                Catch ex As Exception

                End Try
                Return True
            End If
            db_con.FireSQL("update os_synclog Set os_commit=1 where os_commit=0")
            Console.WriteLine("Remote Logrecords auf os_commit=1 gesetzt.")
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
        Dim settings As New Einstellungen
        Dim strSQL As String
        Dim Charge = "NULL"
        If Not (String.IsNullOrEmpty(vacc.charge) Or vacc.charge = KEINE_CHARGE) Then Charge = "'" & vacc.charge & "'"


        Dim BHNR = "NULL"
        If vacc.bhnr <> NO_VAL Then
            BHNR = vacc.bhnr
        End If

        Dim RsnNobilling As String = "'BH-Impfung, Online'"
        Dim bRsnNobilling As String = "-1"
        Dim Satz As String = "0"
        Dim hsatz As New HASatz

        'Console.WriteLine("vacc.prog: " & vacc.prog)
        'Console.WriteLine("vacc.bhnr: " & vacc.bhnr)


        If vacc.prog = Programme.GrippeImpfung65Plus And vacc.bhnr = NO_VAL Then
            Satz = settings.Grippeimpfung65Plus_Honorar.ToString.Replace(",", ".")
            RsnNobilling = "NULL"
            bRsnNobilling = "0"
            hsatz = GetHASätze(vacc.arztnr, vacc.heftnr, vacc.prog, trans)
        End If



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



        'Arztnr - Amtsarztdublette? - Umbuchen !

        Dim tb As DataTable = db_con.GetRecordSet("select original from ghdaten..aa_dubletten where dublette=" & vacc.arztnr, trans)
        If tb.Rows.Count > 0 Then vacc.arztnr = tb.Rows(0)("original")


        'If Not Einstellungen.IsTestEnvironment Then

        Dim wp As Integer = 0
        Dim wp_satz As String = "0"
        Dim wp_mwst As String = "0"



        If vacc.wegpauschale Then
            wp = 1
            wp_satz = settings.Grippeimpfung65Plus_Wegpauschale.ToString.Replace(",", ".")
            wp_mwst = settings.Grippeimpfung65Plus_Wegpauschale_Mwst.ToString.Replace(",", ".")
        End If




        'Console.WriteLine("vacc.wegpauschale: " & vacc.wegpauschale)
        'Console.WriteLine("WP: " & wp)
        'Console.WriteLine("wp_satz: " & wp_satz)
        'Console.WriteLine("wp_mwst: " & wp_mwst)
        'Console.WriteLine("Satz: " & settings.Grippeimpfung65Plus_Wegpauschale)


        If String.IsNullOrEmpty(vacc.plid) Then


            'Dim Username As String

            'If Not String.IsNullOrEmpty(vacc.officialusername) Then
            '    Username = vacc.officialusername
            'Else
            '    Username = vacc.geandert
            'End If

            strSQL = "insert into ghdaten..impfdoku  (" &
                         "datum,nname,vname,gebdat,chargenr," &
                         "arztnr,serum,bis6,eingang,geandert," &
                         "geandertam,satz,mwsthon,aposatz,apomwst," &
                         "wegpauschale,wegpauschale_satz,wegpauschale_mwst," &
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
                         Satz & ",0," & hsatz.satz.ToString.ToString.Replace(",", ".") & "," & hsatz.mwst.ToString.Replace(",", ".") & "," &
                         wp & "," &
                         wp_satz & "," &
                         wp_mwst & "," &
                         vacc.heftnr & ", " &
                         bc & "," &
                         vacc.impfung & "," &
                          bRsnNobilling & "," & RsnNobilling & "," &
                         vacc.prog & "," &
                         BHNR & "," &
                         standort & "," &
                         standortid &
                         ")"


            Try
                db_con.FireSQL(strSQL, trans)

            Catch ex As Exception
                If Not ex.Message.IndexOf("PRIMARY KEY", 0) > 0 Then
                    Throw New Exception(ex.Message)
                Else
                    AktionsLog("Duplkat Impfung wurde übergangen: " & strSQL, AktionslogKat.Integration_BH_Impfungen, trans)
                    Console.WriteLine("Duplkat Impfung wurde übergangen.")
                End If
            End Try


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
                        "satz=" & Satz & ",mwsthon=0,aposatz=0,apomwst=0," &
                        "heftnr=" & vacc.heftnr & ", " &
                        "boncode=" & bc & "," &
                        "impfung=" & vacc.impfung & "," &
                        "NoBilling=" & bRsnNobilling & ",RsnNoBilling=" & RsnNobilling & "," &
                         "wegpauschale=" & wp & "," &
                         "wegpauschale_satz=" & wp_satz & "," &
                         "wegpauschale_mwst=" & wp_mwst & "," &
                        "Prog=" & vacc.prog & "," &
                        "bhnr=" & BHNR & "," &
                        "standort=" & standort & "," &
                        "standortid=" & standortid & " " &
                        "where datum='" & s(0) & "' and boncode=" & bc

                db_con.FireSQL(strSQL, trans)


                    AktionsLog("Impfung wurde verändert: " & strSQL, AktionslogKat.Integration_BH_Impfungen, trans)


                End If

            End If



        'End If

        Return bc

    End Function
    Private Function GetBoncode(Heftnr As Integer, Impfstoff As Integer, Impfung As Integer) As String
        Return Heftnr & Format(CInt(Impfstoff), "00") & Impfung
    End Function


    Private Function GetHASätze(Arztnr As Integer, Hefrtnr As Integer, prog As Integer, trans As SqlClient.SqlTransaction) As HASatz

        Dim hs As New HASatz
        Dim settings As New Einstellungen
        If IsDOCHausapotheker(Arztnr, trans) Then
            If IsMobilerDienstIMpfling(Hefrtnr, trans) Then
                If Not HAKontingentVerbraucht(Arztnr, prog, trans) Then
                    hs.satz = settings.Grippeimpfung65Plus_HASatz
                    hs.mwst = settings.Grippeimpfung65Plus_HAMwst
                End If
            End If
        End If

        Return hs


    End Function


    Private Function HAKontingentVerbraucht(Arztnr As Integer, Prog As Integer, trans As SqlClient.SqlTransaction) As Boolean
        Dim db_con As New cls_db_con
        Dim kontingent As Integer = 0
        Dim verbraucht As Integer = 0
        Dim tb_kontinget As DataTable = db_con.GetRecordSet("select sum(anzahl) from gi65p_einleselog where apotheke=" & Arztnr & " and prog=" & Prog, trans)
        Dim tb_verbraucht As DataTable = db_con.GetRecordSet("select count(*) from impfdoku where arztnr=" & Arztnr & " and prog=" & Prog, trans)


        If Not IsDBNull(tb_kontinget.Rows(0)(0)) Then kontingent = tb_kontinget.Rows(0)(0)
        If Not IsDBNull(tb_verbraucht.Rows(0)(0)) Then verbraucht = tb_verbraucht.Rows(0)(0)



        If verbraucht >= kontingent Then
            Return True
        Else
            Return False
        End If


    End Function

    Private Function IsMobilerDienstIMpfling(heftnr As Integer, trans As SqlClient.SqlTransaction) As Boolean
        Dim db_con As New cls_db_con

        Dim strSQL As String = "select betreuendestelle from ghdaten..frauen where heftnrf=" & heftnr

        Console.WriteLine("IsMobilerDienstIMpfling: " & strSQL)

        Dim tb As DataTable = db_con.GetRecordSet("select betreuendestelle from ghdaten..frauen where heftnrf=" & heftnr, trans)

        If tb.Rows.Count > 0 Then
            If IsDBNull(tb.Rows(0)(0)) Then Return False

            If tb.Rows(0)(0) >= MD_TRÄGER_ID_MIN Then
                Return True
            Else
                Return False

            End If

        Else
            Return False

        End If



    End Function

    Private Function HausapothekeAktuell(von As String, bis As String) As Boolean

        If von = "" Then Return False


        If Date.Compare(CDate(von), Date.Today) <= 0 Then

            If bis = "" Then
                Return True
            Else
                If Date.Compare(CDate(bis), Date.Today) >= 0 Then
                    Return True

                Else
                    Return False

                End If



            End If


        End If

        Return False

    End Function

    Private Function IsDOCHausapotheker(arztnr As Integer, trans As SqlClient.SqlTransaction) As Boolean
        Dim db_con As New cls_db_con

        Dim tb As DataTable = db_con.GetRecordSet("select ha_begin, ha_ende from aerzteliste where arztnr=" & arztnr, trans)
        Dim von As String = ""
        Dim bis As String = ""
        If Not IsDBNull(tb.Rows(0)("ha_begin")) Then von = tb.Rows(0)("ha_begin")
        If Not IsDBNull(tb.Rows(0)("ha_ende")) Then bis = tb.Rows(0)("ha_ende")

        If HausapothekeAktuell(von, bis) Then
            Return True
        Else
            Return False
        End If


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

        If pat.KindUpdateMöglich Then
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
            gebdat = "'" & GetGBShort(pat.gebdat) & "'"
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
            Console.WriteLine("KindUpdateMöglich: Select heftnrf from frauen where heftnrf=" & pat.heftnr_os)
            tb = db_con.GetRecordSet("Select heftnrf from frauen where heftnrf=" & pat.heftnr_os, trans)


            If tb.Rows.Count > 0 Then Return False






            'Exisitiert Kind-DS?
            'Wir können aus einem Kind-DS nicht den Erwachsenen ableiten!

            Console.WriteLine("KindUpdateMöglich: Select heftnr from eeinh where heftnr=" & pat.heftnr_os)
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

        Console.WriteLine("heftnr=" & heftnr)

        If Not hasSVN And heftnr > 0 Then
            'Kind mit SVN?
            Console.WriteLine("1")

            tb = db_con.GetRecordSet("Select SVNRKIND, k_gbdatum from eeinh where heftnr=" & heftnr, trans)
            If tb.Rows.Count > 0 Then

                Console.WriteLine("1a")

                If Not IsDBNull(tb.Rows(0)("SVNRKIND")) And Not IsDBNull(tb.Rows(0)("k_gbdatum")) Then
                    Console.WriteLine("1b")


                    pat.svn = tb.Rows(0)("SVNRKIND") & tb.Rows(0)("k_gbdatum")
                    Console.WriteLine("1c")


                    hasSVN = True
                End If
            End If
            Console.WriteLine("2")

        End If

        Console.WriteLine("3")


        Dim ic As New ImportGeocodierung(pat.strasse & " " & pat.hnr, pat.plz, pat.ort)


        Console.WriteLine("4")


        Dim svnrdata As String = "NULL"
        Dim gebdat As String = "NULL"

        If hasSVN Then
            svnrdata = "'" & pat.svn.Substring(0, 4) & "'"
            gebdat = "'" & pat.svn.Substring(4, 6) & "'"
        Else
            gebdat = "'" & GetGBShort(pat.gebdat) & "'"
        End If

        Console.WriteLine("5")


        Dim sql As String = "insert into frauen (heftnrf,anrede,nachname,vorname,gebdat,gbjahr,strasse,plz,ort,svnrdata," &
                        "gbdatum,sva,kreiertam,geandertam,AlterAddress,kreiert,geandert, " &
                        "nneltern,vneltern,titeleltern,titel_suffix_eltern," &
                        "breitengrad,laengengrad," &
                        "GeocodeLevel,gemeindeid) values ( " &
                       IIf(heftnr = NO_VAL, "NULL", heftnr) & "," &
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


        Console.WriteLine("6")


        db_con.FireSQL(sql, trans)

        Console.WriteLine("7")


        AktionsLog("Neuer Impfling: " & sql, AktionslogKat.Integration_BH_Impfungen, trans)

        Console.WriteLine("8")


        If heftnr = NO_VAL Then
            Console.WriteLine("8a")


            heftnr = InsertHeftnr(pat,, trans)

            Console.WriteLine("8b")


        End If

        Console.WriteLine("9")

        Return heftnr

    End Function

    Public Function GetGBShort(Gebdat As DateTime) As String
        Return Gebdat.Day.ToString.PadLeft(2, "0") &
            Gebdat.Month.ToString.PadLeft(2, "0") &
            Gebdat.Year.ToString.Substring(2, 2)

    End Function

    Private Function InsertHeftnr(pat As OSImpfling, Optional SVNRID As Integer = NO_VAL, Optional trans As SqlClient.SqlTransaction = Nothing) As Integer
        Dim db_con As New cls_db_con
        If SVNRID = NO_VAL Then SVNRID = getSVNRID(pat, trans)
        Dim heftnr As Integer = HEFTBASIS + SVNRID

        Dim sql As String = "update frauen set heftnrf=" & heftnr & " where svnr_id=" & SVNRID

        db_con.FireSQL(sql, trans)

        AktionsLog("Neue Impf-ID: " & sql, AktionslogKat.Integration_BH_Impfungen, trans)

        Return heftnr
    End Function





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

            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

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
    Public aa As List(Of amtsarzt_voll)
    Public Vacc As List(Of Idata)
    Public pers As List(Of OSImpfling)
    Public ais As List(Of ais_data_log)
    Public mobile_vorwahlen As String
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


    Public Property KindUpdateMöglich As Boolean = False
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
    Public Property officialusername As String

    Public Property geandertam As Date

    Public Property prog As Integer
    Public Property plid As String = ""

    Public Property del As Integer = 0
    Public Property wegpauschale As Boolean = False


End Class


Public Class ais_data_log
    Public Property id As Integer
    Public Property arztnr As Integer
    Public Property json As String
    Public Property datum As Date
    Public Property inplattform As Short

End Class

Public Class AIS_Update

    Public Enum UpdateTyp

        Adresse = 1
        NameEltern = 2
        NameImpfling = 3
        TelEltern = 4
        TelImpfling = 5
        EmailEltern = 6
        EmailImpfling = 7
        SVNEltern = 8
        SVNImpfling = 9
    End Enum

    Public Property Typ As UpdateTyp


    Public Property text As String


    Public Property Strasse As String

    Public Property PLZ As String
    Public Property Ort As String


    Public Property Titel_Pra As String
    Public Property Titel_Suf As String
    Public Property Vorname As String
    Public Property Nachname As String


    Public Property heftnr As Integer

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




Public Class amtsarzt_voll

    Private m_NN As String
    Private m_VN As String
    Private m_Titel As String
    Private m_Strasse As String
    Private m_PLZ As Integer
    Private m_Ort As String
    Private m_Email As String
    Private m_BH As String
    Private m_BHNR As Integer
    Private m_BPK As String
    Private m_Tel As String


    Public Arztnr As Integer

    'Notwendig für Deserialisierung
    Public Sub New()

    End Sub

    Public Sub New(strNN As String, strVN As String, strTitel As String, strStrasse As String, iPLZ As Integer, strOrt As String, strEmail As String, strBH As String, iBHNR As Integer, strBPK As String, strTel As String, iArztnr As Integer)


        NN = strNN
        VN = strVN
        Titel = strTitel
        Strasse = strStrasse
        PLZ = iPLZ
        Ort = strOrt
        eMail = strEmail
        BH = strBH
        BHNR = iBHNR
        BPK = strBPK
        Tel = strTel

        Arztnr = iArztnr

    End Sub


    Public Property NN As String
        Get
            Return m_NN.Replace("'", "''")
        End Get
        Set(value As String)
            m_NN = value
        End Set
    End Property


    Public Property VN As String
        Get
            Return m_VN.Replace("'", "''")

        End Get
        Set(value As String)
            m_VN = value
        End Set
    End Property


    Public Property Titel As String
        Get
            Return m_Titel.Replace("'", "''")

        End Get
        Set(value As String)
            m_Titel = value
        End Set
    End Property

    Public Property Strasse As String
        Get
            Return m_Strasse.Replace("'", "''")

        End Get
        Set(value As String)
            m_Strasse = value
        End Set
    End Property


    Public Property PLZ As Integer
        Get
            Return m_PLZ

        End Get
        Set(value As Integer)
            m_PLZ = value
        End Set
    End Property

    Public Property Ort As String
        Get
            Return m_Ort.Replace("'", "''")

        End Get
        Set(value As String)
            m_Ort = value
        End Set
    End Property


    Public Property eMail As String
        Get
            Return m_Email.Replace("'", "''")

        End Get
        Set(value As String)
            m_Email = value
        End Set
    End Property


    Public Property BH As String
        Get
            Return m_BH.Replace("'", "''")

        End Get
        Set(value As String)
            m_BH = value
        End Set
    End Property


    Public Property BHNR As Integer
        Get
            Return m_BHNR

        End Get
        Set(value As Integer)
            m_BHNR = value
        End Set
    End Property

    Public Property BPK As String
        Get
            Return m_BPK

        End Get
        Set(value As String)
            m_BPK = value
        End Set
    End Property

    Public Property Tel As String
        Get
            Return m_Tel.Replace("'", "''")

        End Get
        Set(value As String)
            m_Tel = value
        End Set
    End Property
End Class

Public Class ghdb_amtsarzt

    Private Const HERR As String = "Herrn"
    Private Const FRAU As String = "Frau"

    Private m_con As SqlClient.SqlConnection = Nothing
    Private m_trans As SqlClient.SqlTransaction = Nothing
    Private db_con As New cls_db_con
    Public Sub Add(aa As amtsarzt_voll, DBOnline As Boolean, trans As SqlClient.SqlTransaction)
        'DBonline=true: wie befundne uns im OS
        'DBonline=false: wie befunden uns im Sync-Programm



        Console.WriteLine("Amtsarzt hinzufügen 1")




        m_trans = trans



        BeginTrans()


        Console.WriteLine("Amtsarzt hinzufügen 2")



        Dim BH As String = ""
        If aa.Arztnr <= 0 Then
            Throw New Exception("Arztnummer des neuen Amtsarztes (" & aa.NN & ") ist unbekannt.")
        End If

        BH = aa.BH

        If Not Einstellungen.IsTestEnvironment Then

            Console.WriteLine("Amtsarzt hinzufügen 3, Arztnr=" & aa.Arztnr)


            Dim strSQL As String

            'Nur Update?
            Dim tb_test As DataTable = db_con.GetRecordSet("select arztnr from ghdaten..aerzteliste where arztnr=" & aa.Arztnr, m_trans)
            If tb_test.Rows.Count > 0 Then

                strSQL = "update  ghdaten..aerzteliste set pvp_gid='" & aa.BPK & "' where arztnr=" & aa.Arztnr




            Else
                'Insert

                strSQL = "insert into ghdaten..AERZTELISTE (arztnr, gruppe, [name], grdtit, anrede, nname, vname, [Str], plz, ort, email, BH, bhnr, TEILNIMPF, PVP_GID,tel) values(" &
                        aa.Arztnr & "," &
                        "'Amtsarzt'," &
                        "'" & aa.NN & " " & aa.VN & "'," &
                        "'" & aa.Titel & "'," &
                        "'" & GetAnrede(aa.VN) & "'," &
                        "'" & aa.NN & "'," &
                        "'" & aa.VN & "'," &
                        "'" & aa.Strasse & "'," &
                        aa.PLZ & "," &
                        "'" & aa.Ort & "'," &
                        "'" & aa.eMail & "'," &
                        "'" & BH & "'," &
                        aa.BHNR & "," &
                        "1," &
                        "'" & aa.BPK & "'," &
                        "'" & aa.Tel & "')"



            End If



            db_con.FireSQL(strSQL, m_trans)

            Try
                Dim osync As New sync


                If tb_test.Rows.Count > 0 Then
                    osync.AktionsLog("BPK für Amtsarzt " & aa.Arztnr & " wurde eingetragen.", sync.AktionslogKat.Integration_BH_Impfungen, m_trans)
                Else
                    osync.AktionsLog("Neuer Amtsarzt " & aa.Arztnr & " wurde erstellt.", sync.AktionslogKat.Integration_BH_Impfungen, m_trans)

                End If
            Catch ex As Exception

            End Try



        End If











    End Sub



    Private Function GetAnrede(ByVal VN As String) As String



        If String.IsNullOrEmpty(VN) Then Return ""

        BeginTrans()


        Dim tb As DataTable = db_con.GetRecordSet("select sex from gemeinden..vornamendb where vname='" & VN & "'", m_trans)
        If tb.Rows.Count = 0 Then
            Return FRAU
        Else
            If tb.Rows(0)(0) = "0" Then
                Return HERR
            Else
                Return FRAU
            End If

        End If
    End Function


    'Private Function NeueArztnr() As Integer
    '    BeginTrans()

    '    Dim db_con As New cls_db_con
    '    Dim tb As DataTable = db_con.GetRecordSet("select isnull(max(arztnr),8000000)+1 from aerzteliste where arztnr between 8000000 and 8999999", m_trans)
    '    Return tb.Rows(0)(0)


    'End Function

    Private Sub BeginTrans()
        If m_trans Is Nothing Then
            m_con.Open()
            m_trans = m_con.BeginTransaction
        End If

    End Sub


End Class


Friend Class HASatz
    Public satz As Single = 0
    Public mwst As Single = 0
End Class