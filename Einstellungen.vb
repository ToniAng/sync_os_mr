Option Compare Text
Imports System
Imports System.IO
Imports System.Security.Cryptography


Public Class Einstellungen

    Dim tb As DataTable
    Dim d As DataView
    Dim cmd As New SqlClient.SqlCommand
    Dim dr As DataRow
    Dim Isloaded As Boolean = False
    Dim IsLoading As Boolean = False

    Public Const NO_VAL As Integer = -1


    Public Const DEF_NAME_STELLE As String = "Stelle"
    Public Const DEF_STR As String = "Strasse"
    Public Const DEF_PLZ As String = "9999"
    Public Const DEF_ORT As String = "Ort"
    Public Const DEF_INSTITUT As String = "Institut"
    Public Const DEF_BLZ As String = "99999"
    Public Const DEF_KONTO As String = "12345678901"
    Public Const DEF_Tel As String = ""
    Public Const DEF_FAX As String = ""
    Public Const DEF_EMAIL As String = ""
    Public Const DEF_V3_STELLE As String = ""
    Public Const DEF_V3_INSTITUT As String = ""
    Public Const DEF_MAXSCHULALTER As Integer = 18
    Public Const DEF_ABLAGE As String = "Standardablage"
    Public Const DEF_AR_INTERVALL As Integer = 60
    Public Const DEF_FIBUKONTO_HA As String = ""
    Public Const DEF_FIBUKONTO_HONORARE As String = ""
    Public Const DEF_IMPF_TIMEOUT As Date = #1/1/2002#
    Public Const DEF_IMPF_TIMEOUT_REL As Integer = 10
    Public Const DEF_GEOREF_TN As Date = #1/1/1990#
    Public Const DEF_LOAD_PUPIL As Integer = 0
    Public Const DEF_MAXKINDESALTER As Integer = 18
    Public Const DEF_MAXFAX As Integer = 2
    Public Const DEF_MISS_DBL_TIMEOUT As Integer = 36
    Public Const DEF_SMTP_SERVER As String = ""
    Public Const DEF_OLAP_WS As String = "http://localhost/olap_data/analysis.asmx"
    Public Const DEF_GEOCODER As String = "http://localhost/geocoder/geoservice.asmx"
    Public Const DEF_MM_START As String = "00:00"
    Public Const DEF_FOLDER_SCHOOLDATA As String = "C:\"
    Public Const OL_PRIV_ENTWUERFE = "Private Entwürfe"
    Public Const NO_STD_PRINTER As String = "-----"

    Public Const INFO1 As String = "U29ycnksIGFiZXIgZGFzIGhhdCBrZWluZSBCZWRldXR1bmch"

    Public Const DEF_INITFILEDIR As String = "\\wissakserv\gemeinsamedaten"



    Public Const MAXTREFFER As Integer = 30

    'Default Altersgrenzen Mammographie
    Public Const MM_MIN_ALTER As Integer = 39
    Public Const MM_MAX_ALTER As Integer = 70

    'Background
    Public Const DEF_BACKG_R As Integer = 225
    Public Const DEF_BACKG_G As Integer = 255
    Public Const DEF_BACKG_B As Integer = 225

    Private db_con As New cls_db_con


    Public Const TESTMACHINE As String = "PROXIMACENTAURI"
    Public Const TESTMACHINE2 As String = "70-OPHIUCHI"



    Public Enum FarbPalette
        MultidimensionalAnalyse = 1
        Geoanalysen = 2
    End Enum


    Public Enum ExchangeVersion
        Exchange_2000
        Exchange_2003
        Exchange_2007

    End Enum

    Public ReadOnly Property CurrentUser() As String
        Get
            Return System.Environment.UserName
        End Get
    End Property

    'Private Function GHDruck_GetDefFont() As String
    '    Select Case ID_VERSION
    '        Case ID_AVOS
    '            Return "Avantgarde-MEDIUM"
    '        Case Else
    '            Return "Arial"
    '    End Select
    'End Function
    Public Shared Function IsTestEnvironment() As Boolean
        If Environment.MachineName.ToUpper = TESTMACHINE.ToUpper Or Environment.MachineName.ToUpper = TESTMACHINE2.ToUpper Then
            Return True
        Else
            Return False
        End If



    End Function

    Public Sub LoadConfig(Optional ByVal trans As SqlClient.SqlTransaction = Nothing)
        Dim x As Integer
        Dim m_Ort As String
        If IsLoading And Not Isloaded Then
            'Ein anderer Thread greift zu
            'Do
            '    Threading.Thread.Sleep(10)
            '    If Isloaded Then Exit Do
            'Loop
            'Debug.Write("Threadmanagement Setting aufgerufen" & Chr(13))

            Return
        End If
        IsLoading = True
        If Isloaded Then Return
        db_con.GetCon()

        tb = db_con.GetRecordSet("select * from config", trans)

        Dim tb1 As DataTable
        tb1 = db_con.GetRecordSet("select zahler, anschrift, blz, institut, konto, tel, fax, email, schuljahr, v3_stelle, v3_institut, maxschulalter, nbquiet, askti, lockImpfdoku, pgpuserid, hontext,hontext1,hontext2,dokerf,kritdur, txtbeflist, txtaftlist , impfdokumissing from stelle", trans)
        AddRowToD("zahler", CStr(tb1.Rows(0)("zahler")))



        x = InStr(1, tb1.Rows(0)("anschrift").ToString, Chr(13))
        AddRowToD("STR", Left(tb1.Rows(0)("anschrift").ToString, x - 1))
        m_Ort = Right(tb1.Rows(0)("anschrift").ToString, Len(tb1.Rows(0)("anschrift")) - x - 1)
        AddRowToD("PLZ", Left(m_Ort, 4).Trim)
        AddRowToD("Ort", Right(m_Ort, Len(m_Ort) - 4).Trim)
        AddRowToD("institut", tb1.Rows(0)("institut").ToString)
        AddRowToD("BLZ", tb1.Rows(0)("blz").ToString)
        AddRowToD("Konto", tb1.Rows(0)("Konto").ToString)
        AddRowToD("Tel", tb1.Rows(0)("tel").ToString)
        If IsDBNull(tb1.Rows(0)("FAX")) Then
            AddRowToD("Fax", "")
        Else
            AddRowToD("Fax", tb1.Rows(0)("fax").ToString)
        End If
        If IsDBNull(tb1.Rows(0)("email")) Then
            AddRowToD("email", "")
        Else
            AddRowToD("email", tb1.Rows(0)("email").ToString)
        End If
        If IsDBNull(tb1.Rows(0)("schuljahr")) Then
            AddRowToD("schuljahr", "")
        Else
            AddRowToD("schuljahr", tb1.Rows(0)("schuljahr").ToString)
        End If
        If IsDBNull(tb1.Rows(0)("v3_stelle")) Then
            AddRowToD("v3_stelle", "")
        Else
            AddRowToD("v3_stelle", tb1.Rows(0)("v3_stelle").ToString)
        End If
        If IsDBNull(tb1.Rows(0)("v3_institut")) Then
            AddRowToD("v3_institut", "")
        Else
            AddRowToD("v3_institut", tb1.Rows(0)("v3_institut").ToString)
        End If
        If IsDBNull(tb1.Rows(0)("maxschulalter")) Then
            AddRowToD("maxschulalter", "")
        Else
            AddRowToD("maxschulalter", tb1.Rows(0)("maxschulalter").ToString)
        End If
        If IsDBNull(tb1.Rows(0)("nbquiet")) Then
            AddRowToD("nbquiet", "0")
        Else
            AddRowToD("nbquiet", tb1.Rows(0)("nbquiet").ToString)
        End If
        If IsDBNull(tb1.Rows(0)("askti")) Then
            AddRowToD("askti", "0")
        Else
            AddRowToD("askti", tb1.Rows(0)("askti").ToString)
        End If
        If IsDBNull(tb1.Rows(0)("lockImpfdoku")) Then
            AddRowToD("lockImpfdoku", "")
        Else
            AddRowToD("lockImpfdoku", tb1.Rows(0)("lockImpfdoku").ToString)
        End If
        If IsDBNull(tb1.Rows(0)("pgpuserid")) Then
            AddRowToD("pgpuserid", "")
        Else
            AddRowToD("pgpuserid", tb1.Rows(0)("pgpuserid").ToString)
        End If
        If IsDBNull(tb1.Rows(0)("hontext")) Then
            AddRowToD("hontext", "")
        Else
            AddRowToD("hontext", tb1.Rows(0)("hontext").ToString)
        End If
        If IsDBNull(tb1.Rows(0)("hontext1")) Then
            AddRowToD("hontext1", "")
        Else
            AddRowToD("hontext1", tb1.Rows(0)("hontext1").ToString)
        End If
        If IsDBNull(tb1.Rows(0)("hontext2")) Then
            AddRowToD("hontext2", "")
        Else
            AddRowToD("hontext2", tb1.Rows(0)("hontext2").ToString)
        End If
        If IsDBNull(tb1.Rows(0)("dokerf")) Then
            AddRowToD("dokerf", "")
        Else
            AddRowToD("dokerf", tb1.Rows(0)("dokerf").ToString)
        End If
        If IsDBNull(tb1.Rows(0)("kritdur")) Then
            AddRowToD("kritdur", "")
        Else
            AddRowToD("kritdur", tb1.Rows(0)("kritdur").ToString)
        End If


        If IsDBNull(tb1.Rows(0)("txtbeflist")) Then
            AddRowToD("txtbeflist", "")
        Else
            AddRowToD("txtbeflist", tb1.Rows(0)("txtbeflist").ToString)
        End If
        If IsDBNull(tb1.Rows(0)("txtaftlist")) Then
            AddRowToD("txtaftlist", "")
        Else
            AddRowToD("txtaftlist", tb1.Rows(0)("txtaftlist").ToString)
        End If

        If IsDBNull(tb1.Rows(0)("impfdokumissing")) Then
            AddRowToD("impfdokumissing", "")
        Else
            AddRowToD("impfdokumissing", tb1.Rows(0)("impfdokumissing").ToString)
        End If



        d = New DataView(tb)



        Isloaded = True
    End Sub


    Private Sub SetProp(ByVal param As String, ByVal v As String, Optional ByVal Trans As SqlClient.SqlTransaction = Nothing, Optional AlterCon As Boolean = False)
        Dim SQLStr As String = ""
        If Not Isloaded Then LoadConfig(Trans)

        d.RowFilter = "param='" & param & "'"
        If d.Count = 0 Then
            SQLStr = "insert into config (param, val) values ('" & param & "','" & v & "') "
            'db_con.FireSQL(SQLStr, Trans)
            SetPropSQL(SQLStr, Trans, AlterCon)
            AddRowToD(param, v)

        Else
            SQLStr = "update config set val='" & v & "' where param='" & param & "' "
            'db_con.FireSQL(SQLStr, Trans)
            SetPropSQL(SQLStr, Trans, AlterCon)

            If v = "NULL" Then v = ""
            d.RowFilter = "param='" & param & "'"
            d(0)(1) = v
        End If

    End Sub

    Private Sub SetPropSQL(SQLStr As String, Optional ByVal Trans As SqlClient.SqlTransaction = Nothing, Optional AlterCon As Boolean = False)
        If AlterCon Then
            db_con.FireSQL(SQLStr, Trans, c)

        Else
            db_con.FireSQL(SQLStr, Trans)

        End If
    End Sub


    Private Sub SetProp_Old(ByVal param As String, ByVal v As String, Optional ByVal Dec As Boolean = False, Optional ByVal trans As SqlClient.SqlTransaction = Nothing)
        Dim SQLStr As String = ""
        If Not Isloaded Then LoadConfig(trans)

        If v = "" Then v = "NULL" Else v = "'" & v & "'"
        If param = "ort" Or param = "str" Or param = "plz" Then

            If param = "str" Then SQLStr = "update stelle set anschrift='" & v.Replace("'", "") & vbCrLf & Me.PLZ & " " & Me.Ort & "' "
            If param = "plz" Then SQLStr = "update stelle set anschrift='" & Me.Strasse & vbCrLf & v.Replace("'", "") & " " & Me.Ort & "' "
            If param = "ort" Then SQLStr = "update stelle set anschrift='" & Me.Strasse & vbCrLf & Me.PLZ & " " & v.Replace("'", "") & "' "

        Else
            If Dec Then v = v.Replace(",", ".")
            SQLStr = "update stelle set " & param & "=" & v & " "
            If Dec Then v = v.Replace(".", ",")

        End If

        db_con.FireSQL(SQLStr, trans)
        'd(0)(0) = param
        If v = "NULL" Then v = ""
        d.RowFilter = "param='" & param & "'"
        d(0)(1) = v.Replace("'", "")

    End Sub

    Private Sub AddRowToD(ByVal param As String, ByVal v As String)
        dr = tb.NewRow
        dr(0) = param
        dr(1) = v
        tb.Rows.Add(dr)

    End Sub

    Public Property ColorPalette(ByVal Pal As FarbPalette) As System.Drawing.Color()


        Get
            Dim PalName As String
            Select Case Pal
                Case FarbPalette.Geoanalysen
                    PalName = "ColorPalette_GEO"
                Case FarbPalette.MultidimensionalAnalyse
                    PalName = "ColorPalette_MD"
                Case Else
                    Throw New Exception("Farbpalette wurde nicht definiert.")

            End Select
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='" & PalName & "'"
            Dim m_color(11) As System.Drawing.Color
            If d.Count > 0 Then

                Dim s(11) As String
                s = d(0)(1).ToString.Split(CChar("|"))
                For i As Integer = 0 To 11
                    m_color(i) = Drawing.Color.FromArgb(CInt(s(i)))
                Next
            Else
                m_color(0) = Drawing.Color.DarkBlue
                m_color(1) = Drawing.Color.Red
                m_color(2) = Drawing.Color.DarkViolet
                m_color(3) = Drawing.Color.Gold
                m_color(4) = Drawing.Color.Violet
                m_color(5) = Drawing.Color.Olive
                m_color(6) = Drawing.Color.Blue
                m_color(7) = Drawing.Color.DarkBlue
                m_color(8) = Drawing.Color.Cyan
                m_color(9) = Drawing.Color.Orange
                m_color(10) = Drawing.Color.DarkRed
                m_color(11) = Drawing.Color.Gold

                SetProp(PalName, Me.ColorArrayToString(m_color))
                Return m_color
            End If
            Return m_color
        End Get
        Set(ByVal Value As System.Drawing.Color())
            Dim PalName As String
            Select Case Pal
                Case FarbPalette.Geoanalysen
                    PalName = "ColorPalette_GEO"
                Case FarbPalette.MultidimensionalAnalyse
                    PalName = "ColorPalette_MD"
                Case Else
                    Throw New Exception("Farbpalette wurde nicht definiert.")

            End Select

            SetProp(palname, Me.ColorArrayToString(Value))
        End Set
    End Property


    Public Property Grippeimpfung65Plus_Impfstoff() As Integer
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Grippeimpfung65Plus_Impfstoff'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 0
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Integer)
            SetProp("Grippeimpfung65Plus_Impfstoff", Value)
        End Set
    End Property

    Public Property Grippeimpfung65Plus_Honorar() As Single
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Grippeimpfung65Plus_Honorar'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 0
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Single)
            SetProp("Grippeimpfung65Plus_Honorar", Value)
        End Set
    End Property

    Public Property Grippeimpfung65Plus_Wegpauschale() As Single
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Grippeimpfung65Plus_Wegpauschale'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 0
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Single)
            SetProp("Grippeimpfung65Plus_Wegpauschale", Value)
        End Set
    End Property


    Public Property HASatz() As Single
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Grippeimpfung65Plus_HASatz'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 2.0
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Single)
            SetProp("Grippeimpfung65Plus_HASatz", Value)
        End Set
    End Property
    Public Property HAMwst() As Single
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Grippeimpfung65Plus_HAMwst'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 0.2
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Single)
            SetProp("Grippeimpfung65Plus_HAMwst", Value)
        End Set
    End Property
    Public Property Grippeimpfung65Plus_Wegpauschale_Mwst() As Single
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Grippeimpfung65Plus_Wegpauschale_Mwst'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 0
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Single)
            SetProp("Grippeimpfung65Plus_Wegpauschale_Mwst", Value)
        End Set
    End Property

    Private Function ColorArrayToString(ByVal m_color() As System.Drawing.Color) As String
        Dim sb As New System.Text.StringBuilder
        For i As Integer = 0 To m_color.Length - 1
            sb.Append(m_color(i).ToArgb)
            If i < m_color.Length - 1 Then sb.Append("|")
        Next
        Return sb.ToString
    End Function

    Public Property Name_Stelle() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='zahler'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("zahler", DEF_NAME_STELLE)
                Return DEF_NAME_STELLE
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("zahler", CStr(Value))
        End Set
    End Property
    Public Property Geoanalysen_Font() As System.Drawing.Font
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Font_Geoanalysen'"
            If d.Count > 0 Then

                Dim s() As String = d(0)(1).ToString.Split(CChar("|"))
                Return New System.Drawing.Font(s(0), CInt(s(1)), Me.GetFontStylefromName(s(2)))

            Else
                SetProp("Font_Geoanalysen", "Aria|9|standard")
                Return New System.Drawing.Font("Arial", 9, System.Drawing.FontStyle.Regular)
            End If
        End Get
        Set(ByVal Value As System.Drawing.Font)
            SetProp("Font_Geoanalysen", Value.Name & "|" & Value.Size & "|" & Me.GetFontStyleName(Value.Style))
        End Set
    End Property
    Public Function GetFontStylefromName(ByVal StyleName As String) As System.Drawing.FontStyle
        Select Case StyleName.ToLower
            Case "kursiv", "italic"
                Return System.Drawing.FontStyle.Italic
            Case "fett", "bold"
                Return System.Drawing.FontStyle.Bold
            Case "unterstrichen", "underline"
                Return System.Drawing.FontStyle.Underline
            Case "durchgestrichen", "strikeout"
                Return System.Drawing.FontStyle.Strikeout
            Case Else
                Return System.Drawing.FontStyle.Regular


        End Select
    End Function
    Public Function GetFontStyleName(ByVal fs As System.Drawing.FontStyle) As String
        Select Case fs
            Case System.Drawing.FontStyle.Bold
                Return "fett"
            Case System.Drawing.FontStyle.Italic
                Return "kursiv"
            Case System.Drawing.FontStyle.Strikeout
                Return "durchgestrichen"
            Case System.Drawing.FontStyle.Underline
                Return "unterstrichen"
            Case Else
                Return "standard"

        End Select

    End Function

    Public Property Strasse() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='STR'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("STR", DEF_STR)
                Return DEF_STR
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("STR", CStr(Value))
        End Set
    End Property



    Public Property Ort() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Ort'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("Ort", DEF_ORT)
                Return DEF_ORT
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("Ort", CStr(Value))
        End Set
    End Property

    Public Property PLZ() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='PLZ'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp_Old("PLZ", DEF_PLZ)
                Return CInt(DEF_PLZ)
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp_Old("PLZ", CStr(Value))
        End Set
    End Property
    Public Property ThreadsSchuelerimport() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='ThreadsSchuelerimport'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp("ThreadsSchuelerimport", 10)
                Return CInt(10)
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp("ThreadsSchuelerimport", CStr(Value))
        End Set
    End Property
    Public Property Search_MaxTreffer() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Search_MaxTreffer" & "_" & UCase(Setting.CurrentUser) & "'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp("Search_MaxTreffer" & "_" & UCase(Setting.CurrentUser), MAXTREFFER.ToString)
                Return MAXTREFFER
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp("Search_MaxTreffer" & "_" & UCase(Setting.CurrentUser), CStr(Value))
        End Set
    End Property

    Public Property Institut() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Institut'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("Institut", DEF_INSTITUT)
                Return DEF_INSTITUT
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("Institut", CStr(Value))
        End Set
    End Property
    Public Property BLZ() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='BLZ'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("BLZ", DEF_BLZ)
                Return DEF_BLZ
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("BLZ", CStr(Value))
        End Set
    End Property

    Public Property Konto() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Konto'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("Konto", DEF_KONTO)
                Return DEF_KONTO
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("Konto", CStr(Value))
        End Set
    End Property
    Public Property Tel() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Tel'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("Tel", DEF_Tel)
                Return DEF_Tel
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("Tel", CStr(Value))
        End Set
    End Property
    Public Property Fax() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Fax'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("Fax", DEF_FAX)
                Return DEF_FAX
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("Fax", CStr(Value))
        End Set
    End Property

    Public Property EMail() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='EMail'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("EMail", DEF_EMAIL)
                Return DEF_EMAIL
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("EMail", CStr(Value))
        End Set
    End Property

    Public Property Schuljahr(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As String
        Get
            If Not Isloaded Then LoadConfig(trans)
            d.RowFilter = "param='Schuljahr'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("Schuljahr", Me.Schuljahr_Aktuell.ToString, , trans)
                Return Me.Schuljahr_Aktuell.ToString
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("Schuljahr", CStr(Value), , trans)
        End Set
    End Property

    Public Property V3_Stelle() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='V3_Stelle'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("V3_Stelle", DEF_V3_STELLE)
                Return DEF_V3_STELLE
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("V3_Stelle", CStr(Value))
        End Set
    End Property

    Public Property V3_Institut() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='V3_Institut'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("V3_Institut", DEF_V3_INSTITUT)
                Return DEF_V3_INSTITUT
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("V3_Institut", CStr(Value))
        End Set
    End Property
    Public Property LastGeocoding() As DateTime
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LASTGEOCODING'"
            If d.Count > 0 Then
                Return CDate(d(0)(1))
            Else
                SetProp("LASTGEOCODING", CStr(Date.Now))
                Return Now
            End If
        End Get
        Set(ByVal Value As DateTime)
            SetProp("LASTGEOCODING", CStr(Value))
        End Set
    End Property


    Public Property LetzteDublettenberechnung() As DateTime
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LetzteDublettenberechnung'"
            If d.Count > 0 Then
                Return CDate(d(0)(1))
            Else
                SetProp("LetzteDublettenberechnung", CStr(Date.Now.AddDays(-1)), , True)
                Return Now
            End If
        End Get
        Set(ByVal Value As DateTime)
            SetProp("LetzteDublettenberechnung", CStr(Value), , True)

        End Set
    End Property
    
    Public Property LastScaneingang() As Date
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LastScaneingang" & "_" & UCase(Setting.CurrentUser) & "'"
            If d.Count > 0 Then
                Return CDate(d(0)(1))
            Else
                Return Date.Today

            End If
        End Get
        Set(ByVal Value As Date)
            SetProp("LastScaneingang" & "_" & UCase(Setting.CurrentUser), Value)
        End Set
    End Property

    Public Property MaxSchulalter() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='MaxSchulalter'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp_Old("MaxSchulalter", DEF_MAXSCHULALTER.ToString)
                Return DEF_MAXSCHULALTER
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp_Old("MaxSchulalter", CStr(Value))
        End Set
    End Property


    Public Property Ablage() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Ablage'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp("Ablage", DEF_ABLAGE)
                Return DEF_ABLAGE
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("Ablage", CStr(Value))
        End Set
    End Property


    Public Property InitialFileDir() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='InitialFileDir'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                'SetProp("InitialFileDir", DEF_INITFILEDIR)
                Return ""
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("InitialFileDir", CStr(Value))
        End Set
    End Property

    Public Property URL_Online_Service() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='URL_Online_Service'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                'SetProp("URL_Online_Service", "https://www.ghdbservices.at/wavm")
                Return "http://localhost:17221"

                Return "https://www.ghdbservices.at/wavm"

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("URL_Online_Service", CStr(Value))
        End Set
    End Property

    Public Property GoogleKey() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='GoogleKey'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                'SetProp("GoogleKey", "")
                Return String.Empty
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("GoogleKey", CStr(Value))
        End Set
    End Property
    Public Property AR_Intervall() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Rechercheintervall'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp("Rechercheintervall", DEF_AR_INTERVALL.ToString)
                Return DEF_AR_INTERVALL
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp("Rechercheintervall", CStr(Value))
        End Set
    End Property


    Public Property FIBUKONTO_HA() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='FIBUKONTO_HA'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp("FIBUKONTO_HA", DEF_FIBUKONTO_HA)
                Return DEF_FIBUKONTO_HA
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("FIBUKONTO_HA", CStr(Value))
        End Set
    End Property

    Public Property GH_Bezeichnung_EZ() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='GH_Bezeichnung_EZ'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else

                Return "das Scheckheft"
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("GH_Bezeichnung_EZ", CStr(Value))
        End Set
    End Property


    Public Property GH_Bezeichnung_MZ() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='GH_Bezeichnung_MZ'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else

                Return "die Scheckhefte"
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("GH_Bezeichnung_MZ", CStr(Value))
        End Set
    End Property

    Public Property EA_Bezeichnung_EZ() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='EA_Bezeichnung_EZ'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else

                Return "der Impfbonbogen"
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("EA_Bezeichnung_EZ", CStr(Value))
        End Set
    End Property

    Public Property EA_Bezeichnung_MZ() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='EA_Bezeichnung_MZ'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else

                Return "die Impfbonbögen"
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("EA_Bezeichnung_MZ", CStr(Value))
        End Set
    End Property

    'Public Property GUPrint_Font() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='GUPrint_Font'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            SetProp("GUPrint_Font", GHDruck_GetDefFont)
    '            Return GHDruck_GetDefFont()
    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("GUPrint_Font", CStr(Value))
    '    End Set
    'End Property
    Public Property FIBUKONTO_HONORARE() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='FIBUKONTO_HONORARE'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp("FIBUKONTO_HONORARE", DEF_FIBUKONTO_HONORARE)
                Return DEF_FIBUKONTO_HONORARE
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("FIBUKONTO_HONORARE", CStr(Value))
        End Set
    End Property

    Public Property IMPF_TIMEOUT() As Date
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='IMPF_TIMEOUT'"
            If d.Count > 0 Then
                Return CDate(d(0)(1))
            Else
                SetProp("IMPF_TIMEOUT", DEF_IMPF_TIMEOUT.ToString)
                Return DEF_IMPF_TIMEOUT
            End If
        End Get
        Set(ByVal Value As Date)
            SetProp("IMPF_TIMEOUT", CStr(Value))
        End Set
    End Property
    Public Property Geoanalysen_MinColor() As System.Drawing.Color
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Geoanalysen_MinColor'"
            If d.Count > 0 Then
                Return Me.ColorFromString(d(0)(1).ToString)
            Else
                SetProp("Geoanalysen_MinColor", Me.ColorToString(System.Drawing.Color.White))
                Return System.Drawing.Color.White
            End If
        End Get
        Set(ByVal Value As System.Drawing.Color)
            SetProp("Geoanalysen_MinColor", Me.ColorToString(Value))
        End Set
    End Property


    Public Property Geoanalysen_MaxColor() As System.Drawing.Color
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Geoanalysen_MaxColor'"
            If d.Count > 0 Then
                Return Me.ColorFromString(d(0)(1).ToString)
            Else
                SetProp("Geoanalysen_MaxColor", Me.ColorToString(System.Drawing.Color.Green))
                Return System.Drawing.Color.Green
            End If
        End Get
        Set(ByVal Value As System.Drawing.Color)
            SetProp("Geoanalysen_MaxColor", Me.ColorToString(Value))
        End Set
    End Property

    Public Property Geoanalysen_MapColor() As System.Drawing.Color
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Geoanalysen_MapColor'"
            If d.Count > 0 Then
                Return Me.ColorFromString(d(0)(1).ToString)
            Else
                Dim DEF_BACKG As String = "0|" & DEF_BACKG_R & "|" & DEF_BACKG_G & "|" & DEF_BACKG_B
                SetProp("Geoanalysen_MapColor", DEF_BACKG)
                Return Me.ColorFromString(DEF_BACKG)
            End If
        End Get
        Set(ByVal Value As System.Drawing.Color)
            SetProp("Geoanalysen_MapColor", Me.ColorToString(Value))
        End Set
    End Property
    Public Property Geoanalysen_SymbColor() As System.Drawing.Color
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Geoanalysen_SymbColor'"
            If d.Count > 0 Then
                Return Me.ColorFromString(d(0)(1).ToString)
            Else
                SetProp("Geoanalysen_SymbColor", Me.ColorToString(System.Drawing.Color.Yellow))
                Return System.Drawing.Color.Yellow
            End If
        End Get
        Set(ByVal Value As System.Drawing.Color)
            SetProp("Geoanalysen_SymbColor", Me.ColorToString(Value))
        End Set
    End Property
    Private Function ColorToString(ByVal c As System.Drawing.Color) As String
        Return c.A & "|" & c.R & "|" & c.G & "|" & c.B
    End Function
    Private Function ColorFromString(ByVal strC As String) As System.Drawing.Color
        Dim s() As String = strC.Split(CChar("|"))
        Return System.Drawing.Color.FromArgb(CInt(s(0)), CInt(s(1)), CInt(s(2)), CInt(s(3)))
    End Function


    Public Property IMPF_TIMEOUT_REL() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='IMPF_TIMEOUT_REL'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp("IMPF_TIMEOUT_REL", DEF_IMPF_TIMEOUT_REL.ToString)
                Return DEF_IMPF_TIMEOUT_REL
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp("IMPF_TIMEOUT_REL", CStr(Value))
        End Set
    End Property
    Public Property Georef_TN() As Date
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='GEOREF_TN'"
            If d.Count > 0 Then
                Return CDate(d(0)(1))
            Else
                SetProp("GEOREF_TN", DEF_GEOREF_TN.ToString)
                Return DEF_IMPF_TIMEOUT
            End If
        End Get
        Set(ByVal Value As Date)
            SetProp("GEOREF_TN", CStr(Value))
        End Set
    End Property

    Public Property LOAD_PUPIL() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LOAD_PUPIL'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp_Old("LOAD_PUPIL", DEF_LOAD_PUPIL.ToString)
                Return DEF_LOAD_PUPIL
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp_Old("LOAD_PUPIL", CStr(Value))
        End Set
    End Property

    'Public Property BG_GEOCODING() As Boolean
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='BG_GEOCODING'"
    '        If d.Count > 0 Then
    '            Return IIf(CInt(d(0)(1)) <> 0, True, False)
    '        Else
    '            SetProp("BG_GEOCODING", 0)
    '            Return False
    '        End If
    '    End Get
    '    Set(ByVal Value As Boolean)
    '        SetProp("BG_GEOCODING", IIf(Value, "1", "0"))
    '    End Set
    'End Property
    'Public Property BG_GEOCODING_CITYONLY() As Boolean
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='BG_GEOCODING_CITYONLY'"
    '        If d.Count > 0 Then
    '            Return IIf(CInt(d(0)(1)) <> 0, True, False)
    '        Else
    '            SetProp("BG_GEOCODING_CITYONLY", 0)
    '            Return False
    '        End If
    '    End Get
    '    Set(ByVal Value As Boolean)
    '        SetProp("BG_GEOCODING_CITYONLY", IIf(Value, "1", "0"))
    '    End Set
    'End Property
    Public Property XT_CITY(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Boolean
        Get
            If Not Isloaded Then LoadConfig(trans)
            d.RowFilter = "param='XT_CITY'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) <> 0, True, False))
            Else
                SetProp("XT_CITY", "0", trans)
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("XT_CITY", IIf(Value, "1", "0").ToString, trans)
        End Set
    End Property


    Public Property MD_HIDE_UNBEKANNT() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='MD_HIDE_UNBEKANNT'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) <> 0, True, False))
            Else
                SetProp("MD_HIDE_UNBEKANNT", "1")
                Return True
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("MD_HIDE_UNBEKANNT", IIf(Value, "1", "0").ToString)
        End Set
    End Property
    Public ReadOnly Property DublSuppressGH() As Boolean
        Get
            Return False

            'If Not Isloaded Then LoadConfig()
            'd.RowFilter = "param='DublSuppressGH'"
            'If d.Count > 0 Then
            '    Return CBool(IIf(CInt(d(0)(1)) <> 0, True, False))
            'Else
            '    SetProp("DublSuppressGH", "1")
            '    Return True
            'End If
        End Get
        'Set(ByVal Value As Boolean)
        '    SetProp("DublSuppressGH", IIf(Value, "1", "0").ToString)
        'End Set
    End Property
    Public Property InfoMails_RegionenExcludieren() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='InfoMails_RegionenExcludieren'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) <> 0, True, False))
            Else
                SetProp("InfoMails_RegionenExcludieren", "1")
                Return True
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("InfoMails_RegionenExcludieren", IIf(Value, "1", "0").ToString)
        End Set
    End Property

    'Public Property BG_GEOCODING_MACHINE() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='BG_GEOCODING_MACHINE'"
    '        If d.Count > 0 Then
    '            Return CStr(d(0)(1))
    '        Else
    '            SetProp("BG_GEOCODING_MACHINE", "")
    '            Return ""
    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("BG_GEOCODING_MACHINE", CStr(Value))
    '    End Set
    'End Property
    Public Property MAX_KINDESALTER() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='MAX_KINDESALTER'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp("MAX_KINDESALTER", DEF_MAXKINDESALTER.ToString)
                Return DEF_MAXKINDESALTER
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp("MAX_KINDESALTER", CStr(Value))
        End Set
    End Property

    Public Property MaxFax() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='MaxRecherchen'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp("MaxRecherchen", DEF_MAXFAX.ToString)
                Return DEF_MAXFAX
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp("MaxRecherchen", CStr(Value))
        End Set
    End Property

    Public Property Mailserver_Version() As ExchangeVersion
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Mailserver_Version'"
            If d.Count > 0 Then
                Return CType(d(0)(1), Einstellungen.ExchangeVersion)
            Else
                SetProp("Mailserver_Version", CStr(ExchangeVersion.Exchange_2000))
                Return ExchangeVersion.Exchange_2000
            End If
        End Get
        Set(ByVal Value As ExchangeVersion)
            SetProp("Mailserver_Version", CStr(Value))
        End Set
    End Property

    Public Property MISS_DBL_TIMEOUT() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='MISS_DBL_TIMEOUT'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp_Old("MISS_DBL_TIMEOUT", DEF_MISS_DBL_TIMEOUT.ToString)
                Return DEF_MISS_DBL_TIMEOUT
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp_Old("MISS_DBL_TIMEOUT", CStr(Value))
        End Set
    End Property


    Public Property SMTP_SERVER() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='SMTP_SERVER'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp("SMTP_SERVER", Me.DomainController)
                Return Me.DomainController
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("SMTP_SERVER", CStr(Value))
        End Set
    End Property

    Public Property SMTP_SERVER_ALTERNATIV() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='SMTP_SERVER_ALTERNATIV'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp("SMTP_SERVER_ALTERNATIV", "192.168.2.254")
                Return "192.168.2.254"

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("SMTP_SERVER_ALTERNATIV", CStr(Value))
        End Set
    End Property

    Public ReadOnly Property Schuljahr_Aktuell() As Integer
        Get

            If Today.Month <= 8 Then
                Return Today.Year - 1
            Else
                Return Today.Year
            End If

        End Get
    End Property
    Public Property URL_OLAP_WEBSERVICE() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='URL_OLAP_WEBSERVICE'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp("URL_OLAP_WEBSERVICE", DEF_OLAP_WS.Replace("localhost", Me.DomainController))
                Return DEF_OLAP_WS.Replace("localhost", Me.DomainController)
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("URL_OLAP_WEBSERVICE", CStr(Value))
        End Set
    End Property

    Public Property URL_GEOCODER() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='URL_GEOCODER'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp("URL_GEOCODER", DEF_GEOCODER.Replace("localhost", Me.DomainController))
                Return DEF_GEOCODER.Replace("localhost", Me.DomainController)
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("URL_GEOCODER", CStr(Value))
        End Set
    End Property

    'Public Property MR4_Pfad() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MR4_PFAD" & "_" & UCase(Setting.CurrentUser) & "'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            SetProp("MR4_PFAD" & "_" & UCase(Setting.CurrentUser), System.Windows.Forms.Application.ExecutablePath)
    '            Return System.Windows.Forms.Application.ExecutablePath
    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MR4_PFAD" & "_" & UCase(Setting.CurrentUser), CStr(Value))
    '    End Set
    'End Property
    'Public Property MR5_Pfad() As String
    '    Get
    '        Try
    '            System.Threading.Monitor.Enter(Me)

    '            If Not Isloaded Then LoadConfig()
    '            d.RowFilter = "param='MR5_Pfad" & "_" & UCase(Setting.CurrentUser) & "'"
    '            If d.Count > 0 Then
    '                Return d(0)(1).ToString
    '            Else
    '                SetProp("MR5_Pfad" & "_" & UCase(Setting.CurrentUser), System.Windows.Forms.Application.ExecutablePath)
    '                Return System.Windows.Forms.Application.ExecutablePath
    '            End If
    '        Catch ex As Exception
    '            Throw New Exception(ex.Message)
    '        Finally
    '            System.Threading.Monitor.Exit(Me)

    '        End Try

    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MR5_Pfad" & "_" & UCase(Setting.CurrentUser), CStr(Value))
    '    End Set
    'End Property


    'Public Property MM_SA_Export_Intervall() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_SA_Export_Intervall'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            SetProp("MM_SA_Export_Intervall", "q")
    '            Return "q"
    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_SA_Export_Intervall", CStr(Value))
    '    End Set
    'End Property

    'Public Property MM_Eval_Export_Intervall() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_Eval_Export_Intervall'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            SetProp("MM_Eval_Export_Intervall", "m")
    '            Return "m"
    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_Eval_Export_Intervall", CStr(Value))
    '    End Set
    'End Property

    'Public Property MM_Bura_Export_Intervall() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_Bura_Export_Intervall'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            SetProp("MM_Bura_Export_Intervall", "m")
    '            Return "m"
    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_Bura_Export_Intervall", CStr(Value))
    '    End Set
    'End Property

    'Public Property MM_MAILTO_SA() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_MAILTO_SA'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            Return ""
    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_MAILTO_SA", CStr(Value))
    '    End Set
    'End Property
    'Public Property MM_MAILTO_BURA(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As String
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='MM_MAILTO_BURA'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            Return ""
    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_MAILTO_BURA", CStr(Value), trans)
    '    End Set
    'End Property

    'Public Property MM_MAILTO_EVAL() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_MAILTO_EVAL'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            Return ""
    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_MAILTO_EVAL", CStr(Value))
    '    End Set
    'End Property

    'Public Property MM_EMAIL_INBOX() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_EMAIL_INBOX'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            Return ""
    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_EMAIL_INBOX", CStr(Value))
    '    End Set
    'End Property

    'Public Property MM_STARTTIME() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_STARTTIME'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            Return DEF_MM_START
    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_STARTTIME", CStr(Value))
    '    End Set
    'End Property

    'Public Property MM_ServiceRuns() As Integer
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_ServiceRuns'"
    '        If d.Count > 0 Then
    '            Return CInt(d(0)(1))
    '        Else
    '            Return 1
    '        End If
    '    End Get
    '    Set(ByVal Value As Integer)
    '        SetProp("MM_ServiceRuns", CStr(Value))
    '    End Set
    'End Property

    Public Property PlayBinPause() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='PlayBinPause'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp("PlayBinPause", "200")
                Return 200
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp("PlayBinPause", CStr(Value))
        End Set
    End Property


    'Public Property MM_Date_SA_Export(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Date
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='MM_Date_SA_Export'"
    '        If d.Count > 0 Then
    '            Return CDate(d(0)(1))
    '        Else
    '            SetProp("MM_Date_SA_Export", DAT_NULL.ToString, trans)
    '            Return DAT_NULL
    '        End If
    '    End Get
    '    Set(ByVal Value As Date)
    '        SetProp("MM_Date_SA_Export", CStr(Value), trans)
    '    End Set
    'End Property
    'Public Property MM_Date_SA_Export_Last(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Date
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='MM_Date_SA_Export_Last'"
    '        If d.Count > 0 Then
    '            Return CDate(d(0)(1))
    '        Else
    '            SetProp("MM_Date_SA_Export_Last", DAT_NULL.ToString, trans)
    '            Return DAT_NULL
    '        End If
    '    End Get
    '    Set(ByVal Value As Date)
    '        SetProp("MM_Date_SA_Export_Last", CStr(Value), trans)
    '    End Set
    'End Property
    'Public Property MM_Date_SA_Export_Last2(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Date
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='MM_Date_SA_Export_Last2'"
    '        If d.Count > 0 Then
    '            Return CDate(d(0)(1))
    '        Else
    '            SetProp("MM_Date_SA_Export_Last2", DAT_NULL.ToString, trans)
    '            Return DAT_NULL
    '        End If
    '    End Get
    '    Set(ByVal Value As Date)
    '        SetProp("MM_Date_SA_Export_Last2", CStr(Value), trans)
    '    End Set
    'End Property

    'Public Property MM_Date_Eval_Export(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Date
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='MM_Date_Eval_Export'"
    '        If d.Count > 0 Then
    '            Return CDate(d(0)(1))
    '        Else
    '            SetProp("MM_Date_Eval_Export", DAT_NULL.ToString, trans)
    '            Return DAT_NULL
    '        End If
    '    End Get
    '    Set(ByVal Value As Date)
    '        SetProp("MM_Date_Eval_Export", CStr(Value), trans)
    '    End Set
    'End Property
    'Public Property MM_Date_Bura_Export(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Date
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='MM_Date_Bura_Export'"
    '        If d.Count > 0 Then
    '            Return CDate(d(0)(1))
    '        Else
    '            SetProp("MM_Date_Bura_Export", DAT_NULL.ToString, trans)
    '            Return DAT_NULL
    '        End If
    '    End Get
    '    Set(ByVal Value As Date)
    '        SetProp("MM_Date_Bura_Export", CStr(Value), trans)
    '    End Set
    'End Property

    'Public Property MM_SendungsID_SA(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Integer
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='MM_SendungsID_SA'"
    '        If d.Count > 0 Then
    '            Return CInt(d(0)(1))
    '        Else
    '            SetProp("MM_SendungsID_SA", "0", trans)
    '            Return 0
    '        End If
    '    End Get
    '    Set(ByVal Value As Integer)
    '        SetProp("MM_SendungsID_SA", CStr(Value), trans)
    '    End Set
    'End Property
    'Public Property MM_SendungsID_Eval(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Integer
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='MM_SendungsID_SA'"
    '        If d.Count > 0 Then
    '            Return CInt(d(0)(1))
    '        Else
    '            SetProp("MM_SendungsID_SA", "0", trans)
    '            Return 0
    '        End If
    '    End Get
    '    Set(ByVal Value As Integer)
    '        SetProp("MM_SendungsID_SA", CStr(Value), trans)
    '    End Set
    'End Property

    'Public Property MM_Last_Account() As Date
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_Last_Account'"
    '        If d.Count > 0 Then
    '            Return CDate(d(0)(1))
    '        Else
    '            SetProp("MM_Last_Account", DAT_NULL.ToString)
    '            Return DAT_NULL
    '        End If
    '    End Get
    '    Set(ByVal Value As Date)
    '        SetProp("MM_Last_Account", CStr(Value))
    '    End Set
    'End Property

    'Public Property MM_MinAlter() As Integer
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_MinAlter'"
    '        If d.Count > 0 Then
    '            Return CInt(d(0)(1))
    '        Else
    '            SetProp("MM_MinAlter", MM_MIN_ALTER.ToString)
    '            Return MM_MIN_ALTER
    '        End If
    '    End Get
    '    Set(ByVal Value As Integer)
    '        SetProp("MM_MinAlter", CStr(Value))
    '    End Set
    'End Property

    'Public Property MM_MaxAlter() As Integer
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_MaxAlter'"
    '        If d.Count > 0 Then
    '            Return CInt(d(0)(1))
    '        Else
    '            SetProp("MM_MaxAlter", MM_MAX_ALTER.ToString)
    '            Return MM_MAX_ALTER
    '        End If
    '    End Get
    '    Set(ByVal Value As Integer)
    '        SetProp("MM_MaxAlter", CStr(Value))
    '    End Set
    'End Property


    Private Property Nr_MKPMail(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As String
        Get
            If Not Isloaded Then LoadConfig(trans)
            d.RowFilter = "param='NR_MAIL'"
            If d.Count > 0 Then
                Return CStr(d(0)(1))
            Else
                SetProp("NR_MAIL", "0", trans)
                Return "0"
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("NR_MAIL", CStr(Value), trans)
        End Set
    End Property
    'Public Property Nr_Mail(ByVal Prog As Integer, Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As String
    '    Get

    '        Select Case Prog
    '            Case GHParam.Programme.MKP, GHParam.Programme.Impfnetzwerk
    '                Return Me.Nr_MKPMail(trans)
    '            Case GHParam.Programme.Mammographiescreening
    '                Return 0 'Me.Nr_MammoMail(trans)
    '            Case Else
    '                Throw New Exception("Für Programm " & Prog & " kann keine Mail-Nr. generiert werden.")
    '        End Select

    '    End Get
    '    Set(ByVal Value As String)
    '        Select Case Prog
    '            Case GHParam.Programme.MKP, GHParam.Programme.Impfnetzwerk
    '                Me.Nr_MKPMail(trans) = Value
    '            Case GHParam.Programme.Mammographiescreening
    '                'Me.Nr_MammoMail(trans) = Value
    '            Case Else
    '                Throw New Exception("Für Programm " & Prog & " kann keine Mail-Nr. generiert werden.")
    '        End Select

    '    End Set
    'End Property
    'Private Property Nr_MammoMail(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As String
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='NR_MAMMOMAIL'"
    '        If d.Count > 0 Then
    '            Return CStr(d(0)(1))
    '        Else
    '            SetProp("NR_MAMMOMAIL", "0", trans)
    '            Return "0"
    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("NR_MAMMOMAIL", CStr(Value), trans)
    '    End Set
    'End Property
    Public Property WebDAVLogon() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='WebDAVLogon'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp("WebDAVLogon", ".")
                Return "."
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("WebDAVLogon", CStr(Value))
        End Set
    End Property

    'Public Property MM_LocalMail() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_LocalMail'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            SetProp("MM_LocalMail", ".")
    '            Return ""
    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_LocalMail", CStr(Value))
    '    End Set
    'End Property

    'Public Property MM_FromMail() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_FromMail'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            SetProp("MM_FromMail", ".")
    '            Return ""
    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_FromMail", CStr(Value))
    '    End Set
    'End Property

    Public Property FaxMethod() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='FaxMethod'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp("FaxMethod", "0")
                Return 0
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp("FaxMethod", CStr(Value))
        End Set
    End Property
    Public Property EmailToFaxDomain() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='EmailToFaxDomain'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp("EmailToFaxDomain", "")
                Return ""
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("EmailToFaxDomain", CStr(Value))
        End Set
    End Property

    Public Property FaxDeckblatt() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='FaxDeckblatt'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp("FaxDeckblatt", "standard")
                Return "standard"
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("FaxDeckblatt", CStr(Value))
        End Set
    End Property
    Public Property Logging() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Logging'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) <> 0, True, False))
            Else
                SetProp("Logging", "0")
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("Logging", IIf(Value, "1", "0").ToString)
        End Set
    End Property
    Public Property GeneralLogging() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='GeneralLogging'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) <> 0, True, False))
            Else
                SetProp("GeneralLogging", "0")
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("GeneralLogging", IIf(Value, "1", "0").ToString)
        End Set
    End Property
    Public Property WriteBackAddress() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='WriteBackAddress'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) <> 0, True, False))
            Else
                SetProp("WriteBackAddress", "1")
                Return True
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("WriteBackAddress", IIf(Value, "1", "0").ToString)
        End Set
    End Property
    Public ReadOnly Property DomainController() As String
        Get
            Dim dnsName As String = String.Empty
            Try
                'Dim rootDse As New System.DirectoryServices.DirectoryEntry("LDAP://RootDSE")
                ''Dim defaultContext As String = DirectCast(rootDse.Properties("defaultNamingContext").Value, String)
                'dnsName = DirectCast(rootDse.Properties("dnsHostName").Value, String)
                'rootDse.Dispose()




            Catch ex As Exception

            End Try
            If String.IsNullOrEmpty(dnsName) Then
                Return String.Empty
            Else

                Return dnsName.Split(CChar("."))(0)
            End If
        End Get

    End Property
    Public Property RepositoriumMailEntwürfe() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='RepositoriumMailEntwürfe'"
            If d.Count > 0 Then
                If Not IsDBNull(d(0)(1)) Then
                    Return d(0)(1).ToString
                Else
                    SetProp("RepositoriumMailEntwürfe", OL_PRIV_ENTWUERFE)
                    Return OL_PRIV_ENTWUERFE
                End If
            Else
                SetProp("RepositoriumMailEntwürfe", OL_PRIV_ENTWUERFE)
                Return OL_PRIV_ENTWUERFE
            End If
        End Get
        Set(ByVal Value As String)
            If String.IsNullOrEmpty(Value) Then Return
            SetProp("RepositoriumMailEntwürfe", CStr(Value))
        End Set
    End Property
    Public Property Openlist() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='OPENLIST" & "_" & UCase(Setting.CurrentUser) & "'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) = 0, False, True))
            Else
                SetProp("OPENLIST" & "_" & UCase(Setting.CurrentUser), "1")
                Return True
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("OPENLIST" & "_" & UCase(Setting.CurrentUser), CStr(IIf(Value, 1, 0)))
        End Set
    End Property

    Public Property TestAddress() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='TestAddress" & "_" & UCase(Setting.CurrentUser) & "'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) = 0, False, True))
            Else
                SetProp("TestAddress" & "_" & UCase(Setting.CurrentUser), "0")
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("TestAddress" & "_" & UCase(Setting.CurrentUser), CStr(IIf(Value, 1, 0)))
        End Set
    End Property

    Public Property Opt_NN() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='OPT_NN" & "_" & UCase(Setting.CurrentUser) & "'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) = 0, False, True))
            Else
                SetProp("OPT_NN" & "_" & UCase(Setting.CurrentUser), "1")
                Return True
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("OPT_NN" & "_" & UCase(Setting.CurrentUser), CStr(IIf(Value, 1, 0)))
        End Set
    End Property

    Public Property Opt_VN() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='OPT_VN" & "_" & UCase(Setting.CurrentUser) & "'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) = 0, False, True))
            Else
                SetProp("OPT_VN" & "_" & UCase(Setting.CurrentUser), "1")
                Return True
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("OPT_VN" & "_" & UCase(Setting.CurrentUser), CStr(IIf(Value, 1, 0)))
        End Set
    End Property

    'Public Property MM_SA_Pool1_Ausständig(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Boolean
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='MM_SA_Pool1_Ausständig'"
    '        If d.Count > 0 Then
    '            Return CBool(IIf(CInt(d(0)(1)) = 0, False, True))
    '        Else
    '            SetProp("MM_SA_Pool1_Ausständig", "0", trans)
    '            Return False
    '        End If
    '    End Get
    '    Set(ByVal Value As Boolean)
    '        SetProp("MM_SA_Pool1_Ausständig", CStr(IIf(Value, 1, 0)), trans)
    '    End Set
    'End Property
    'Public Property MM_SA_Pool2_Ausständig(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Boolean
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='MM_SA_Pool2_Ausständig'"
    '        If d.Count > 0 Then
    '            Return CBool(IIf(CInt(d(0)(1)) = 0, False, True))
    '        Else
    '            SetProp("MM_SA_Pool2_Ausständig", "0", trans)
    '            Return False
    '        End If
    '    End Get
    '    Set(ByVal Value As Boolean)
    '        SetProp("MM_SA_Pool2_Ausständig", CStr(IIf(Value, 1, 0)), trans)
    '    End Set
    'End Property


    Public Property Opt_Addr() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='OPT_ADDR" & "_" & UCase(Setting.CurrentUser) & "'"

            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) = 0, False, True))
            Else
                SetProp("OPT_ADDR" & "_" & UCase(Setting.CurrentUser), "1")
                Return True
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("OPT_ADDR" & "_" & UCase(Setting.CurrentUser), CStr(IIf(Value, 1, 0)))
        End Set
    End Property

    Public Property Opt_Arzt() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='OPT_ARZT" & "_" & UCase(Setting.CurrentUser) & "'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) = 0, False, True))
            Else
                SetProp("OPT_ARZT" & "_" & UCase(Setting.CurrentUser), "1")
                Return True
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("OPT_ARZT" & "_" & UCase(Setting.CurrentUser), CStr(IIf(Value, 1, 0)))
        End Set
    End Property
    Public Property Intervall_fehlende_Datenblätter() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='MISS_DBL_TIMEOUT'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp("MISS_DBL_TIMEOUT", "0")
                Return 0
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp("MISS_DBL_TIMEOUT", CStr(Value))
        End Set
    End Property

    Public Property NBQuiet() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='NBQuiet'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp_Old("NBQuiet", "0")
                Return 0
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp_Old("NBQuiet", CStr(Value))
        End Set
    End Property


    'Public Property LockImpfdoku() As Date
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='LockImpfdoku'"
    '        If d.Count > 0 Then
    '            Return CDate(d(0)(1))
    '        Else
    '            SetProp_Old("LockImpfdoku", DAT_NULL.ToString)
    '            Return DAT_NULL
    '        End If
    '    End Get
    '    Set(ByVal Value As Date)
    '        SetProp_Old("LockImpfdoku", CStr(Value))
    '    End Set
    'End Property

    Public Property LetzteAufräumarbeiten() As Date
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LetzteAufräumarbeiten'"
            If d.Count > 0 Then
                Return CDate(d(0)(1))
            Else
                Return Date.Today.AddDays(-1)
            End If
        End Get
        Set(ByVal Value As Date)
            SetProp("LetzteAufräumarbeiten", CStr(Value), , True)
        End Set
    End Property

    Public Property LetzterMRBericht() As Date
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LetzterMRBericht'"
            If d.Count > 0 Then
                Return CDate(d(0)(1))
            Else
                Return Date.Today.AddYears(-1)
            End If
        End Get
        Set(ByVal Value As Date)
            SetProp("LetzterMRBericht", CStr(Value))
        End Set
    End Property


#If Not NET_2_0 Then

    'Public Property ZeitplanMRBericht() As ZeitpLanMR5Bericht
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='ZeitplanMRBericht'"
    '        If d.Count > 0 Then
    '            Return CType(d(0)(1), ZeitpLanMR5Bericht)
    '        Else
    '            Return ZeitpLanMR5Bericht.KeinBericht
    '        End If
    '    End Get
    '    Set(ByVal Value As ZeitpLanMR5Bericht)
    '        SetProp("ZeitplanMRBericht", CInt(Value))
    '    End Set
    'End Property
#End If
    Public Property FOLDER_SCHOOLDATA() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='FOLDER_SCHOOLDATA'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp("FOLDER_SCHOOLDATA", DEF_FOLDER_SCHOOLDATA)
                Return DEF_FOLDER_SCHOOLDATA
            End If
        End Get
        Set(ByVal Value As String)
            SetProp("FOLDER_SCHOOLDATA", CStr(Value))
        End Set
    End Property

    Public Property SCHOOLFOLDERASK() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='SCHOOLFOLDERASK'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) <> 0, True, False))
            Else
                SetProp("SCHOOLFOLDERASK", "0")
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("SCHOOLFOLDERASK", IIf(Value, "1", "0").ToString)
        End Set
    End Property

    Public Property AskTI() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='AskTI'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) <> 0, True, False))
            Else
                SetProp_Old("AskTI", "0")
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp_Old("AskTI", IIf(Value, "1", "0").ToString)
        End Set
    End Property
    'Public Property MM_Pool1Check() As Boolean
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_Pool1Check'"
    '        If d.Count > 0 Then
    '            Return CBool(IIf(CInt(d(0)(1)) <> 0, True, False))
    '        Else
    '            SetProp("MM_Pool1Check", "0")
    '            Return False
    '        End If
    '    End Get
    '    Set(ByVal Value As Boolean)
    '        SetProp("MM_Pool1Check", IIf(Value, "1", "0").ToString)
    '    End Set
    'End Property

    Public Property PGPUserID() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='PGPUserID'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("PGPUserID", "")
                Return ""
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("PGPUserID", CStr(Value))
        End Set
    End Property
    Public Property Hontext() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Hontext'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("Hontext", "")
                Return ""
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("Hontext", CStr(Value))
        End Set
    End Property
    Public Property Hontext1() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Hontext1'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("Hontext1", "")
                Return ""
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("Hontext1", CStr(Value))
        End Set
    End Property
    Public Property Hontext2() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Hontext2'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("Hontext2", "")
                Return ""
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("Hontext2", CStr(Value))
        End Set
    End Property
    Public Property KritDur() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='KritDur'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp_Old("KritDur", "0")
                Return 0
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp_Old("KritDur", CStr(Value))
        End Set
    End Property
    'Public Property MM_Pool1Intermix() As Integer
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_Pool1Intermix'"
    '        If d.Count > 0 Then
    '            Return CInt(d(0)(1))
    '        Else
    '            SetProp("MM_Pool1Intermix", "0")
    '            Return 0
    '        End If
    '    End Get
    '    Set(ByVal Value As Integer)
    '        SetProp("MM_Pool1Intermix", CStr(Value))
    '    End Set
    'End Property
    Public Property GUPrint_Offset_X() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='GUPrint_Offset_X'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp("GUPrint_Offset_X", "0")
                Return 0
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp("GUPrint_Offset_X", CStr(Value))
        End Set
    End Property
    Public Property GUPrint_Offset_Y() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='GUPrint_Offset_Y'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp("GUPrint_Offset_Y", "0")
                Return 0
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp("GUPrint_Offset_Y", CStr(Value))
        End Set
    End Property
    Public Property Dokerf() As Single
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Dokerf'"
            If d.Count > 0 Then
                Return CSng(d(0)(1))
            Else
                SetProp_Old("Dokerf", "0")
                Return 0
            End If
        End Get
        Set(ByVal Value As Single)
            SetProp_Old("Dokerf", CStr(Value), True)
        End Set
    End Property

    Public Property TxtBefList() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='TxtBefList'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("TxtBefList", "")
                Return ""
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("TxtBefList", CStr(Value))
        End Set
    End Property
    Public Property TxtAftList() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='TxtAftList'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                SetProp_Old("TxtAftList", "")
                Return ""
            End If
        End Get
        Set(ByVal Value As String)
            SetProp_Old("TxtAftList", CStr(Value))
        End Set
    End Property
    Public Property BehaveImpfdokumissing() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='impfdokumissing'"
            If d.Count > 0 Then
                Return CInt(d(0)(1))
            Else
                SetProp_Old("impfdokumissing", "0")
                Return 0
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp_Old("impfdokumissing", CStr(Value))
        End Set
    End Property



    'Public ReadOnly Property ErfasserID() As Integer
    '    Get
    '        Select Case ID_VERSION
    '            Case ID_WISSAK
    '                ErfasserID = 2
    '            Case Else
    '                Err.Raise(65001, "", "Erfasser-ID für die Stelle nicht gefunden.")
    '                Return NO_VAL
    '        End Select
    '    End Get
    'End Property

    'Public Property MM_MaxDelay_Date(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Integer
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='MM_MaxDelay_Date'"
    '        If d.Count > 0 Then
    '            Return CInt(d(0)(1))
    '        Else
    '            SetProp("MM_MaxDelay_Date", "3", trans)
    '            Return 3
    '        End If
    '    End Get
    '    Set(ByVal Value As Integer)
    '        SetProp("MM_MaxDelay_Date", CStr(Value), trans)
    '    End Set
    'End Property


    Public Property Ladeschüler() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='load_pupil'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) <> 0, True, False))
            Else
                SetProp("load_pupil", "0")
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("load_pupil", IIf(Value, "1", "0").ToString)
        End Set
    End Property
    ''' <summary>
    ''' Bundesland der Stelle aus Tab "Orte"
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Bundesland() As String
        Get
            Dim rs As DataTable
            Try
                rs = db_con.GetRecordSet("select bl from orte where postleitzahl=" & Me.PLZ)

                If rs.Rows.Count > 0 Then
                    Return rs.Rows(0)("bl").ToString
                Else
                    Return ""
                End If
            Catch ex As Exception
                Return ""
            End Try
        End Get

    End Property

    Public ReadOnly Property NUTS2Stelle(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As String
        Get
            Dim rs As DataTable
            Try
                rs = db_con.GetRecordSet("select distinct nuts2 from nuts where plz=" & Me.PLZ, trans)

                If rs.Rows.Count > 0 Then
                    Return rs.Rows(0)(0).ToString
                Else
                    Return ""
                End If
            Catch ex As Exception
                Return ""
            End Try
        End Get

    End Property

    ' ''' <summary>
    ' ''' Letzter Dateidownload für Mammographie (Default: 1.1.1900)
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Property MM_LastTransfer() As Date
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_LastTransfer'"
    '        If d.Count > 0 Then
    '            Return CDate(d(0)(1))
    '        Else
    '            SetProp("MM_LastTransfer", DAT_NULL.ToString)
    '            Return DAT_NULL
    '        End If
    '    End Get
    '    Set(ByVal Value As Date)
    '        SetProp("MM_LastTransfer", CStr(Value))
    '    End Set
    'End Property


    'Public Property SI_LastTransfer() As Date
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='SI_LastTransfer'"
    '        If d.Count > 0 Then
    '            Return CDate(d(0)(1))
    '        Else
    '            SetProp("SI_LastTransfer", DAT_NULL.ToString)
    '            Return DAT_NULL
    '        End If
    '    End Get
    '    Set(ByVal Value As Date)
    '        SetProp("SI_LastTransfer", CStr(Value))
    '    End Set
    'End Property

    'Public Property MM_SigningKeyPP() As String
    '    Get

    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_SigningKeyPP'"
    '        If d.Count > 0 Then
    '            Dim c As New SimpleCrypto
    '            Dim pw As String = System.Text.Encoding.Default.GetString(System.Convert.FromBase64String(INFO1))
    '            Return c.DecryptString(d(0)(1).ToString, pw)
    '        Else
    '            Return ""
    '        End If


    '    End Get
    '    Set(ByVal value As String)
    '        Dim c As New SimpleCrypto
    '        Dim pw As String = System.Text.Encoding.Default.GetString(System.Convert.FromBase64String(INFO1))
    '        SetProp("MM_SigningKeyPP", c.EncryptString(value, pw))

    '    End Set

    'End Property

    'Public Property MM_SigningKey() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_SigningKey'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            Return ""

    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_SigningKey", CStr(Value))
    '    End Set
    'End Property
    'Public Property MM_HREF_Inbox() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_HREF_Inbox'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            Return ""

    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_HREF_Inbox", CStr(Value))
    '    End Set
    'End Property
    'Public Property MM_FileDir() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_FileDir'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            Return ""

    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_FileDir", CStr(Value))
    '    End Set
    'End Property

    'Public Property MM_HREF_BCC() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_HREF_BCC'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            Return ""

    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_HREF_BCC", CStr(Value))
    '    End Set
    'End Property
    'Public Property MM_PGPKey_SA(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As String
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='MM_PGPKey_SA'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            Return ""

    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_PGPKey_SA", CStr(Value), trans)
    '    End Set
    'End Property
    'Public Property MM_PGPKey_EVAL(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As String
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='MM_PGPKey_EVAL'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            Return ""

    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_PGPKey_EVAL", CStr(Value), trans)
    '    End Set
    'End Property
    'Public Property MM_PGPKey_BURA() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_PGPKey_BURA'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            Return ""

    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_PGPKey_BURA", CStr(Value))
    '    End Set
    'End Property

    'Public Property MM_BCC() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='MM_BCC'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            Return ""

    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("MM_BCC", CStr(Value))
    '    End Set
    'End Property
    Public Property STD_PRT_ETIKETT() As String
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='STD_PRT_ETIKETT" & "_" & UCase(Environment.MachineName) & "_" & Environment.UserName & "'"
                If d.Count > 0 Then
                    Return d(0)(1).ToString
                Else
                    Return NO_STD_PRINTER
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As String)
            SetProp("STD_PRT_ETIKETT" & "_" & UCase(Environment.MachineName) & "_" & Environment.UserName, CStr(Value))
        End Set
    End Property


    Public Property Zoom_Abschnitt() As Integer
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Zoom_Abschnitt" & "_" & Environment.UserName & "'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 100
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Integer)
            SetProp("Zoom_Abschnitt" & "_" & Environment.UserName, CStr(Value))
        End Set
    End Property

    Public Property Explorer_Eingaben_MaxRec() As Integer
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Explorer_Eingaben_MaxRec" & "_" & Environment.UserName & "'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 10000
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Integer)
            SetProp("Explorer_Eingaben_MaxRec" & "_" & Environment.UserName, CStr(Value))
        End Set
    End Property
    Public Property Explorer_Scans_MaxRec() As Integer
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Explorer_Scans_MaxRec" & "_" & Environment.UserName & "'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 1000
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Integer)
            SetProp("Explorer_Scans_MaxRec" & "_" & Environment.UserName, CStr(Value))
        End Set
    End Property
    Public Property Explorer_Dokumente_MaxRec() As Integer
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Explorer_Dokumente_MaxRec" & "_" & Environment.UserName & "'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 1000
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Integer)
            SetProp("Explorer_Dokumente_MaxRec" & "_" & Environment.UserName, CStr(Value))
        End Set
    End Property
    Public Property Explorer_Eingaben_AutoUpdate() As Boolean
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Explorer_Eingaben_AutoUpdate" & "_" & Environment.UserName & "'"
                If d.Count > 0 Then
                    Return IIf(CInt(d(0)(1)) <> 0, True, False)
                Else
                    Return False
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Boolean)
            SetProp("Explorer_Eingaben_AutoUpdate" & "_" & Environment.UserName, IIf(Value, "1", "0").ToString)
        End Set
    End Property
    Public Property Explorer_Scans_AutoUpdate() As Boolean
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Explorer_Scans_AutoUpdate" & "_" & Environment.UserName & "'"
                If d.Count > 0 Then
                    Return IIf(CInt(d(0)(1)) <> 0, True, False)
                Else
                    Return True
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Boolean)
            SetProp("Explorer_Scans_AutoUpdate" & "_" & Environment.UserName, IIf(Value, "1", "0").ToString)
        End Set
    End Property
    Public Property Explorer_Dokumente_AutoUpdate() As Boolean
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Explorer_Dokumente_AutoUpdate" & "_" & Environment.UserName & "'"
                If d.Count > 0 Then
                    Return IIf(CInt(d(0)(1)) <> 0, True, False)
                Else
                    Return False
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Boolean)
            SetProp("Explorer_Dokumente_AutoUpdate" & "_" & Environment.UserName, IIf(Value, "1", "0").ToString)
        End Set
    End Property

    Public Property Zoom_Dbl() As Integer
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Zoom_Dbl" & "_" & Environment.UserName & "'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 100
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Integer)
            SetProp("Zoom_Dbl" & "_" & Environment.UserName, CStr(Value))
        End Set
    End Property
    Public Property Zoom_Amtsbon() As Integer
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Zoom_Amtsbon" & "_" & Environment.UserName & "'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 100
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Integer)
            SetProp("Zoom_Amtsbon" & "_" & Environment.UserName, CStr(Value))
        End Set
    End Property
    Public Property Zoom_Sonstiges() As Integer
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Zoom_Sonstiges" & "_" & Environment.UserName & "'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 100
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Integer)
            SetProp("Zoom_Sonstiges" & "_" & Environment.UserName, CStr(Value))
        End Set
    End Property
    Public ReadOnly Property RecherchenLaufNr(Optional trans As SqlClient.SqlTransaction = Nothing) As String
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                Dim ret As Integer = 0
                Dim RLNR As String = ""
                Dim LNROrig As String = ""

                Do Until ret = 1
                    Dim tb As DataTable = db_con.GetRecordSet("select val from config where param='RecherchenLaufNr'", trans)
                    If tb.Rows.Count = 0 Then
                        SetProp("RecherchenLaufNr", GetRechercheNr(10), trans)
                        Return GetRechercheNr(10)

                    Else
                        LNROrig = tb.Rows(0)("val")
                        Dim LNR As String = CInt(tb.Rows(0)("val").ToString.Substring(0, 2)) '  + 1).ToString & tb.Rows(0)("val").ToString.Substring(2)
                        If tb.Rows(0)("val").ToString.Substring(2) = Date.Today.Day.ToString("00") & Date.Today.Month.ToString("00") & CInt(Date.Today.Year.ToString.Substring(2, 2)).ToString("00") Then
                            RLNR = GetRechercheNr(LNR + 1)
                        Else
                            RLNR = GetRechercheNr(10)
                        End If
                    End If

                    ret = db_con.FireSQL("update config set val='" & RLNR & "' where param='RecherchenLaufNr' and val='" & LNROrig & "'", trans)

                Loop




                Return RLNR





            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get


    End Property

    Private Function GetRechercheNr(Laufnr As Integer) As String
        Return Laufnr.ToString("00") & Date.Today.Day.ToString("00") & Date.Today.Month.ToString("00") & CInt(Date.Today.Year.ToString.Substring(2, 2)).ToString("00")


    End Function


    Public Property Stelle_Arztnr() As Integer
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='Stelle_Arztnr'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 9000091
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Integer)
            SetProp("Stelle_Arztnr", CStr(Value))
        End Set
    End Property

    Public Property RecherchenWiedervorlageIntervall_FaxEmail() As Integer
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='RecherchenWiedervorlageIntervall_FaxEmail'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 7
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Integer)
            SetProp("RecherchenWiedervorlageIntervall_FaxEmail", CStr(Value))
        End Set
    End Property


    Public Property RecherchenWiedervorlageIntervall_DBErsatz() As Integer
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='RecherchenWiedervorlageIntervall_DBErsatz'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 14
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Integer)
            SetProp("RecherchenWiedervorlageIntervall_DBErsatz", CStr(Value))
        End Set
    End Property

    Public Property RecherchenWiedervorlageIntervall_Brief() As Integer
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='RecherchenWiedervorlageIntervall_Brief'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 14
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Integer)
            SetProp("RecherchenWiedervorlageIntervall_Brief", CStr(Value))
        End Set
    End Property


    Public Property RecherchenAnzahlWiedervorlagen() As Integer
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='RecherchenAnzahlWiedervorlagen'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 1
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Integer)
            SetProp("RecherchenAnzahlWiedervorlagen", CStr(Value))
        End Set
    End Property
    'Public Property MM_EP1R_Deadline(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Date
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='MM_EP1R_Deadline'"
    '        If d.Count > 0 Then
    '            Return CDate(d(0)(1))
    '        Else
    '            SetProp("MM_EP1R_Deadline", DAT_NULL.ToString, trans)
    '            Return DAT_NULL
    '        End If
    '    End Get
    '    Set(ByVal Value As Date)
    '        SetProp("MM_EP1R_Deadline", CStr(Value), trans)
    '    End Set
    'End Property
    'Public Property MM_EP2R_Deadline(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Date
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='MM_EP2R_Deadline'"
    '        If d.Count > 0 Then
    '            Return CDate(d(0)(1))
    '        Else
    '            SetProp("MM_EP2R_Deadline", DAT_NULL.ToString, trans)
    '            Return DAT_NULL
    '        End If
    '    End Get
    '    Set(ByVal Value As Date)
    '        SetProp("MM_EP2R_Deadline", CStr(Value), trans)
    '    End Set
    'End Property


    Public Property WF_SCANDIR() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='WF_SCANDIR" & "_" & UCase(System.Environment.MachineName) & "'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("WF_SCANDIR" & "_" & UCase(System.Environment.MachineName), CStr(Value))
        End Set
    End Property

    Public Property Online_User() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Online_User" & "_" & UCase(System.Environment.UserName) & "'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("Online_User" & "_" & UCase(System.Environment.UserName), CStr(Value))
        End Set
    End Property
    Public Property Online_PW() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Online_PW" & "_" & UCase(System.Environment.UserName) & "'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("Online_PW" & "_" & UCase(System.Environment.UserName), CStr(Value))
        End Set
    End Property

    'Public Property WF_LastScan() As String
    '    Get
    '        If Not Isloaded Then LoadConfig()
    '        d.RowFilter = "param='WF_LastScan" & "_" & UCase(Setting.CurrentUser) & "'"
    '        If d.Count > 0 Then
    '            Return d(0)(1).ToString
    '        Else
    '            'Scantyp|Post-ID|AbsenderID|Absender|Abrechnung|Dokuemntentyp
    '            Return Scantypus.Recherche & "|-1|-1|||" & Dokumenttypus.Recherche_Ein & "|" & Date.Today

    '        End If
    '    End Get
    '    Set(ByVal Value As String)
    '        SetProp("WF_LastScan" & "_" & UCase(Setting.CurrentUser), CStr(Value))
    '    End Set
    'End Property

    Public Property WF_HTMLEditor() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='WF_HTMLEditor" & "_" & UCase(Setting.CurrentUser) & "'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("WF_HTMLEditor" & "_" & UCase(Setting.CurrentUser), CStr(Value))
        End Set
    End Property
    Public Property User_EMail() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='User_EMail" & "_" & UCase(Setting.CurrentUser) & "'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return "administrator@scheckheft-gesundheit.at"

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("User_EMail" & "_" & UCase(Setting.CurrentUser), CStr(Value))
        End Set
    End Property


    Public Property WF_DBL_Vorbelegen() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='WF_DBL_Vorbelegen" & "_" & UCase(Setting.CurrentUser) & "'"
            If d.Count > 0 Then
                Return IIf(d(0)(1) <> 0, True, False)
            Else
                Return True

            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("WF_DBL_Vorbelegen" & "_" & UCase(Setting.CurrentUser), IIf(Value, 1, 0))
        End Set
    End Property

    Public Property Impfdoku_Überspringen() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Impfdoku_Überspringen" & "_" & UCase(Setting.CurrentUser) & "'"
            If d.Count > 0 Then
                Return IIf(d(0)(1) <> 0, True, False)
            Else
                Return False

            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("Impfdoku_Überspringen" & "_" & UCase(Setting.CurrentUser), IIf(Value, 1, 0))
        End Set
    End Property


    Public Property AutoProtokollGedächtnis() As Boolean
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='AutoProtokollGedächtnis" & "_" & UCase(Setting.CurrentUser) & "'"
            If d.Count > 0 Then
                Return IIf(d(0)(1) <> 0, True, False)
            Else
                Return True

            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("AutoProtokollGedächtnis" & "_" & UCase(Setting.CurrentUser), IIf(Value, 1, 0))
        End Set
    End Property


    Public Property WF_NASDIR(Optional trans As SqlClient.SqlTransaction = Nothing) As String
        Get
            If Not Isloaded Then LoadConfig(trans)
            d.RowFilter = "param='WF_NASDIR'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("WF_NASDIR", CStr(Value), trans)
        End Set
    End Property

    'Public ReadOnly Property GetNextPostID(Anzahl As Integer) As Integer
    '    Get
    '        Dim GotRec As Boolean = False
    '        Dim tb As DataTable
    '        Dim MaxID As Integer
    '        Do Until GotRec
    '            tb = db_con.GetRecordSet("select max(pi_bis) from wf_postid")
    '            If tb.Rows.Count = 0 Then
    '                MaxID = 0
    '            Else
    '                If IsDBNull(tb.Rows(0)(0)) Then
    '                    MaxID = 0
    '                Else
    '                    MaxID = tb.Rows(0)(0)
    '                End If
    '            End If
    '            Try
    '                db_con.FireSQL("insert into wf_postid (pi_von,pi_bis) values (" & MaxID + 1 & "," & MaxID + Anzahl & ")")
    '                GotRec = True
    '            Catch ex As Exception

    '            End Try
    '        Loop


    '        Return MaxID + 1


    '    End Get
    'End Property



    Public Property WF_NASDIR_ERSATZ(Optional trans As SqlClient.SqlTransaction = Nothing) As String
        Get
            If Not Isloaded Then LoadConfig(trans)
            d.RowFilter = "param='WF_NASDIR_ERSATZ'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("WF_NASDIR_ERSATZ", CStr(Value), trans)
        End Set
    End Property


    Public Property BIC_STelle() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='BIC_STelle'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("BIC_STelle", CStr(Value))
        End Set
    End Property
    Public Property IBAN_STelle() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='IBAN_STelle'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("IBAN_STelle", CStr(Value))
        End Set
    End Property


    'Public Property WF_NASServerInBetrieb(Optional trans As SqlClient.SqlTransaction = Nothing) As Integer
    '    Get
    '        If Not Isloaded Then LoadConfig(trans)
    '        d.RowFilter = "param='WF_NASServerInBetrieb'"
    '        If d.Count > 0 Then
    '            Return CInt(d(0)(1))
    '        Else
    '            SetProp("WF_NASServerInBetrieb", "0", trans)
    '            Return DokumentenServerInBetrieb.Hauptserver
    '        End If
    '    End Get
    '    Set(ByVal Value As Integer)
    '        SetProp("WF_NASServerInBetrieb", CStr(Value), trans)
    '    End Set
    'End Property


    Public Property WF_RechercheID() As Integer
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='WF_RechercheID'"
            If d.Count > 0 Then
                SetProp("WF_RechercheID", CInt(d(0)(1)) + 1)
                Return CInt(d(0)(1)) + 1
            Else
                SetProp("WF_RechercheID", "0")
                Return 0
            End If
        End Get
        Set(ByVal Value As Integer)
            SetProp("WF_RechercheID", CStr(Value))
        End Set
    End Property


    Public Property CriticalMessage_ScanIntvl() As Integer
        Get
            Try
                System.Threading.Monitor.Enter(Me)
                If Not Isloaded Then LoadConfig()
                Dim d As New DataView(tb)
                d.RowFilter = "param='CriticalMessage_ScanIntvl'"
                If d.Count > 0 Then
                    Return d(0)(1)
                Else
                    Return 10
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                System.Threading.Monitor.Exit(Me)

            End Try
        End Get
        Set(ByVal Value As Integer)
            SetProp("CriticalMessage_ScanIntvl", Value)
        End Set
    End Property

    Public Property ShowPDFToolbar(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Boolean
        Get
            If Not Isloaded Then LoadConfig(trans)
            d.RowFilter = "param='ShowPDFToolbar'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) <> 0, True, False))
            Else
                SetProp("ShowPDFToolbar", "0", trans)
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("ShowPDFToolbar", IIf(Value, "1", "0").ToString, trans)
        End Set
    End Property
    Public Property PDF_DBL_Darstellung_optimiert(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Boolean
        Get
            If Not Isloaded Then LoadConfig(trans)
            d.RowFilter = "param='PDF_DBL_Darstellung_optimiert'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) <> 0, True, False))
            Else
                SetProp("PDF_DBL_Darstellung_optimiert", "0", trans)
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("PDF_DBL_Darstellung_optimiert", IIf(Value, "1", "0").ToString, trans)
        End Set
    End Property

    Public Property RecherchenAlleBonsAnzeigen(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Boolean
        Get
            If Not Isloaded Then LoadConfig(trans)
            d.RowFilter = "param='RecherchenAlleBonsAnzeigen'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) <> 0, True, False))
            Else
                SetProp("RecherchenAlleBonsAnzeigen", "0", trans)
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("RecherchenAlleBonsAnzeigen", IIf(Value, "1", "0").ToString, trans)
        End Set
    End Property



    Public Property WF_Papierkorb(Optional ByVal trans As SqlClient.SqlTransaction = Nothing) As Boolean
        Get
            If Not Isloaded Then LoadConfig(trans)
            d.RowFilter = "param='WF_Papierkorb'"
            If d.Count > 0 Then
                Return CBool(IIf(CInt(d(0)(1)) <> 0, True, False))
            Else
                SetProp("WF_Papierkorb", "0", trans)
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            SetProp("WF_Papierkorb", IIf(Value, "1", "0").ToString, trans)
        End Set
    End Property

    Public Sub SetLetzteAbrechnung(ZeitraumEingang As String, VonEingang As String, BisEingang As String, ZeitraumImpfdatum As String, VonIMpfdatum As String, BisImpfdatum As String)
        Dim SignAbr As String = ZeitraumEingang & "," & VonEingang & "," & BisEingang & "|" & ZeitraumImpfdatum & "," & VonIMpfdatum & "," & BisImpfdatum
        If SignAbr = LetzteAbrechnung1 Or SignAbr = LetzteAbrechnung2 Or SignAbr = LetzteAbrechnung3 Or SignAbr = LetzteAbrechnung4 Or SignAbr = LetzteAbrechnung5 Then Return
        LetzteAbrechnung5 = LetzteAbrechnung4
        LetzteAbrechnung4 = LetzteAbrechnung3
        LetzteAbrechnung3 = LetzteAbrechnung2
        LetzteAbrechnung2 = LetzteAbrechnung1
        LetzteAbrechnung1 = SignAbr



    End Sub

    Public Function SetLetzterBericht(ZeitraumEingang As String, VonEingang As String, BisEingang As String, ZeitraumImpfdatum As String, VonIMpfdatum As String, BisImpfdatum As String) As String
        Dim SignAbr As String = Zeitfilter(ZeitraumEingang, VonEingang, BisEingang, ZeitraumImpfdatum, VonIMpfdatum, BisImpfdatum)
        If SignAbr = LetzteBerichte1 Or SignAbr = LetzteBerichte2 Or SignAbr = LetzteBerichte3 Or SignAbr = LetzteBerichte4 Or SignAbr = LetzteBerichte5 Then Return SignAbr
        LetzteBerichte5 = LetzteBerichte4
        LetzteBerichte4 = LetzteBerichte3
        LetzteBerichte3 = LetzteBerichte2
        LetzteBerichte2 = LetzteBerichte1
        LetzteBerichte1 = SignAbr
        Return SignAbr


    End Function

    Public ReadOnly Property CurAbrechnugsQuartal As List(Of Date)
        Get
            Dim h As Date = Date.Today
            Dim y As Integer = Date.Today.Year()
            Dim l As New List(Of Date)


            Dim q1_b As Date = CDate("16.01." & y)
            Dim q1_e As Date = CDate("15.04." & y)

            Dim q2_b As Date = CDate("16.04." & y)
            Dim q2_e As Date = CDate("15.07." & y)

            Dim q3_b As Date = CDate("16.07." & y)
            Dim q3_e As Date = CDate("15.10." & y)

            Dim q4_b As Date = CDate("16.10." & y)
            Dim q4_e As Date = CDate("15.01." & y + 1)


            If Date.Today.Month = 1 And Date.Today.Day <= 15 Then
                q4_b = CDate("16.10." & y - 1)
                q4_e = CDate("15.01." & y)


            End If


            If Date.Compare(h, q1_b) >= 0 And Date.Compare(h, q1_e) <= 0 Then

                l.Add(q1_b)
                l.Add(q1_e)
            ElseIf Date.Compare(h, q2_b) >= 0 And Date.Compare(h, q2_e) <= 0 Then
                l.Add(q2_b)
                l.Add(q2_e)

            ElseIf Date.Compare(h, q3_b) >= 0 And Date.Compare(h, q3_e) <= 0 Then
                l.Add(q3_b)
                l.Add(q3_e)

            Else
                l.Add(q4_b)
                l.Add(q4_e)


            End If

            Return l

        End Get
    End Property


    Public Function Zeitfilter(ZeitraumEingang As String, VonEingang As String, BisEingang As String, ZeitraumImpfdatum As String, VonIMpfdatum As String, BisImpfdatum As String) As String
        Return ZeitraumEingang & "," & VonEingang & "," & BisEingang & "|" & ZeitraumImpfdatum & "," & VonIMpfdatum & "," & BisImpfdatum

    End Function





#If NET_2_0 Then
#Else
    'Public Sub Liste_LetzteAbrechnungen(cb As ComboBox)

    '    Dim stan As New StandardTab
    '    stan.AddItem(NO_OPTION, NO_VAL)
    '    If LetzteAbrechnung1 <> "" Then stan.AddItem(GetAbrechungstext(LetzteAbrechnung1), LetzteAbrechnung1)
    '    If LetzteAbrechnung2 <> "" Then stan.AddItem(GetAbrechungstext(LetzteAbrechnung2), LetzteAbrechnung2)
    '    If LetzteAbrechnung3 <> "" Then stan.AddItem(GetAbrechungstext(LetzteAbrechnung3), LetzteAbrechnung3)
    '    If LetzteAbrechnung4 <> "" Then stan.AddItem(GetAbrechungstext(LetzteAbrechnung4), LetzteAbrechnung4)
    '    If LetzteAbrechnung5 <> "" Then stan.AddItem(GetAbrechungstext(LetzteAbrechnung5), LetzteAbrechnung5)

    '    Dim q1_von As Date
    '    Dim q1 As Integer


    '    Select Case Date.Today.Month
    '        Case 1, 2, 3
    '            q1_von = CDate("16.10." & Date.Today.Year - 1)
    '            q1 = 4
    '        Case 4, 5, 6
    '            q1_von = CDate("16.01." & Date.Today.Year)
    '            q1 = 1
    '        Case 7, 8, 9
    '            q1_von = CDate("16.04." & Date.Today.Year)
    '            q1 = 2
    '        Case 10, 11, 12
    '            q1_von = CDate("16.07." & Date.Today.Year)
    '            q1 = 2
    '    End Select


    '    stan.AddItem("Q" & q1 & " " & q1_von.Year & ", Eingang: " & q1_von & " - " & q1_von.AddMonths(3).AddDays(-1) & ", Impfdatum: Alle", _
    '                 "anderer Zeitraum," & q1_von & "," & q1_von.AddMonths(3).AddDays(-1) & "|Alle,,")
    '    stan.AddItem("Q" & IIf(q1 - 1 < 0, q1 - 1 + 4, q1 - 1) & " " & q1_von.AddMonths(-3).Year & ", Eingang: " & q1_von.AddMonths(-3) & " - " & q1_von.AddDays(-1) & ", Impfdatum: Alle", _
    '                 "anderer Zeitraum," & q1_von.AddMonths(-3) & "," & q1_von.AddDays(-1) & "|Alle,,")











    '    Try


    '        cb.ValueMember = StandardTab.FLD_PARAM
    '        cb.DisplayMember = StandardTab.FLD_VAL
    '        cb.DataSource = stan.GetTab


    '        If stan.GetTab.Rows.Count > 1 Then cb.SelectedIndex = 1

    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try



    'End Sub





    'Public Sub Liste_LetzteBerichte(cb As ComboBox)

    '    Dim stan As New StandardTab
    '    stan.AddItem(NO_OPTION, NO_VAL)
    '    If LetzteBerichte1 <> "" Then stan.AddItem(GetAbrechungstext(LetzteBerichte1), LetzteBerichte1)
    '    If LetzteBerichte2 <> "" Then stan.AddItem(GetAbrechungstext(LetzteBerichte2), LetzteBerichte2)
    '    If LetzteBerichte3 <> "" Then stan.AddItem(GetAbrechungstext(LetzteBerichte3), LetzteBerichte3)
    '    If LetzteBerichte4 <> "" Then stan.AddItem(GetAbrechungstext(LetzteBerichte4), LetzteBerichte4)
    '    If LetzteBerichte5 <> "" Then stan.AddItem(GetAbrechungstext(LetzteBerichte5), LetzteBerichte5)


    '    Dim q1_von_impfdatum As Date
    '    Dim q1_von_eingang As Date
    '    Dim q1_bis_impfdatum As Date
    '    Dim q1_bis_eingang As Date
    '    Dim q1 As Integer


    '    Select Case Date.Today.Month
    '        Case 1, 2, 3
    '            q1_von_eingang = CDate("16.10." & Date.Today.Year - 1)
    '            q1_von_impfdatum = CDate("01.07." & Date.Today.Year - 1)
    '            q1 = 4
    '        Case 4, 5, 6
    '            q1_von_eingang = CDate("16.01." & Date.Today.Year)
    '            q1_von_impfdatum = CDate("01.10." & Date.Today.Year - 1)
    '            q1 = 1
    '        Case 7, 8, 9
    '            q1_von_eingang = CDate("16.04." & Date.Today.Year)
    '            q1_von_impfdatum = CDate("01.01." & Date.Today.Year)
    '            q1 = 2
    '        Case 10, 11, 12
    '            q1_von_eingang = CDate("16.07." & Date.Today.Year)
    '            q1_von_impfdatum = CDate("01.04." & Date.Today.Year - 1)

    '            q1 = 2
    '    End Select

    '    q1_bis_impfdatum = q1_von_impfdatum.AddMonths(6).AddDays(-1)
    '    q1_bis_eingang = q1_von_eingang.AddMonths(3).AddDays(-1)

    '    Dim q2 As Integer = IIf(q1 - 1 < 0, q1 - 1 + 4, q1 - 1)

    '    Dim q2_von_eingang As Date = q1_von_eingang.AddMonths(-3)
    '    Dim q2_von_impfdatum As Date = q1_von_impfdatum.AddMonths(-3)

    '    Dim q2_bis_impfdatum As Date = q1_bis_impfdatum.AddMonths(-3)
    '    Dim q2_bis_eingang As Date = q1_bis_eingang.AddMonths(-3)



    '    stan.AddItem("Q" & q1 & " " & q1_von_eingang.Year & ", Eingang: " & q1_von_eingang & " - " & q1_bis_eingang & ", Impfdatum: " & q1_von_impfdatum & " - " & q1_bis_impfdatum, _
    '                 "anderer Zeitraum," & q1_von_eingang & "," & q1_bis_eingang & "|anderer Zeitraum," & q1_von_impfdatum & "," & q1_bis_impfdatum)
    '    stan.AddItem("Q" & q2 & " " & q2_von_eingang.Year & ", Eingang: " & q2_von_eingang & " - " & q2_bis_eingang & ", Impfdatum: " & q2_von_impfdatum & " - " & q2_bis_impfdatum, _
    '                 "anderer Zeitraum," & q2_von_eingang & "," & q2_bis_eingang & "|anderer Zeitraum," & q2_von_impfdatum & "," & q2_bis_impfdatum)






    '    Try


    '        cb.ValueMember = StandardTab.FLD_PARAM
    '        cb.DisplayMember = StandardTab.FLD_VAL
    '        cb.DataSource = stan.GetTab


    '        If stan.GetTab.Rows.Count > 1 Then cb.SelectedIndex = 1

    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try



    'End Sub


#End If

    'Private Function GetAbrechungstext(SignAbr As String) As String
    '    If SignAbr = "" Then Return ""
    '    Dim arrZ As String() = SignAbr.Split("|")
    '    Dim arrEingang As String() = arrZ(0).Split(",")
    '    Dim arrImpfdatum As String() = arrZ(1).Split(",")

    '    Return "Eingang: " & IIf(arrEingang(0) = ZR_AND_ZR, arrEingang(1) & " - " & arrEingang(2), arrEingang(0)) & "; Impfdatum: " & IIf(arrImpfdatum(0) = ZR_AND_ZR, arrImpfdatum(1) & " - " & arrImpfdatum(2), arrImpfdatum(0))





    'End Function

    Private Property LetzteAbrechnung1() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LetzteAbrechnung1'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("LetzteAbrechnung1", CStr(Value))
        End Set
    End Property
    Private Property LetzteAbrechnung2() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LetzteAbrechnung2'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("LetzteAbrechnung2", CStr(Value))
        End Set
    End Property
    Private Property LetzteAbrechnung3() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LetzteAbrechnung3'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("LetzteAbrechnung3", CStr(Value))
        End Set
    End Property

    Private Property LetzteAbrechnung4() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LetzteAbrechnung4'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("LetzteAbrechnung4", CStr(Value))
        End Set
    End Property
    Private Property LetzteAbrechnung5() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LetzteAbrechnung5'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("LetzteAbrechnung5", CStr(Value))
        End Set
    End Property


    Private Property LetzteBerichte1() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LetzteBerichte1'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("LetzteBerichte1", CStr(Value))
        End Set
    End Property

    Private Property LetzteBerichte2() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LetzteBerichte2'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("LetzteBerichte2", CStr(Value))
        End Set
    End Property


    Private Property LetzteBerichte3() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LetzteBerichte3'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("LetzteBerichte3", CStr(Value))
        End Set
    End Property
    Private Property LetzteBerichte4() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LetzteBerichte4'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("LetzteBerichte4", CStr(Value))
        End Set
    End Property
    Private Property LetzteBerichte5() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='LetzteBerichte5'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ""

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("LetzteBerichte5", CStr(Value))
        End Set
    End Property


    Public Property AnredeAbschluss() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='AnredeAbschluss'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return ","

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("AnredeAbschluss", CStr(Value))
        End Set
    End Property


    Public Property Nr_GH() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Nr_GH'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return "Heft Nr."

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("Nr_GH", CStr(Value))
        End Set
    End Property
    Public Property Nr_EA() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='Nr_EA'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return "Bonbogen Nr."

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("Nr_EA", CStr(Value))
        End Set
    End Property

    Public Property AnredeDirektorAnonym() As String
        Get
            If Not Isloaded Then LoadConfig()
            d.RowFilter = "param='AnredeDirektorAnonym'"
            If d.Count > 0 Then
                Return d(0)(1).ToString
            Else
                Return "Sehr geehrte/r Herr/Frau Direktor/in"

            End If
        End Get
        Set(ByVal Value As String)
            SetProp("AnredeDirektorAnonym", CStr(Value))
        End Set
    End Property

End Class




Public Class SimpleCrypto



    Private Function EncryptString(ByVal clearText As Byte(), ByVal Key As Byte(), ByVal IV As Byte()) As Byte()



        Dim ms As MemoryStream = New MemoryStream()
        Dim alg As Rijndael = Rijndael.Create()
        alg.Key = Key
        alg.IV = IV
        Dim cs As CryptoStream = New CryptoStream(ms, alg.CreateEncryptor(), CryptoStreamMode.Write)
        cs.Write(clearText, 0, clearText.Length)
        cs.Close()
        Dim encryptedData As Byte() = ms.ToArray()
        Return encryptedData
    End Function


    ''' <summary>
    ''' Encrypts the string.
    ''' </summary>
    ''' <param name="clearText">The clear text.</param>
    ''' <param name="Password">The password.</param>
    ''' <returns></returns>
    Public Function EncryptString(ByVal clearText As String, ByVal Password As String) As String

        Dim clearBytes As Byte() = System.Text.Encoding.Unicode.GetBytes(clearText)
        Dim pdb As Rfc2898DeriveBytes = New Rfc2898DeriveBytes(Password, _
            New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, &H65, &H64, &H76, &H65, &H64, &H65, &H76})
        Dim encryptedData As Byte() = EncryptString(clearBytes, pdb.GetBytes(32), pdb.GetBytes(16))
        Return Convert.ToBase64String(encryptedData)
    End Function

    ''' <summary>
    ''' Decrypts the string.
    ''' </summary>
    ''' <param name="cipherData">The cipher data.</param>
    ''' <param name="Key">The key.</param>
    ''' <param name="IV">The IV.</param>
    ''' <returns></returns>
    Private Function DecryptString(ByVal cipherData As Byte(), ByVal Key As Byte(), ByVal IV As Byte()) As Byte()

        Dim ms As MemoryStream = New MemoryStream()
        Dim alg As Rijndael = Rijndael.Create()
        alg.Key = Key
        alg.IV = IV
        Dim cs As CryptoStream = New CryptoStream(ms, alg.CreateDecryptor(), CryptoStreamMode.Write)
        cs.Write(cipherData, 0, cipherData.Length)
        cs.Close()
        Dim decryptedData As Byte() = ms.ToArray()
        Return decryptedData
    End Function

    ''' <summary>
    ''' Decrypts the string.
    ''' </summary>
    ''' <param name="cipherText">The cipher text.</param>
    ''' <param name="Password">The password.</param>
    ''' <returns></returns>
    Public Function DecryptString(ByVal cipherText As String, ByVal Password As String) As String

        Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
        Dim pdb As Rfc2898DeriveBytes = New Rfc2898DeriveBytes(Password, _
            New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, &H65, &H64, &H76, &H65, &H64, &H65, &H76})
        Dim decryptedData As Byte() = DecryptString(cipherBytes, pdb.GetBytes(32), pdb.GetBytes(16))
        Return System.Text.Encoding.Unicode.GetString(decryptedData)
    End Function

End Class
