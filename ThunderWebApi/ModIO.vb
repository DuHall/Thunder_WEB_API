Imports System.Collections.Specialized

Module ModIO

    ''**********************************Local Variable Definition**********************************
    ''these variables are defined in web.config
    Public g_intNumberOfRetriesToAccessDatabase As Integer = ConfigurationManager.AppSettings("inNumberOfRetriesToAccessDatabase")
    Public g_ConnectionToUse As String = ConfigurationManager.AppSettings("ConnectionType")
    Public g_ConnectionSchema As String = ConfigurationManager.AppSettings("schema")
    Public g_blnAbort As Boolean = False
    Private m_nvcTables As New NameValueCollection               ' key=tablename, data=index to nvcFields 
    Private m_nvcFields() As NameValueCollection                 ' (=index#fromnvcTables),Key=FieldName,Data=index to nvcFieldAttributes
    Private m_nvcFieldAttributes() As NameValueCollection        ' (=index#fromnvcFields),Key=AttributeName,Data=AttributeValue
    Private m_IntLastRecid As Integer                          ' passes recid from insert down to recid request (in MSSQL only)

    Public Function g_IO_Execute_SQL(ByVal strSQL As String, ByRef blnReturnErrorCode As Boolean) As DataTable
        ''Function is used to query the database.
        Dim blnInsert As Boolean = False
        If strSQL.ToUpper.IndexOf("INSERT ") = 0 Then
            Try
                ' convert MySQL True/False references to 0/1
                Dim strFields As String = strSQL.Substring(0, strSQL.ToUpper.IndexOf(" VALUES"))
                Dim strValues As String = strSQL.Substring(strSQL.ToUpper.IndexOf(" VALUES"))
                strValues = strValues.Replace(", True", ", 1")
                strValues = strValues.Replace(",True", ",1")
                strValues = strValues.Replace(", TRUE", ", 1")
                strValues = strValues.Replace(",TRUE", ",1")
                strValues = strValues.Replace(", False", ", 0")
                strValues = strValues.Replace(",False", ",0")
                strValues = strValues.Replace(", FALSE", ", 0")
                strValues = strValues.Replace(",FALSE", ",0")
                strSQL = strFields & strValues

                strSQL &= ";Select @@IDENTITY as ID;"    ' on MSSQL inserts must retrieve the new RECID now
                blnInsert = True
            Catch : End Try
        ElseIf strSQL.ToUpper.IndexOf("UPDATE ") = 0 Then
            Try
                ' convert MySQL True/False references to 0/1
                Dim strFields As String = strSQL.Substring(0, strSQL.ToUpper.IndexOf(" SET"))
                Dim strValues As String = strSQL.Substring(strSQL.ToUpper.IndexOf(" SET"))
                strValues = strValues.Replace("= True", "=1")
                strValues = strValues.Replace("=True", "=1")
                strValues = strValues.Replace("= TRUE", "= 1")
                strValues = strValues.Replace("=TRUE", "=1")
                strValues = strValues.Replace("= False", "= 0")
                strValues = strValues.Replace("=False", "=0")
                strValues = strValues.Replace("= FALSE", "= 0")
                strValues = strValues.Replace("=FALSE", "=0")
                strSQL = strFields & strValues
            Catch : End Try
        ElseIf strSQL.ToUpper.IndexOf("SELECT ") = 0 Then
            If strSQL.ToUpper.IndexOf("CONCAT(") = -1 Then
            Else
                Call convertConcatStatement(strSQL)
            End If
        End If

        Try
            Dim tblTemp As DataTable = IO_Execute_MSSQL(strSQL, blnReturnErrorCode)
            If strSQL.ToUpper.IndexOf("SELECT ") = 0 Then
                ' trim all text entries
                For Each rowTemp As DataRow In tblTemp.Rows
                    For i = 0 To tblTemp.Columns.Count - 1
                        Try : rowTemp(i) = Trim(rowTemp(i)) : Catch : End Try
                    Next
                Next
            ElseIf blnInsert Then
                System.Web.HttpContext.Current.Session("NewMSSQLRECID") = tblTemp.Rows(0)("ID")

            End If
            Return tblTemp
        Catch ex As Exception
            System.Web.HttpContext.Current.Session("SQLERROR") = ex.Message.ToString & " - " & ex.StackTrace.ToString
        End Try

    End Function

    Private Sub convertConcatStatement(ByRef strSQL As String)

        Dim strEndOfFieldMarker As String = ","
        Dim blnThisCharIsAQuote As Boolean = False
        Dim intIndex As Integer = 0
        Dim strFieldValue As String = ""
        Dim blnEndOfFieldFound As Boolean = False
        Dim intLargestRowIndex As Integer = 0
        Dim blnConversionComplete As Boolean = False



        Dim intStart As Integer = strSQL.ToUpper.IndexOf("CONCAT(")
        Dim strFrontSQL As String = Left(strSQL, intStart) & "("
        Dim strBackSQL As String = Mid(strSQL, intStart + 8)
        Dim I As Integer = 1

        Do Until blnConversionComplete
            Dim strCurrentChar As String = Mid(strBackSQL, I, 1)

            blnEndOfFieldFound = False

            If strCurrentChar = ")" And (strEndOfFieldMarker = "," Or strEndOfFieldMarker = "") Then

                Exit Do
            ElseIf strEndOfFieldMarker = "" Then
                ' end the middle of evaluating a string

                Select Case strCurrentChar
                    Case ","
                        strEndOfFieldMarker = ","
                        strFrontSQL &= " + "
                    Case "'"
                        strEndOfFieldMarker = "'"
                End Select
            ElseIf strEndOfFieldMarker = "'" Then

                ' this is a string under evaluation
                If Mid(strBackSQL, I, 2) = "''" Then
                    'this is a quote within the string
                    blnThisCharIsAQuote = True
                    I += 1   ' this quote takes up two spaces
                Else
                    blnThisCharIsAQuote = False
                End If

                If blnThisCharIsAQuote Then
                    strFrontSQL &= "'"
                Else
                    ' is this the end of the string?
                    If strCurrentChar = strEndOfFieldMarker Then
                        strFrontSQL &= "'"
                        strEndOfFieldMarker = ""
                        intIndex += 1
                    Else
                        ' build current field char by char
                        strFrontSQL &= strCurrentChar
                    End If

                End If


            Else
                ' this should be a field name (might not be if there are spaces before the starting quote of a string)
                If strCurrentChar = "'" Then
                    ' oops, this is a string after all
                    strEndOfFieldMarker = "'"
                    strFrontSQL &= "'"
                ElseIf strCurrentChar = strEndOfFieldMarker Then
                    strEndOfFieldMarker = ","
                    strFrontSQL &= " + "
                    intIndex += 1
                Else
                    strFrontSQL &= strCurrentChar
                End If
            End If

            I += 1
        Loop

        strSQL = strFrontSQL & Mid(strBackSQL, I)

    End Sub
    '
    Private Function IO_Execute_MSSQL(ByVal strSQL As String, ByRef ReturnCode As Boolean) As DataTable

        'System.Web.HttpContext.Current.Response.Write(ConfigurationManager.ConnectionStrings("ConnectionString").ToString & "<br>")

        Try

            '            System.Web.HttpContext.Current.Response.Write("IO_Execute: 1<br>")
            Dim adapt As New SqlClient.SqlDataAdapter

            Dim blnProcessWasSuccessful As Boolean = True
            Dim blnRetry As Boolean = True
            Dim tblDatatable As New DataTable
            '           System.Web.HttpContext.Current.Response.Write("IO_Execute: 2<br>")

            Dim g_strConnectionStringMSSQL As New System.Data.SqlClient.SqlConnection

            ' System.Web.HttpContext.Current.Response.Write("IO_Execute: 3<br>")
            g_strConnectionStringMSSQL = New System.Data.SqlClient.SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ToString)

            'System.Web.HttpContext.Current.Response.Write("IO_Execute: 4<br>")

            Dim command As New SqlClient.SqlCommand(strSQL, g_strConnectionStringMSSQL)
            Debug.WriteLine(strSQL)
            'System.Web.HttpContext.Current.Response.Write("IO_Execute: 5<br>")

            command.CommandTimeout = 600
            adapt.SelectCommand = command

            Dim strErrorCode As String = ""
            For i = 1 To g_intNumberOfRetriesToAccessDatabase
                Debug.WriteLine(strSQL)
                adapt.Fill(tblDatatable)
                'System.Web.HttpContext.Current.Response.Write("IO_Execute: 6<br>")

                blnRetry = False
                ReturnCode = True
                Return tblDatatable
                Exit For
            Next
        Catch ex As Exception
            ReturnCode = False
            Debug.WriteLine(ex.Message)
            'System.Web.HttpContext.Current.Response.Write("IO_Execute: 7<br>")
            System.Web.HttpContext.Current.Session("SystemErrorMessage") = ex.Message

            'System.Web.HttpContext.Current.Response.Write(ConfigurationManager.ConnectionStrings("ConnectionString").ToString & "<br>")
            'System.Web.HttpContext.Current.Response.Write(ex.Message & "<br>")
        End Try
        Return Nothing
    End Function

End Module