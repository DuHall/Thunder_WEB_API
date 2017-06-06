Imports System.Net.Mail
Imports System.Net.Mail.SmtpClient
Imports System.Web.HttpContext
Imports System.Configuration
Imports System.Web.Services
Imports System.Data.SqlClient
Imports System.IO

Imports System.Security.Principal
Imports System.DirectoryServices.AccountManagement

Module ModMain

    Dim g_EmailHostIPAddress As String = ConfigurationManager.AppSettings("EmailIPAddress")
    Dim g_EmailPort As String = ConfigurationManager.AppSettings("EmailPort")
    Dim g_EmailUserName As String = ConfigurationManager.AppSettings("EmailUserName")
    Dim g_EmailPassword As String = ConfigurationManager.AppSettings("EmailPassword")
    Public g_EmailAudit As String = ConfigurationManager.AppSettings("EmailAudit")
    Public g_EmailAuditEmail As String = ConfigurationManager.AppSettings("EmailAuditEmail")
    Public g_EmailFromAddress As String = ConfigurationManager.AppSettings("EmailFromAddress")
    Dim g_EmailEnabled As Boolean = IIf(UCase(ConfigurationManager.AppSettings("EmailEnabled")) = "TRUE", True, False)
    Dim g_SiteLocation As String = ConfigurationManager.AppSettings("siteLocation")
    Public g_EfilesReferralsBaseDirectory As String = ConfigurationManager.AppSettings("EfilesReferralsBaseDirectory")
    Public g_PDFTempDirectory As String = ConfigurationManager.AppSettings("TempPDFDir")
    Public g_StringArrayOuterSplitParameter = "||"
    Public g_StringArrayValueSplitParameter = "~~"
    Public g_intUserRecid = -1
    Public g_strFormObjectTypes = "input+select+textarea"
    Public g_Debug As Boolean = False
    Public g_portalURL As String = ConfigurationManager.AppSettings("portalURL")
    Public g_uploadURL As String = HttpContext.Current.Server.MapPath("/" & ConfigurationManager.AppSettings("uploadURL"))
    Public g_portalEnvironment As String = ConfigurationManager.AppSettings("portalEnvironment")
    Public g_EmailURL As String = ConfigurationManager.AppSettings("EmailURL")

    Public Sub g_SendEmail(ByVal ToAddress As String, ByVal Subject As String, ByVal Message As String)
        g_SendEmail(ToAddress:=ToAddress, Subject:=Subject, Message:=Message, FromAddress:="", BCC_Address:="", CC_Address:="", Attachment:="", emailTemplate:="ITG_Form.html")
    End Sub

    Public Sub g_SendEmail(ByVal ToAddress As String, ByVal Subject As String, ByVal Message As String, CC_Address As String)
        g_SendEmail(ToAddress:=ToAddress, FromAddress:="", CC_Address:=CC_Address, BCC_Address:="", Subject:=Subject, Attachment:="", Message:=Message, emailTemplate:="ITG_Form.html")
    End Sub

    Public Sub g_SendEmail(ByVal ToAddress As String, ByVal FromAddress As String, ByVal CC_Address As String, ByVal BCC_Address As String, ByVal Subject As String, ByVal Message As String, ByVal emailTemplate As String)
        g_SendEmail(ToAddress:=ToAddress, FromAddress:="", CC_Address:="", BCC_Address:="", Subject:=Subject, Attachment:="", Message:=Message, emailTemplate:="ITG_Form.html")
    End Sub

    Public Sub g_SendEmail(ByVal ToAddress As String, ByVal FromAddress As String, ByVal CC_Address As String, ByVal BCC_Address As String, ByVal Subject As String, ByVal Attachment As String, ByVal Message As String, ByVal emailTemplate As String)
        Debug.WriteLine("Email Module -- To: " & ToAddress & "  Subject: " & Subject & " Attachment: " & Attachment & "  Message: " & Message)
        If ConfigurationManager.AppSettings("EmailEnabled") = True Then
            Dim Mail As New System.Net.Mail.MailMessage
            Mail.Subject = Subject
            If ToAddress = "" Then
                Debug.Print("ModMain (g_SendEmail): No email address provided. Can't send it.")
            Else
                For Each strEmailAddress As String In Split(ToAddress, ";")
                    If strEmailAddress <> "" And strEmailAddress.Length > 5 Then
                        Mail.To.Add(strEmailAddress)
                    End If
                Next
                ''Add CC
                If CC_Address <> "" Then
                    For Each strCCEmailAddress As String In Split(CC_Address, ";")
                        If strCCEmailAddress <> "" And strCCEmailAddress.Length > 5 Then
                            Mail.CC.Add(strCCEmailAddress.Replace(",", ""))
                        End If
                    Next
                End If

                ''Add BCC
                If BCC_Address <> "" Then
                    For Each strBCCEmailAddress As String In Split(BCC_Address, ";")
                        If strBCCEmailAddress <> "" And strBCCEmailAddress.Length > 5 Then
                            Mail.CC.Add(strBCCEmailAddress)
                        End If
                    Next
                End If

                ''Add From (If blank, use default from web.config
                If FromAddress = "" Then
                    Mail.From = New System.Net.Mail.MailAddress(g_EmailFromAddress)
                Else
                    Mail.From = New System.Net.Mail.MailAddress(FromAddress)
                End If


                If emailTemplate = "" Then
                Else
                    Dim rdrEmailTemplate As StreamReader = New StreamReader(HttpContext.Current.Server.MapPath("/" & emailTemplate))
                    Dim htmlBody As String = rdrEmailTemplate.ReadToEnd
                    Message = htmlBody.Replace("{Body}", Message)
                End If

                ''Add Attachment
                If Attachment <> "" Then
                    Message = Message.Replace("{ATTACHMENT}", Attachment)
                Else
                    Message = Message.Replace("{ATTACHMENT}", "")
                End If

                Mail.Body = Message

                Dim strHTMLCheck As String = UCase(Message)
                Mail.IsBodyHtml = strHTMLCheck.ToUpper.Contains("<BODY") Or strHTMLCheck.ToUpper.Contains("<TABLE") Or strHTMLCheck.ToUpper.Contains("<DIV") Or strHTMLCheck.ToUpper.Contains("<BR") Or strHTMLCheck.ToUpper.Contains("<P") Or strHTMLCheck.ToUpper.Contains("<SPAN")
                Dim SMTPServer As New System.Net.Mail.SmtpClient()

                SMTPServer.Timeout = 100000
                SMTPServer.Host = g_EmailHostIPAddress
                SMTPServer.Port = g_EmailPort
                SMTPServer.EnableSsl = False
                ''SMTPServer.Credentials = New System.Net.NetworkCredential(g_EmailUserName, g_EmailPassword)
                Debug.Print(ToAddress & " - " & Subject)
                If g_EmailEnabled Then
                    SMTPServer.Send(Mail)
                End If
            End If
        End If
    End Sub

    Function GetNextDate(ByVal d As DayOfWeek, Optional ByVal StartDate As Date = Nothing) As Date
        If StartDate = DateTime.MinValue Then
            StartDate = Now
            For p As Integer = 1 To 7
                If StartDate.AddDays(p).DayOfWeek = d Then Return StartDate.AddDays(p)
            Next
        Else
            Return Date.Now
        End If
    End Function

    Public Function validateUser(ByVal USERNAME As String, ByVal PASSWORD As String)
        Dim User As String = ""
        Dim strSql As String = ""
        Dim tblResults As DataTable = Nothing

        ''''''''''VERIFY USER''''''''''''
        If IsNothing(System.Web.HttpContext.Current.Session("User_Name")) Then
            strSql = "Select * from dbo.SYS_USERS where Username = '" & USERNAME & "' and password = '" & PASSWORD & "'"
            tblResults = g_IO_Execute_SQL(strSql, False)

            If tblResults.Rows.Count > 0 Then
                System.Web.HttpContext.Current.Session("USER_NAME") = tblResults.Rows(0)("Username").ToString()
                System.Web.HttpContext.Current.Session("EMAIL") = tblResults.Rows(0)("Email").ToString()
                System.Web.HttpContext.Current.Session("IS_COACH") = tblResults.Rows(0)("IsCoach").ToString()

                strSql = "Insert into dbo.AUDIT (AUDIT_TYPE, AUDIT_MESSAGE, AUDIT_USER, AUDIT_DETAILS) VALUES (" &
                    "'USER AUTHENTICATION','AUTHENTICATION SUCCESS','" & USERNAME & "','LOGIN FAILURE. TIME: " & Date.Now.Hour & ":" & Date.Now.Minute & ":" & Date.Now.Second & "')"
                'g_IO_Execute_SQL(strSql, False)
                System.Web.HttpContext.Current.Session("isLoggedIn") = "true"
                Return True
            Else
                strSql = "Insert into dbo.AUDIT (AUDIT_TYPE, AUDIT_MESSAGE, AUDIT_USER, AUDIT_DETAILS) VALUES (" &
                    "'USER AUTHENTICATION','AUTHENTICATION FAILED','" & USERNAME & "','LOGIN FAILURE. TIME: " & Date.Now.Hour & ":" & Date.Now.Minute & ":" & Date.Now.Second & "')"
                'g_IO_Execute_SQL(strSql, False)
                Return False
            End If
        Else
            Return True
            Exit Function
        End If
    End Function

    Public Function validateAdmin()
        If System.Web.HttpContext.Current.Session("USER_NAME").ToString.ToUpper = "IS_COACH" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function validateAOR(ByVal Area As String)
        If System.Web.HttpContext.Current.Session("USER_NAME").ToString.ToUpper = "IS_COACH" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub createDefaultRow(ByRef tbl As DataTable, ByVal strDefault As String, ByVal strTextField As String, ByVal strValueField As String)
        ''Create the default
        Dim defaultRow As DataRow = tbl.NewRow
        defaultRow(strTextField) = strDefault
        defaultRow(strValueField) = "-1"
        tbl.Rows.InsertAt(defaultRow, 0)
    End Sub

    Public Function createCSVLink(ByVal URL As String)
        Dim strResponse As String = "<div><img src=""images/download-csv.png"" style=""cursor: pointer"" onclick=""redirect('" & URL & "')"" /></div>"
        Return strResponse
    End Function

    Public Sub create_csv_from_datatable(ByVal table As DataTable, ByVal strFileName As String)
        Dim myCsv As New StringBuilder
        Dim strBuilder As String = ""
        Dim delimiter As String = ""
        ''Build the header
        For Each col In table.Columns
            strBuilder &= delimiter & col.ToString.ToUpper
            delimiter = ","
        Next
        ''Append the column header to the csv
        myCsv.AppendLine(strBuilder)

        For Each row In table.Rows
            delimiter = ""
            strBuilder = ""
            For Each col In table.Columns
                strBuilder &= delimiter & """" & IIf(IsDBNull(row(col)), "", row(col).ToString.ToUpper) & """"
                delimiter = ","
            Next
            myCsv.AppendLine(strBuilder)
        Next

        'foreach(myRow r in myRows) {
        '  myCsv.AppendFormat("\"{0}\",{1}", r.MyCol1, r.MyCol2);
        '}

        System.Web.HttpContext.Current.Response.ContentType = "application/csv"
        System.Web.HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=" & strFileName)
        System.Web.HttpContext.Current.Response.Write(myCsv.ToString())
        System.Web.HttpContext.Current.Response.End()

    End Sub

    Public Function createHeaderWithSorts(ByVal headerList As String, ByVal redirectPage As String, ByVal QueryString As NameValueCollection)
        Dim strResponse As String = ""
        For Each col In headerList.Split(";")
            ''rpt_GetReportDetails.aspx?Report=LOC_W_USR
            If Not IsNothing(QueryString("sort")) Then
                ''It's been sorted already  Check the sort order
                If QueryString("sort").ToString.ToUpper = col.Split("~")(1).ToUpper AndAlso QueryString("order").ToString.ToUpper = "ASC" Then
                    strResponse &= "<th style=""cursor: pointer;"" onclick=""redirect('" & redirectPage & "&sort=" & col.Split("~")(1) & "&order=desc')"">" & col.Split("~")(0) & "</th>"
                Else
                    strResponse &= "<th style=""cursor: pointer;"" onclick=""redirect('" & redirectPage & "&sort=" & col.Split("~")(1) & "&order=asc')"">" & col.Split("~")(0) & "</th>"
                End If
            Else
                ''It hasn't been sorted, so set the defaults...
                strResponse &= "<th style=""cursor: pointer;"" onclick=""redirect('" & redirectPage & "&sort=" & col.Split("~")(1) & "&order=asc')"">" & col.Split("~")(0) & "</th>"
            End If
        Next
        Return strResponse
    End Function

    Public Function createSortedTblView(ByVal tblToSort As DataTable, ByVal colToSort As String, ByVal sortOrder As String)
        Dim dataView As New DataView(tblToSort)

        If Not IsNothing(colToSort) Then
            ''Sort in the order given
            dataView.Sort = colToSort & " " & sortOrder ''" AutoID DESC, Name DESC"
            tblToSort = dataView.ToTable()
        End If

        Return tblToSort

    End Function

    Public Function returnQueryString(ByVal queryString As NameValueCollection)
        Dim strVariables As String = ""
        Dim delimiter As String = ""

        For Each item In queryString
            strVariables &= delimiter & item.ToString & "=" & queryString(item)
            delimiter = "&"
        Next

        Return strVariables
    End Function

    Public Sub logActivity(ByVal queryString As NameValueCollection, ByVal PAGE As String, ByVal Action As String)
        If queryString.Count > 0 Then
            PAGE &= "?" & returnQueryString(queryString)
        End If

        Dim userID As String = System.Web.HttpContext.Current.Session("USER_ID")
        Dim strSQL As String = "udp_LogActivity @USERID = '" & userID.Replace("'", "''") & "', @PAGE = '" & PAGE.Replace("'", "''") & "', @ACTION = '" & Action.Replace("'", "''") & "'"
        g_IO_Execute_SQL(strSQL, False)
    End Sub

    Public Sub decrementProductInventory(ByVal productID As Integer, ByVal qty As Integer)
        Dim strSQL As String = "Update inventory.vw_Products set ONHAND_QTY = (ONHAND_QTY - " & qty & ") where product_id = " & productID
        g_IO_Execute_SQL(strSQL, False)
    End Sub

    Public Sub incrementProductInventory(ByVal productID As Integer, ByVal qty As Integer)
        Dim strSQL As String = "Update inventory.vw_Products set ONHAND_QTY = (ONHAND_QTY + " & qty & ") where product_id = " & productID
        g_IO_Execute_SQL(strSQL, False)
    End Sub
End Module