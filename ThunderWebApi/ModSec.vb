Imports System.Net.Mail
Imports System.Configuration
Imports System.Web.Services
Imports System.Data.SqlClient
Imports System.IO
Imports System.DirectoryServices
Imports System.Security.Principal
Imports System.DirectoryServices.AccountManagement

''    Ensure the following lines are in the web.config file

''    <add key = "secGroupName" value="PORTAL_INFRASTRUCTURE"/>
''    <add key = "auditLoginActivity" value="true" />


Module ModSec

    Public g_SecurityGroupName As String = ConfigurationManager.AppSettings("secGroupName")
    Public g_AuditLoginActivity As String = ConfigurationManager.AppSettings("auditLoginActivity")
    Dim _path As String
    Dim _filterAttribute As String

    Public Function validateAdmin()
        If IsNothing(System.Web.HttpContext.Current.Session("IS_ADMIN")) Then
            Return False
        End If

        If System.Web.HttpContext.Current.Session("IS_ADMIN") = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function validateEndUser()
        If IsNothing(System.Web.HttpContext.Current.Session("END_USER")) Then
            Return False
        End If
    End Function

    Public Function validateStockClerk()
        If IsNothing(System.Web.HttpContext.Current.Session("STOCKCLERK")) Then
            Return False
        End If
    End Function

    Public Function validateDirector()
        If IsNothing(System.Web.HttpContext.Current.Session("APPROVER")) Then
            Return False
        End If
    End Function


    Public Sub logActivity(ByVal queryString As NameValueCollection, ByVal PAGE As String, ByVal Action As String)
        If queryString.Count > 0 Then
            PAGE &= "?" & returnQueryString(queryString)
        End If

        Dim userID As String = System.Web.HttpContext.Current.Session("USER_ID")
        Dim strSQL As String = "udp_LogActivity @USERID = '" & userID.Replace("'", "''") & "', @PAGE = '" & PAGE.Replace("'", "''") & "', @ACTION = '" & Action.Replace("'", "''") & "'"
        g_IO_Execute_SQL(strSQL, False)
    End Sub

    Public Function returnQueryString(ByVal queryString As NameValueCollection)
        Dim strVariables As String = ""
        Dim delimiter As String = ""

        For Each item In queryString
            strVariables &= delimiter & item.ToString & "=" & queryString(item)
            delimiter = "&"
        Next

        Return strVariables
    End Function

    Public Function IsAuthenticated(ByVal domain As String, ByVal username As String, ByVal pwd As String) As Integer
        ''Return 0 if successful
        ''Return 1 if bad password
        Dim domainAndUsername As String = domain & "\" & username
        _path = ""
        Dim entry As DirectoryEntry = New DirectoryEntry(_path, domainAndUsername, pwd)

        Try
            System.Web.HttpContext.Current.Session("IS_ADMIN") = 0
            System.Web.HttpContext.Current.Session("END_USER") = 0
            System.Web.HttpContext.Current.Session("STOCKCLERK") = 0
            System.Web.HttpContext.Current.Session("APPROVER") = 0

            'Bind to the native AdsObject to force authentication.			
            Dim obj As Object = entry.NativeObject
            Dim search As DirectorySearcher = New DirectorySearcher(entry)

            search.Filter = "(SAMAccountName=" & username & ")"
            search.PropertiesToLoad.Add("cn")
            search.PropertiesToLoad.Add("mail")
            search.PropertiesToLoad.Add("givenName")
            search.PropertiesToLoad.Add("sn")
            Dim result As SearchResult = search.FindOne()

            If (result Is Nothing) Then
                Return False
            End If

            'Update the new path to the user in the directory.
            _path = result.Path
            _filterAttribute = CType(result.Properties("cn")(0), String)
            System.Web.HttpContext.Current.Session("USER_ID") = username.ToUpper
            System.Web.HttpContext.Current.Session("USER_EMAIL") = CType(result.Properties("mail")(0), String).ToUpper
            System.Web.HttpContext.Current.Session("USER_NAME") = CType(result.Properties("givenName")(0), String).ToUpper & " " & CType(result.Properties("sn")(0), String).ToUpper

            ''Get the Users Groups
            Dim UserGroups As String = GetGroups()
            System.Web.HttpContext.Current.Session("MEMBER_OF") = UserGroups
            Dim is_readonly As Integer = 0
            Dim is_admin As Integer = 0
            Dim is_enduser As Integer = 0

            ''Add the users
            If UserGroups.ToUpper.IndexOf(g_SecurityGroupName + "_ADMIN") >= 0 Then
                System.Web.HttpContext.Current.Session("IS_ADMIN") = 1
            Else
                System.Web.HttpContext.Current.Session("IS_ADMIN") = 0
            End If

            If UserGroups.ToUpper.IndexOf(g_SecurityGroupName + "_ENDUSER") >= 0 Then
                System.Web.HttpContext.Current.Session("END_USER") = 1
            Else
                System.Web.HttpContext.Current.Session("END_USER") = 0
            End If

            If UserGroups.ToUpper.IndexOf(g_SecurityGroupName + "_STOCKCLERK") >= 0 Then
                System.Web.HttpContext.Current.Session("STOCKCLERK") = 1
            Else
                System.Web.HttpContext.Current.Session("STOCKCLERK") = 0
            End If

            If UserGroups.ToUpper.IndexOf(g_SecurityGroupName + "_APPROVER") >= 0 Then
                System.Web.HttpContext.Current.Session("APPROVER") = 1
            Else
                System.Web.HttpContext.Current.Session("APPROVER") = 0
            End If

            If g_AuditLoginActivity Then
                Dim strSql As String = "Insert into dbo.AUDIT 
                   (AUDIT_TYPE, AUDIT_MESSAGE, AUDIT_USER, AUDIT_DETAILS, IS_ADMIN, IS_ENDUSER, IS_STOCKCLERK, IS_APPROVER) VALUES (" &
                   "'USER AUTHENTICATED'
                   ,'USER GROUP LISTING'
                   ,'" & System.Web.HttpContext.Current.Session("USER_ID").ToString & "'
                   ,'" & System.Web.HttpContext.Current.Session("MEMBER_OF").ToString.ToUpper.Replace("|", " | ") & "'
                   ,'" & System.Web.HttpContext.Current.Session("IS_ADMIN").ToString & "'
                   ,'" & System.Web.HttpContext.Current.Session("END_USER").ToString & "'
                   ,'" & System.Web.HttpContext.Current.Session("STOCKCLERK").ToString & "'
                   ,'" & System.Web.HttpContext.Current.Session("APPROVER").ToString & "')"
                g_IO_Execute_SQL(strSql, False)
            End If

            System.Web.HttpContext.Current.Session("isLoggedIn") = "true"

        Catch ex As Exception
            Dim messageString = ex.Message.ToString
            If messageString.ToUpper.IndexOf("BAD PASSWORD") >= 0 Then
                Return 1
            End If
        End Try
        Return 0
    End Function

    Public Sub IsAuthenticatedNoLogon()
        ''Return 1 if bad password
        Dim User As String = HttpContext.Current.Request.LogonUserIdentity.Name.ToString.Replace("D700\", "").Replace("ITGBRANDS\", "")
        System.Web.HttpContext.Current.Session("USER_ID") = User.ToUpper

        Try

            System.Web.HttpContext.Current.Session("END_USER") = 0
            System.Web.HttpContext.Current.Session("IS_ADMIN") = 0
            System.Web.HttpContext.Current.Session("STOCKCLERK") = 0
            System.Web.HttpContext.Current.Session("APPROVER") = 0

            ''Get the Users Groups
            Dim UserGroups As String = GetGroups()
            System.Web.HttpContext.Current.Session("MEMBER_OF") = UserGroups
            Dim is_readonly As Integer = 0
            Dim is_admin As Integer = 0
            Dim is_enduser As Integer = 0

            ''Add the users
            If UserGroups.ToUpper.IndexOf(g_SecurityGroupName + "_ADMIN") >= 0 Then
                System.Web.HttpContext.Current.Session("IS_ADMIN") = 1
            Else
                System.Web.HttpContext.Current.Session("IS_ADMIN") = 0
            End If

            If UserGroups.ToUpper.IndexOf(g_SecurityGroupName + "_ENDUSER") >= 0 Then
                System.Web.HttpContext.Current.Session("END_USER") = 1
            Else
                System.Web.HttpContext.Current.Session("END_USER") = 0
            End If

            If UserGroups.ToUpper.IndexOf(g_SecurityGroupName + "_STOCKCLERK") >= 0 Then
                System.Web.HttpContext.Current.Session("STOCKCLERK") = 1
            Else
                System.Web.HttpContext.Current.Session("STOCKCLERK") = 0
            End If

            If UserGroups.ToUpper.IndexOf(g_SecurityGroupName + "_APPROVER") >= 0 Then
                System.Web.HttpContext.Current.Session("APPROVER") = 1
            Else
                System.Web.HttpContext.Current.Session("APPROVER") = 0
            End If

            If g_AuditLoginActivity Then
                Dim strSql As String = "Insert into dbo.AUDIT 
                   (AUDIT_TYPE, AUDIT_MESSAGE, AUDIT_USER, AUDIT_DETAILS, IS_ADMIN, IS_ENDUSER, IS_STOCKCLERK, IS_APPROVER) VALUES (" &
                   "'USER AUTHENTICATED'
                   ,'USER GROUP LISTING'
                   ,'" & System.Web.HttpContext.Current.Session("USER_ID").ToString & "'
                   ,'" & System.Web.HttpContext.Current.Session("MEMBER_OF").ToString.ToUpper.Replace("|", " | ") & "'
                   ,'" & System.Web.HttpContext.Current.Session("IS_ADMIN").ToString & "'
                   ,'" & System.Web.HttpContext.Current.Session("END_USER").ToString & "'
                   ,'" & System.Web.HttpContext.Current.Session("STOCKCLERK").ToString & "'
                   ,'" & System.Web.HttpContext.Current.Session("APPROVER").ToString & "')"
                g_IO_Execute_SQL(strSql, False)
            End If

            System.Web.HttpContext.Current.Session("isLoggedIn") = "true"

        Catch ex As Exception
            Dim messageString = ex.Message.ToString
        End Try
    End Sub

    Public Function GetGroups() As String
        Dim search As DirectorySearcher = New DirectorySearcher(_path.Trim)
        search.Filter = "(cn=" & _filterAttribute.Trim & ")"
        search.PropertiesToLoad.Add("memberOf")
        Dim groupNames As StringBuilder = New StringBuilder()

        Try
            Dim result As SearchResult = search.FindOne()
            Dim propertyCount As Integer = result.Properties("memberOf").Count

            Dim dn As String
            Dim equalsIndex, commaIndex

            Dim propertyCounter As Integer

            For propertyCounter = 0 To propertyCount - 1
                dn = CType(result.Properties("memberOf")(propertyCounter), String)

                equalsIndex = dn.IndexOf("=", 1)
                commaIndex = dn.IndexOf(",", 1)
                If (equalsIndex = -1) Then
                    Return Nothing
                End If

                groupNames.Append(dn.Substring((equalsIndex + 1), (commaIndex - equalsIndex) - 1))
                groupNames.Append("|")
            Next

        Catch ex As Exception
            Throw New Exception("Error obtaining group names. " & ex.Message)
        End Try

        Return groupNames.ToString()
    End Function

    'Public Function GetEmail(ByVal domain As String, ByVal username As String, ByVal pwd As String) As String
    '    Dim userEmail As String = String.Empty

    '    Using search As DirectorySearcher = New DirectorySearcher(_path)
    '        search.Filter = "(SAMAccountName=" & username & ")"
    '        search.PropertiesToLoad.Add("mail")
    '        Try
    '            Dim result As SearchResult = search.FindOne()
    '            If result IsNot Nothing Then userEmail = result.Properties("mail").ToString
    '        Catch ex As Exception

    '        End Try
    '    End Using

    '    Return userEmail
    'End Function
End Module
