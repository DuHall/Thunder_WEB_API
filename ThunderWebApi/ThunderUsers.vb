Imports System.Data.SqlClient
Imports System.Web.Configuration.WebConfigurationManager

Public Class ThunderUsers
    Public Property Username As String
    Public Property Firstname As String
    Public Property Lastname As String


    Public Sub New()

    End Sub
    Public Sub New(Username As String, Firstname As String, Lastname As String)
        Me.Username = Username
        Me.Firstname = Firstname
        Me.Lastname = Lastname

    End Sub

    Public Function ConvertReader(ByVal reader As SqlDataReader)
        Dim users As List(Of ThunderUsers) = New List(Of ThunderUsers)
        If reader.HasRows Then
            While reader.Read()
                users.Add(New ThunderUsers(reader.Item("Username").ToString(),
                                         reader.Item("Firstname").ToString(),
                                         reader.Item("Lastname").ToString()))
            End While
        End If
        Return users
    End Function


    Public Function GetUsers() As IEnumerable(Of ThunderUsers)
        Dim cn As SqlConnection =
            New SqlConnection(ConnectionStrings("ThunderConnect").ToString)
        Dim cmd As New SqlCommand("[dbo].[SP_GET_USERS]", cn)
        cmd.CommandType = CommandType.StoredProcedure
        cn.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader()
        Dim users As List(Of ThunderUsers) = ConvertReader(reader)
        cn.Close()

        Return users
    End Function




End Class
