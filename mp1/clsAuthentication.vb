Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Web.Script.Serialization
Imports System.IdentityModel.Tokens.Jwt
Imports System.IdentityModel.Tokens
Imports System.Security.Claims
Imports System.DateTime


Public Class clsAuthentication
    Public Property Url As String 'http://127.0.0.1:8000
    Public Property Username As String
    Public Property Password As String

    Private TokenService As String = ""
    Private vMessage As String = ""
    Private vAccess_Token As String = ""
    Private vRefresh_Token As String = ""
    Private vUser_id As String = ""
    Private vExp As Integer = 0

    Public Sub New(ByVal url As String,
               ByVal username As String,
               ByVal password As String)
        Me.Url = url
        Root_url = url

        Me.Username = username
        Me.Password = password

        Me.TokenService = url + "/api/token/"
    End Sub

    Public Property Message() As String
        Get
            Return vMessage
        End Get
        Set(ByVal Value As String)
            Me.vMessage = Value
        End Set
    End Property


    Public Property user_id_token() As String
        Get
            Return vUser_id
        End Get
        Set(ByVal Value As String)
            Me.vUser_id = Value
        End Set
    End Property

    Public Property access_token() As String
        Get
            Return vAccess_Token
        End Get
        Set(ByVal Value As String)
            Me.vAccess_Token = Value
        End Set
    End Property

    Public Property refresh_token() As String
        Get
            Return vRefresh_Token
        End Get
        Set(ByVal Value As String)
            Me.vRefresh_Token = Value
        End Set
    End Property

    Public Property exp() As Integer
        Get
            Return vExp
        End Get
        Set(ByVal Value As Integer)
            Me.vExp = Value
        End Set
    End Property

    Public Function requestToken() As Boolean
        Dim request As WebRequest = WebRequest.Create(Me.TokenService)
        request.Method = "POST"
        Dim postData As String
        postData = "username=" & Username & "&password=" & Password
        Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postData)
        request.ContentType = "application/x-www-form-urlencoded"
        request.ContentLength = byteArray.Length
        Dim dataStream As Stream = request.GetRequestStream()
        dataStream.Write(byteArray, 0, byteArray.Length)
        dataStream.Close()
        Dim response As WebResponse
        Try
            response = request.GetResponse()
            'Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)
            dataStream = response.GetResponseStream()
            Dim reader As New StreamReader(dataStream)
            Dim responseFromServer As String = reader.ReadToEnd()
            reader.Close()
            dataStream.Close()
            response.Close()
            Dim jss = New JavaScriptSerializer()
            Dim data As Object = jss.DeserializeObject(responseFromServer)
            vMessage = "Login Successful...."
            'Assign Token to global variable



            For Each item As KeyValuePair(Of String, Object) In data
                Console.WriteLine(item.Key & ": " & item.Value)
                If item.Key = "refresh" Then
                    vRefresh_Token = item.Value
                End If

                If item.Key = "access" Then
                    vAccess_Token = item.Value
                End If
            Next

            Dim Tokens As New JwtSecurityToken
            Dim handler As New JwtSecurityTokenHandler

            Tokens = handler.ReadToken(vAccess_Token)
            vUser_id = Tokens.Payload.Item("user_id").ToString
            vExp = Tokens.Payload.Exp   'Payload.Item("exp").ToString

            Return True
        Catch ex As Exception
            vMessage = "Login failed :No active account found with the given credentials"
            vRefresh_Token = ""
            vAccess_Token = ""
            Return False
        End Try
    End Function


End Class
