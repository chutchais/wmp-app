Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Web.Script.Serialization

Public Class clsAuthentication
    Public Property Url As String 'http://127.0.0.1:8000
    Public Property Username As String
    Public Property Password As String

    Private TokenService As String = ""
    Private vMessage As String = ""

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
            Return Me.vMessage
        End Get
        Set(ByVal Value As String)
            Me.vMessage = Value
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
                    Token_refresh = item.Value
                End If

                If item.Key = "access" Then
                    Token_access = item.Value
                End If
            Next

            Return True
        Catch ex As Exception
            vMessage = "Login failed :No active account found with the given credentials"
            Token_refresh = ""
            Token_access = ""
            Return False
        End Try
    End Function


End Class
