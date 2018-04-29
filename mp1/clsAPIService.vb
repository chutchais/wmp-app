Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Web.Script.Serialization

Public Class clsAPIService
    Private vUrl As String
    Private vAccessToken As String

    Public Property Url() As String
        Get
            Return vUrl
        End Get
        Set(ByVal Value As String)
            Me.vUrl = Value
        End Set
    End Property

    Public Property access_token() As String
        Get
            Return vAccessToken
        End Get
        Set(ByVal Value As String)
            Me.vAccessToken = Value
        End Set
    End Property

    Public Function getJsonObject(ByVal address As String) As Object
        'Support request with Token
        Dim request As WebRequest = WebRequest.Create(address)
        request.Method = "GET"
        Dim byteArray As Byte() = Encoding.UTF8.GetBytes("")
        request.PreAuthenticate = True
        request.Headers.Add("Authorization", "Bearer " + vAccessToken)


        Dim myHttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
        Dim myWebSource As New StreamReader(myHttpWebResponse.GetResponseStream())
        Dim myPageSource As String = myWebSource.ReadToEnd()

        Dim jss = New JavaScriptSerializer()
        Dim data As Object = jss.Deserialize(Of Object)(myPageSource)
        Return data


    End Function

    Public Function getObjectByUrl(vUrl As String) As Object
        Dim json As Object
        json = getJsonObject(vUrl)
        getObjectByUrl = json
    End Function

    Public Function getObjectBySlug(vApp As String, vSlug As String) As Object
        Dim json As Object
        Dim vUrlSlug As String
        vUrlSlug = vUrl + "/api/" + vApp + "/" + vSlug & "/"
        json = getJsonObject(vUrlSlug)
        getObjectBySlug = json
    End Function


    Public Function getRoutingDetail(vRoute As String, vOperation As String) As Object
        'VRoute = Routing name
        'vOperation = Operation name
        Dim json As Object
        json = getJsonObject(vUrl + "/api/routing-detail/?route=" + vRoute + "&operation=" & vOperation)
        Return json
    End Function
End Class
