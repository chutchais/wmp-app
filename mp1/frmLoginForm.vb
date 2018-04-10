Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Web.Script.Serialization

Public Class frmLoginForm

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See https://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        'Dim resp As String
        'Dim client As HttpClient
        'client = New HttpClient()
        'client.BaseAddress = New Uri("http://127.0.0.1:8000/api/token")
        'client.Timeout = New TimeSpan(0, 0, 60)

        'Dim _ContentType As String
        '_ContentType = "application/json"
        'client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue(_ContentType))


        Dim user As String
        Dim passt As String
        user = UsernameTextBox.Text
        passt = PasswordTextBox.Text

        Dim request As WebRequest = WebRequest.Create("http://127.0.0.1:8000/api/token/")
        request.Method = "POST"
        Dim postData As String
        postData = "username=" & user & "&password=" & passt
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
            lblStatus.Text = "Login Successful...." : lblStatus.ForeColor = Color.Green
        Catch ex As Exception
            lblStatus.Text = "Login failed :No active account found with the given credentials" : lblStatus.ForeColor = Color.Red
        End Try

        'Me.Close()
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

End Class
