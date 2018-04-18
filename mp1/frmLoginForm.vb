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
        Dim vTokenResult As Boolean
        Dim user As String
        Dim password As String

        user = UsernameTextBox.Text
        password = PasswordTextBox.Text

        Dim clsAuthen As clsAuthentication = New clsAuthentication(
            "http://127.0.0.1:8000", user, password)


        vTokenResult = clsAuthen.requestToken()
        lblStatus.Text = clsAuthen.Message
        lblStatus.ForeColor = IIf(vTokenResult, Color.Green, Color.Red)

        'Assign public variable (MUST -- Very Importance)
        user_id = clsAuthen.user_id_token
        access_token = clsAuthen.access_token
        refresh_token = clsAuthen.requestToken
        '------------------------

        If vTokenResult Then
            frmParameter.Show()
            'Me.Close()
        End If
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub UsernameTextBox_TextChanged(sender As Object, e As EventArgs) Handles UsernameTextBox.TextChanged
        lblStatus.Text = ""
    End Sub

    Private Sub PasswordTextBox_TextChanged(sender As Object, e As EventArgs) Handles PasswordTextBox.TextChanged
        lblStatus.Text = ""
    End Sub
End Class
