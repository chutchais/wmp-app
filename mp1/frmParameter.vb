Imports System.ComponentModel
Imports System.Text.RegularExpressions

<Runtime.InteropServices.ComVisible(True)>
Public Class frmParameter
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim vPosBottom As Integer
        Dim X As New ucParaText()
        With X
            .Name = "Param1"
            .title = "Param Name"
            .message = "Testing Mesasge"
            .value = "ABC"
            .regExpress = "[0-9]{3}"
            .slug = "item2-cisco-test-product"
            .url = "http://127.0.0.1:8000"
            .CurrentForm = Me
            .Location = New Point(50, 10)
            vPosBottom = .Location.Y + .Height + 5
        End With
        Me.Controls.Add(X)
        X.Show()
        X.showOpject()

        Dim Y As New ucParaText()
        With Y
            .Name = "Param2"
            .title = "Param Name"
            .message = "Testing Mesasge"
            .value = "ABC"
            .regExpress = "[0-9]{3}"
            .slug = "item2-cisco-test-product"
            .url = "http://127.0.0.1:8000"
            .CurrentForm = Me
            .Location = New Point(50, vPosBottom)
            vPosBottom = .Location.Y + .Height + 5
            '.Dock = DockStyle.Bottom
        End With
        Me.Controls.Add(Y)
        Y.Show()
        Y.showOpject()

        Dim Z As New ucParaText()
        With Z
            .Name = "Param2"
            .title = "Param Name"
            .message = "Testing Mesasge"
            .value = "ABC"
            .regExpress = "[0-9]{3}"
            .slug = "item2-cisco-test-product"
            .url = "http://127.0.0.1:8000"
            .CurrentForm = Me
            .Location = New Point(50, vPosBottom)
            vPosBottom = .Location.Y + .Height + 5
            '.Dock = DockStyle.Bottom
        End With
        Me.Controls.Add(Z)
        Z.Show()
        Z.showOpject()

        'For i = 1 To 2

        '    Dim txt As New TextBox
        '    With txt
        '        .Text = "name" & i.ToString()
        '        .Location = New Point(50, (i * 100) + 10)
        '        .BackColor = Color.Red
        '        .Size = New Size(50, 20)
        '    End With
        '    Me.Controls.Add(txt)
        '    txt.Show()


        'Next

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_Validated(sender As Object, e As EventArgs) Handles TextBox1.Validated

    End Sub

    Private Sub TextBox1_Validating(sender As Object, e As CancelEventArgs) Handles TextBox1.Validating
        'If Not IsValid(TextBox1.Text) Then
        '    e.Cancel = True
        '    TextBox1.Select(0, TextBox1.Text.Length)
        'End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress

    End Sub

    Function IsValid(ByRef value As String) As Boolean
        Return Regex.IsMatch(value, "[0-9]{3}")
        'Dim myRegEx As New Regex("[0-9]{3}", RegexOptions.None)
        'Dim myMatch As Match
        'myMatch = myRegEx.Match(value)
        'Return myMatch.Success
    End Function

    'Private Function IsValid(ByVal Contents As String) As Boolean
    '    Dim myRegEx As New Regex("*", RegexOptions.IgnoreCase)
    '    Dim myMatch As Match
    '    myMatch = myRegEx.Match(Contents)
    '    Return myMatch.Success
    'End Function

End Class