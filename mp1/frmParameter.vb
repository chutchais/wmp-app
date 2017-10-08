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
            .Name = "Param3"
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

        Dim L As New ucParamList()
        With L
            .Name = "Param4"
            .title = "Fiber Type"
            .message = "Testing Mesasge"
            .value = "ABC"
            .regExpress = "[0-9]{3}"
            .slug = "choice1-none"
            .url = "http://127.0.0.1:8000"
            .CurrentForm = Me
            .Location = New Point(50, vPosBottom)
            vPosBottom = .Location.Y + .Height + 5
            .Left = X.Left
            .Width = L.Width
            '.Dock = DockStyle.Bottom
        End With
        Me.Controls.Add(L)
        L.Show()
        L.showOpject()


        Dim R As New ucParamRadio()
        With R
            .Name = "Param5"
            .title = "Fiber Type"
            .message = "Testing Mesasge"
            .value = "ABC"
            .regExpress = "[0-9]{3}"
            .slug = "choice1-none"
            .url = "http://127.0.0.1:8000"
            .CurrentForm = Me
            .Location = New Point(50, vPosBottom)
            vPosBottom = .Location.Y + .Height + 5
            .Left = X.Left
            .Width = L.Width
            '.Dock = DockStyle.Bottom
        End With
        Me.Controls.Add(R)
        R.Show()
        R.showOpject()



    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) 

    End Sub

    Private Sub TextBox1_Validated(sender As Object, e As EventArgs) 

    End Sub

    Private Sub TextBox1_Validating(sender As Object, e As CancelEventArgs) 
        'If Not IsValid(TextBox1.Text) Then
        '    e.Cancel = True
        '    TextBox1.Select(0, TextBox1.Text.Length)
        'End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) 

    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) 

    End Sub

    Function IsValid(ByRef value As String) As Boolean
        Return Regex.IsMatch(value, "[0-9]{3}")
        'Dim myRegEx As New Regex("[0-9]{3}", RegexOptions.None)
        'Dim myMatch As Match
        'myMatch = myRegEx.Match(value)
        'Return myMatch.Success
    End Function

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) 

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) 

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

    End Sub

    Private Sub RadioButton1_Click(sender As Object, e As EventArgs) Handles RadioButton1.Click

    End Sub



    'Private Function IsValid(ByVal Contents As String) As Boolean
    '    Dim myRegEx As New Regex("*", RegexOptions.IgnoreCase)
    '    Dim myMatch As Match
    '    myMatch = myRegEx.Match(Contents)
    '    Return myMatch.Success
    'End Function

End Class