Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.Web.Script.Serialization
Imports System.Dynamic

<Runtime.InteropServices.ComVisible(True)>
Public Class frmParameter
    Public Shared Function getJsonString(ByVal address As String) As String

        Dim client As WebClient = New WebClient()
        Dim reply As String = client.DownloadString(address)
        Return reply
    End Function

    Public Shared Function getJsonObject(ByVal address As String) As Object
        Dim client As WebClient = New WebClient()
        Dim json As String = client.DownloadString(address)
        Dim jss = New JavaScriptSerializer()
        Dim data As Object = jss.Deserialize(Of Object)(json)
        Return data
    End Function

    Private Function getParameterBySlug(vRoutingDetailSlug As String) As Object
        Dim json As Object
        json = getJsonObject(vUrl + "/api/routing-detail/" + vRoutingDetailSlug)
        getParameterBySlug = json
    End Function

    Private Function getItemBySlug(vParameterSlug As String) As Object
        Dim json As Object
        json = getJsonObject(vUrl + "/api/parameter/" + vParameterSlug)
        getItemBySlug = json
    End Function

    'Private Function getSnippetBySlug(vSnippetSlug As String) As Object
    '    Dim json As Object
    '    json = getJsonObject(vUrl + "/api/snippet/" + vSnippetSlug)
    '    Return json
    'End Function

    Private vUrl As String = "http://127.0.0.1:8000"
    Private ixRadio As Integer = 0

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CreateObject()
    End Sub

    Sub CreateObject()
        Dim objParams As Object
        objParams = getParameterBySlug("a001-visual-inspection-routing1")

        Dim tabControl As New TabControl
        'Dim tabPage As New TabPage
        With tabControl
            .Name = "Parameter"
            .Location = New Point(0, 130)
            .Size = New Size(Me.Width, 300)
            .SizeMode = TabSizeMode.Normal
            .AutoSize = True
        End With

        Dim objParam As Object
        Dim vParamSlug As String
        For Each objParam In objParams("parameter")
            Dim tabPage As New TabPage
            tabPage.Text = IIf(IsDBNull(objParam("title")), objParam("name"), objParam("title"))
            tabPage.Name = objParam("name")
            tabPage.AutoScroll = True
            tabPage.AutoSize = True
            ' AddHandler Text.Validating, AddressOf text_Validating
            tabControl.TabPages.Add(tabPage)
            ''------Add Parameter to Page---
            vParamSlug = objParam("name")
            addParameterToPage(vParamSlug, tabPage)
            ''------------------------------
        Next
        tabControl.Show()
        Me.Controls.Add(tabControl)
    End Sub


    Sub addParameterToPage(vParameterSlug As String, page As TabPage)
        Dim objItems As Object
        Dim objItem As Object
        objItems = getItemBySlug(vParameterSlug)

        Dim vItemName As String
        Dim vTitle As String
        Dim vDefaultValue As String
        Dim vRegExp As String
        Dim vItemSlug As String
        Dim vItemType As String
        Dim vRequired As Boolean

        Dim vPosBottom As Integer = 10
        Dim ucText As New ucParaText
        Dim ucList As New ucParamList
        Dim ucRadio As ucParamRadio

        Dim ucOption As New ucParamOption
        For Each objItem In objItems("items")
            vItemName = objItem("name")
            vTitle = objItem("title")
            vDefaultValue = objItem("default_value")
            vRegExp = objItem("regexp")
            vItemSlug = objItem("slug")
            vItemType = objItem("input_type")
            vRequired = objItem("required")
            Select Case vItemType
                Case "TEXT"
                    ucText = New ucParaText With {
                                .Name = vItemName,
                                .title = vTitle,
                                .message = "Testing Mesasge",
                                .value = vDefaultValue,
                                .regExpress = vRegExp,
                                .slug = vItemSlug,
                                .url = vUrl,
                                .CurrentForm = Me,
                                .Location = New Point(50, vPosBottom),
                                .required = vRequired
                                }
                    vPosBottom = ucText.Location.Y + ucText.Height + 5
                    page.Controls.Add(ucText)
                    ucText.Show()
                    ucText.showOpject()
                Case "LIST"
                    ucList = New ucParamList With {
                        .Name = vItemName,
                        .title = vTitle,
                        .message = "Testing Mesasge",
                        .value = vDefaultValue,
                        .regExpress = vRegExp,
                        .slug = vItemSlug,
                        .url = vUrl,
                        .CurrentForm = Me,
                        .Location = New Point(50, vPosBottom),
                        .required = vRequired
                    }
                    vPosBottom = ucList.Location.Y + ucList.Height + 5
                    page.Controls.Add(ucList)
                    ucList.Show()
                    ucList.showOpject()

                Case "RADIO"
                    ucRadio = New ucParamRadio With {
                        .Name = vItemName,
                        .title = vTitle,
                        .message = "",
                        .value = vDefaultValue,
                        .regExpress = vRegExp,
                        .slug = vItemSlug,
                        .url = vUrl,
                        .CurrentForm = Me,
                        .Location = New Point(50, vPosBottom),
                        .required = vRequired
                    }

                    page.Controls.Add(ucRadio)
                    'vcc = ucRadio.Height_New
                    ucRadio.Show()
                    ucRadio.AutoSize = True
                    vPosBottom = ucRadio.Location.Y + ucRadio.Height + 5
                    ucRadio.showOpject()
                Case "OPTION"
                    ucOption = New ucParamOption With {
                        .Name = vItemName,
                        .title = vTitle,
                        .message = "",
                        .value = vDefaultValue,
                        .regExpress = vRegExp,
                        .slug = vItemSlug,
                        .url = vUrl,
                        .CurrentForm = Me,
                        .Location = New Point(50, vPosBottom),
                        .required = vRequired
                    }
                    page.Controls.Add(ucOption)
                    ucOption.Show()
                    ucOption.AutoSize = True
                    vPosBottom = ucOption.Location.Y + ucOption.Height + 5
                    ucOption.showOpject()
                Case "SCRIPT"
            End Select
        Next



        '    Dim X As New ucParaText()
        '    With X
        '        .Name = "Param1"
        '        .title = "Param Name"
        '        .message = "Testing Mesasge"
        '        .value = "ABC"
        '        .regExpress = "[0-9]{3}"
        '        .slug = "item2-cisco-test-product"
        '        .url = "http://127.0.0.1:8000"
        '        .CurrentForm = Me
        '        .Location = New Point(50, 10)
        '        vPosBottom = .Location.Y + .Height + 5
        '    End With
        '    page.Controls.Add(X)
        '    X.Show()
        '    X.showOpject()

        '    Dim Y As New ucParaText()
        '    With Y
        '        .Name = "Param2"
        '        .title = "Param Name"
        '        .message = "Testing Mesasge"
        '        .value = "ABC"
        '        .regExpress = "[0-9]{3}"
        '        .slug = "item2-cisco-test-product"
        '        .url = "http://127.0.0.1:8000"
        '        .CurrentForm = Me
        '        .Location = New Point(50, vPosBottom)
        '        vPosBottom = .Location.Y + .Height + 5
        '        '.Dock = DockStyle.Bottom
        '    End With
        '    page.Controls.Add(Y)
        '    Y.Show()
        '    Y.showOpject()

        '    Dim Z As New ucParaText()
        '    With Z
        '        .Name = "Param3"
        '        .title = "Param Name"
        '        .message = "Testing Mesasge"
        '        .value = "ABC"
        '        .regExpress = "[0-9]{3}"
        '        .slug = "item2-cisco-test-product"
        '        .url = "http://127.0.0.1:8000"
        '        .CurrentForm = Me
        '        .Location = New Point(50, vPosBottom)
        '        vPosBottom = .Location.Y + .Height + 5
        '        '.Dock = DockStyle.Bottom
        '    End With
        '    page.Controls.Add(Z)
        '    Z.Show()
        '    Z.showOpject()

        '    Dim L As New ucParamList()
        '    With L
        '        .Name = "Param4"
        '        .title = "Fiber Type"
        '        .message = "Testing Mesasge"
        '        .value = "ABC"
        '        .regExpress = "[0-9]{3}"
        '        .slug = "choice1-none"
        '        .url = "http://127.0.0.1:8000"
        '        .CurrentForm = Me
        '        .Location = New Point(50, vPosBottom)
        '        vPosBottom = .Location.Y + .Height + 5
        '        .Left = X.Left
        '        .Width = L.Width
        '        '.Dock = DockStyle.Bottom
        '    End With
        '    page.Controls.Add(L)
        '    L.Show()
        '    L.showOpject()


        '    Dim R As New ucParamRadio()
        '    With R
        '        .Name = "Param5"
        '        .title = "Fiber Type"
        '        .message = "Testing Mesasge"
        '        .value = "ABC"
        '        .regExpress = "[0-9]{3}"
        '        .slug = "choice1-none"
        '        .url = "http://127.0.0.1:8000"
        '        .CurrentForm = Me
        '        .Location = New Point(50, vPosBottom)
        '        vPosBottom = .Location.Y + .Height + 5
        '        .Left = X.Left
        '        .Width = L.Width
        '        .AutoSize = True
        '        '.Dock = DockStyle.Bottom
        '    End With
        '    page.Controls.Add(R)
        '    R.Show()
        '    R.Refresh()
        '    vPosBottom = R.Location.Y + R.Height + 5
        '    R.showOpject()


        '    Dim O As New ucParamOption()
        '    With O
        '        .Name = "Param6"
        '        .title = "Fiber Type"
        '        .message = "Testing Mesasge"
        '        .value = "ABC"
        '        .regExpress = "[0-9]{3}"
        '        .slug = "choice1-none"
        '        .url = "http://127.0.0.1:8000"
        '        .CurrentForm = Me
        '        .Location = New Point(50, vPosBottom)
        '        vPosBottom = .Location.Y + .Height + 5
        '        .Left = X.Left
        '        .Width = L.Width
        '        .AutoSize = True
        '        '.Dock = DockStyle.Bottom
        '    End With
        '    page.Controls.Add(O)
        '    O.Show()
        '    O.showOpject()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' TabControl = Nothing
        Dim aa As Object

        For Each aa In Me.Controls
            If TypeOf aa Is TabControl Then
                aa.Dispose()
            End If
        Next
        CreateObject()
        getOperation()
    End Sub

    Private Sub frmParameter_Load(sender As Object, e As EventArgs) Handles Me.Load
        tss1.Text = vUrl

        getOperation()
    End Sub

    Sub getOperation()
        Dim operations As Object
        Dim operation As Object
        operations = getJsonObject(vUrl + "/api/operation/?name=")
        If Not operations Is Nothing Then
            Dim comboSource As New Dictionary(Of String, String)()
            For Each operation In operations
                comboSource.Add(operation("name") & ":" & operation("title"), operation("name"))
            Next
            With cbOperation
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = New BindingSource(comboSource, Nothing)
                .DisplayMember = "Key"
                .ValueMember = "Value"
                '.SelectedValue = vDefaultValue
            End With
        End If

        'getItemBySlug = json
    End Sub

End Class