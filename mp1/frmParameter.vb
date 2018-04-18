Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.Web.Script.Serialization
Imports System.Dynamic
Imports System.IO
Imports System.Text

<Runtime.InteropServices.ComVisible(True)>
Public Class frmParameter
    Public Shared Function getJsonString(ByVal address As String) As String

        Dim client As WebClient = New WebClient()
        Dim reply As String = client.DownloadString(address)
        Return reply
    End Function

    Public Shared Function getJsonObject3(ByVal address As String) As Object
        Dim client As WebClient = New WebClient()
        Dim json As String = client.DownloadString(address)
        Dim jss = New JavaScriptSerializer()
        Dim data As Object = jss.Deserialize(Of Object)(json)
        Return data
    End Function

    Public Shared Function getJsonObject(ByVal address As String) As Object
        'Support request with Token
        Dim request As WebRequest = WebRequest.Create(address)
        request.Method = "GET"
        Dim byteArray As Byte() = Encoding.UTF8.GetBytes("")
        request.PreAuthenticate = True
        request.Headers.Add("Authorization", "Bearer " + access_token)


        Dim myHttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
        Dim myWebSource As New StreamReader(myHttpWebResponse.GetResponseStream())
        Dim myPageSource As String = myWebSource.ReadToEnd()

        Dim jss = New JavaScriptSerializer()
        Dim data As Object = jss.Deserialize(Of Object)(myPageSource)
        Return data


    End Function

    Private Function getSerialNumber(vSerialNumber As String) As Object
        Dim json As Object
        json = getJsonObject(vUrl + "/api/serialnumber/?wip=true&number=" + vSerialNumber)
        getSerialNumber = json
    End Function

    Private Function getParameterBySlug(vRoutingDetailSlug As String) As Object
        Dim json As Object
        json = getJsonObject(vUrl + "/api/routingdetail/" + vRoutingDetailSlug)
        getParameterBySlug = json
    End Function

    Private Function getItemBySlug(vParameterSlug As String) As Object
        Dim json As Object
        json = getJsonObject(vUrl + "/api/parameter/" + vParameterSlug)
        getItemBySlug = json
    End Function

    Private Function getRoutingDetail(vRoute As String, vOperation As String) As Object
        Dim json As Object
        json = getJsonObject(vUrl + "/api/routingdetail/?route=" + vRoute + "&operation=" & vOperation)
        getRoutingDetail = json
    End Function

    Private Function getRoutingDetail(vSlug As String) As Object
        Dim json As Object
        json = getJsonObject(vUrl + "/api/routingdetail/" & vSlug)
        getRoutingDetail = json
    End Function

    Private Function getSnippetBySlug(vSnippetSlug As String) As Object
        Dim json As Object
        json = getJsonObject(vUrl + "/api/snippet/" + vSnippetSlug)
        Return json
    End Function

    'Private Function getSnippetBySlug(vSnippetSlug As String) As Object
    '    Dim json As Object
    '    json = getJsonObject(vUrl + "/api/snippet/" + vSnippetSlug)
    '    Return json
    'End Function

    Private vUrl As String = Root_url
    Private vCurrentRoutingDetailSlug As String
    Private vDefaultNextPassOperation As String
    Private vDefaultNextFailOperation As String


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        If Not checkSerialNumber(txtSn.Text) Then
            txtSn.Select(0, txtSn.Text.Length)
            txtSn.Select()
        Else
            CreateObject(vCurrentRoutingDetailSlug)
            btnStart.Enabled = False
            btnRefresh.Enabled = True
            txtSn.Enabled = False

            btnPass.Enabled = True
            btnFail.Enabled = True
        End If

    End Sub

    Function checkSerialNumber(vSn As String) As Boolean
        Dim objSns As Object
        objSns = getSerialNumber(vSn)

        If objSns.length() = 0 Then
            MsgBox("Serial number " & vSn & " doesn't exits in system",
            MsgBoxStyle.Critical, "Not found Serial number")
            Exit Function
        End If

        Dim vSnSlug As String = ""
        Dim vWip As Boolean = False
        Dim vSerialNumber As String = ""
        Dim vWorkOrder As Object
        Dim vRoute As Object
        Dim objSn As Object = Nothing
        For Each objSn In objSns
            vSerialNumber = objSn("number")
            vWorkOrder = objSn("workorder")
            vSnSlug = objSn("slug")
            vWip = objSn("wip")
            vRoute = objSn("routing")

        Next

        If Not vWip Then
            MsgBox("Serial number " & vSn & " is not in WIP",
            MsgBoxStyle.Critical, "Not in WIP")
            Exit Function
        End If


        'Get Routing for Serial number
        vRoute = getSerialNumberRouteObject(vSerialNumber, vWorkOrder, vRoute, objSn)
        If IsNothing(vRoute) Then
            Exit Function
        End If

        'Get Routing Details (Route + Operation)
        Dim objRouteDetail As Object
        Dim vOperation As String = cbOperation.SelectedValue
        objRouteDetail = getRoutingDetail(vRoute("name"), vOperation)
        If objRouteDetail.length() = 0 Then
            MsgBox("Operation :" & cbOperation.SelectedValue & " is not exist in routing :" & vRoute("name"),
                   MsgBoxStyle.Critical, "Operation not exist.")
            Exit Function
        Else
            'Final Route Detail Slug
            vCurrentRoutingDetailSlug = objRouteDetail(0)("slug")
        End If

        'Check Current Operation-
        Dim vCurrentOpr As String
        Dim vObjRoutingDetail As Object
        vObjRoutingDetail = getRoutingDetail(vCurrentRoutingDetailSlug)

        If Not IsNothing(objSn("current_operation")) Then
            vCurrentOpr = objSn("current_operation")("name")
        Else
            vCurrentOpr = objSn("current_operation")
        End If

        'get Next pass and Next Fail.
        If Not IsNothing(vObjRoutingDetail("next_pass")) Then
            vDefaultNextPassOperation = vObjRoutingDetail("next_pass")("name")
        End If
        If Not IsNothing(vObjRoutingDetail("next_fail")) Then
            vDefaultNextFailOperation = vObjRoutingDetail("next_fail")("name")
        End If
        '-----------------------------

        If vCurrentOpr <> vOperation Then
            'Case Current in system with working operation is not same
            'Need to check Acceptance Code
            If Not checkAcceptance(vObjRoutingDetail("accept_code")) Then
                'exit return False
                MsgBox("Either Wrong operation or No acceptace condition accepted",
                        MsgBoxStyle.Critical, "Not accept to perform")
                Exit Function
            End If
        End If


        If checkExceptance(vObjRoutingDetail("except_code")) Then
            'exit return False
            MsgBox("Not allow to operate on this operation",
                        MsgBoxStyle.Critical, "Exceptance check")
            Exit Function
        End If


        'If everything is Okay ,it will return True
        Return True
    End Function

    Function checkAcceptance(vAcceptObjs As Object) As Boolean
        'ANY True -- return True
        Dim vAcceptObj As Object
        Dim objSnippet As Object
        Dim vSnippetSlug As String = ""
        Dim vCode As String
        checkAcceptance = False
        For Each vAcceptObj In vAcceptObjs
            vSnippetSlug = vAcceptObj("snippet")("slug")
            objSnippet = getSnippetBySlug(vSnippetSlug)
            vCode = objSnippet("code")
            'Execute Script-----
            If vCode <> "" And objSnippet("status") = "A" Then
                If executeScript(vCode) Then
                    checkAcceptance = True
                    Exit For
                End If
            End If
            '-------------------
        Next
    End Function

    Function checkExceptance(vExceptObjs As Object) As Boolean
        'ANY True -- return True
        Dim vExceptObj As Object
        Dim objSnippet As Object
        Dim vSnippetSlug As String = ""
        Dim vCode As String
        checkExceptance = False
        For Each vExceptObj In vExceptObjs
            vSnippetSlug = vExceptObj("snippet")("slug")
            objSnippet = getSnippetBySlug(vSnippetSlug)
            vCode = objSnippet("code")
            'Execute Script-----
            If vCode <> "" And objSnippet("status") = "A" Then
                If executeScript(vCode) Then
                    checkExceptance = True
                    Exit For
                End If
            End If
            '-------------------
        Next
    End Function

    Private Function executeScript(vCode As String) As Boolean
        Dim vCls As New clsMPFlex
        'initial
        vCls.Form = Me
        vCls.Url = vUrl

        Dim vReturn As Object
        vReturn = vCls.executeScritp(vCode)
        'MsgBox(vCls.Url)
        'lblSuccess.Text = vCls.success
        'lblMsg.Text = vCls.message
        If Not vCls.success Then
            MsgBox(vCls.error_message)
        End If

        Return vReturn
    End Function

    Function getSerialNumberRouteObject(vSn As String, vWorkOrder As Object, vRoute As Object, objSn As Object) As Object
        'Return Route of Serial number
        If Not IsNothing(vRoute) Then
            Return vRoute
        End If

        'Check WorkOrder Routing
        If Not IsNothing(vWorkOrder("routing")) Then
            Return vWorkOrder("routing")
        End If

        'Check Product Routing
        If Not IsNothing(vWorkOrder("product")("routing")) Then
            Return vWorkOrder("product")("routing")
        End If

        MsgBox("Not found Routing setting for this serial number", MsgBoxStyle.Critical, "Not found routing setting")
        Return Nothing
    End Function


    Sub CreateObject(vRoutingDetailSlug As String)
        Dim objParams As Object
        objParams = getParameterBySlug(vRoutingDetailSlug)

        Dim tabControl As New TabControl
        'Dim tabPage As New TabPage
        With tabControl
            .Name = "Parameter"
            .TabIndex = 3
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
            vParamSlug = objParam("slug")
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        ' TabControl = Nothing
        Dim aa As Object

        For Each aa In Me.Controls
            If TypeOf aa Is TabControl Then
                aa.Dispose()
            End If
        Next
        CreateObject(vCurrentRoutingDetailSlug)
        getOperation()
    End Sub

    Private Sub frmParameter_Load(sender As Object, e As EventArgs) Handles Me.Load
        tss1.Text = vUrl

        'Get authorized operation
        getOperation(user_id)


        txtSn.Select()
        showObjectName()
    End Sub

    Sub getOperation(Optional vUserId As String = "")
        Dim operations As Object
        Dim operation As Object

        'if vUserId is Blank --> get all existing operation
        'if exist -->get only authorize operation

        If vUserId <> "" Then
            operations = getJsonObject(vUrl + "/api/users/" & vUserId & "/")("operations")
        Else
            operations = getJsonObject(vUrl + "/api/operation/")
        End If




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

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim aa As Object
        For Each aa In Me.Controls
            If TypeOf aa Is TabControl Then
                aa.Dispose()
            End If
        Next
        txtSn.Enabled = True
        txtSn.Select(0, txtSn.Text.Length)
        txtSn.Select()
        btnStart.Enabled = True
        btnRefresh.Enabled = False

        btnPass.Enabled = False
        btnFail.Enabled = False
    End Sub



    Private Sub txtSn_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSn.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            Button1_Click(sender, e)
        End If
    End Sub

    Private Sub btnFail_Click(sender As Object, e As EventArgs) Handles btnFail.Click
        checkNextCondition(vDefaultNextFailOperation)
    End Sub

    Private Sub btnPass_Click(sender As Object, e As EventArgs) Handles btnPass.Click
        checkNextCondition(vDefaultNextPassOperation)
    End Sub

    Function checkNextCondition(vDefaultNextOpr As String) As Boolean
        Dim objRoutingDetail As Object
        objRoutingDetail = getRoutingDetail(vCurrentRoutingDetailSlug)

        Dim vNextoprObjs As Object
        vNextoprObjs = objRoutingDetail("next_code")
        'ANY True -- return True
        Dim vNextoprObj As Object
        Dim objSnippet As Object
        Dim vSnippetSlug As String = ""
        Dim vTitle As String
        Dim vNextOpr As String = ""
        Dim vCode As String
        checkNextCondition = False
        For Each vNextoprObj In vNextoprObjs
            vSnippetSlug = vNextoprObj("slug")
            objSnippet = getSnippetBySlug(vSnippetSlug)
            vTitle = vNextoprObj("title")
            vNextOpr = vNextoprObj("operation")
            vCode = objSnippet("code")
            'Execute Script-----
            If vCode <> "" And objSnippet("status") = "A" Then
                If executeScript(vCode) Then
                    ' checkAcceptance = True
                    MsgBox(vTitle & " is correct condition " & vbCrLf &
                           "System will move unit to operation " & vNextOpr,
                           MsgBoxStyle.Information, "Routing condition information")
                    Exit Function
                End If
            End If
            '-------------------
        Next
        MsgBox("System will move unit to operation " & vDefaultNextOpr,
                           MsgBoxStyle.Information, "Default Routing")
    End Function

    'Must Have function
    Dim toolTip1 As New ToolTip()
    Private Sub showObjectName()
        Dim nControl As Control
        Dim strTooltrip As String
        For Each nControl In Me.Controls
            strTooltrip = nControl.Name
            toolTip1.SetToolTip(nControl, strTooltrip)

        Next
    End Sub


    '------------------
End Class