Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.Web.Script.Serialization
Imports System.Dynamic
Imports System.IO
Imports System.Text

<Runtime.InteropServices.ComVisible(True)>
Public Class frmParameter

    Dim objApiService As New clsAPIService

    'Using Function()
    Public Shared Function getJsonString(ByVal address As String) As String

        Dim client As WebClient = New WebClient()
        Dim reply As String = client.DownloadString(address)
        Return reply
    End Function

    'Public Shared Function getJsonObject3(ByVal address As String) As Object
    '    Dim client As WebClient = New WebClient()
    '    Dim json As String = client.DownloadString(address)
    '    Dim jss = New JavaScriptSerializer()
    '    Dim data As Object = jss.Deserialize(Of Object)(json)
    '    Return data
    'End Function

    'Public Shared Function getJsonObject(ByVal address As String) As Object
    '    'Support request with Token
    '    Dim request As WebRequest = WebRequest.Create(address)
    '    request.Method = "GET"
    '    Dim byteArray As Byte() = Encoding.UTF8.GetBytes("")
    '    request.PreAuthenticate = True
    '    request.Headers.Add("Authorization", "Bearer " + access_token)


    '    Dim myHttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
    '    Dim myWebSource As New StreamReader(myHttpWebResponse.GetResponseStream())
    '    Dim myPageSource As String = myWebSource.ReadToEnd()

    '    Dim jss = New JavaScriptSerializer()
    '    Dim data As Object = jss.Deserialize(Of Object)(myPageSource)
    '    Return data


    'End Function

    Private Function getSerialNumber(vSerialNumber As String) As Object
        Dim json As Object
        json = objApiService.getJsonObject(vUrl + "/api/serialnumber/?q=" + vSerialNumber)
        getSerialNumber = json
    End Function


    'Private Function getObjectByUrl(vUrl As String) As Object
    '    Dim json As Object
    '    json = getJsonObject(vUrl)
    '    getObjectByUrl = json
    'End Function

    'Private Function getObjectBySlug(vApp As String, vSlug As String) As Object
    '    Dim json As Object
    '    Dim vUrlSlug As String
    '    vUrlSlug = vUrl + "/api/" + vApp + "/" + vSlug & "/"
    '    json = getJsonObject(vUrlSlug)
    '    getObjectBySlug = json
    'End Function

    'Private Function getRoutingDetail(vRoute As String, vOperation As String) As Object
    '    'VRoute = Routing name
    '    'vOperation = Operation name
    '    Dim json As Object
    '    json = getJsonObject(vUrl + "/api/routing-detail/?route=" + vRoute + "&operation=" & vOperation)
    '    Return json
    'End Function

    ''End used Function

    'Private Function getParameterBySlug(vRoutingDetailSlug As String) As Object
    '    Dim json As Object
    '    json = getJsonObject(vUrl + "/api/routingdetail/" + vRoutingDetailSlug)
    '    getParameterBySlug = json
    'End Function

    'Private Function getItemBySlug(vParameterSlug As String) As Object
    '    Dim json As Object
    '    json = getJsonObject(vUrl + "/api/parameter/" + vParameterSlug)
    '    getItemBySlug = json
    'End Function



    'Private Function getRouting(vRouteSlug As String) As Object
    '    Dim json As Object
    '    json = getJsonObject(vUrl + "/api/routing/" + vRouteSlug)
    '    getRouting = json
    'End Function



    'Private Function getRoutingDetail(vSlug As String) As Object
    '    Dim json As Object
    '    json = getJsonObject(vUrl + "/api/routingdetail/" & vSlug)
    '    getRoutingDetail = json
    'End Function

    'Private Function getSnippetBySlug(vSnippetSlug As String) As Object
    '    Dim json As Object
    '    json = getJsonObject(vUrl + "/api/snippet/" + vSnippetSlug)
    '    Return json
    'End Function

    'Private Function getWorkOrder(vWorkOrderSlug As String) As Object
    '    Dim json As Object
    '    json = getJsonObject(vUrl + "/api/workorder/" + vWorkOrderSlug + "/")
    '    getWorkOrder = json
    'End Function

    'Private Function getProduct(vProductSlug As String) As Object
    '    Dim json As Object
    '    json = getJsonObject(vUrl + "/api/product/" + vProductSlug + "/")
    '    getProduct = json
    'End Function
    'Private Function getSnippetBySlug(vSnippetSlug As String) As Object
    '    Dim json As Object
    '    json = getJsonObject(vUrl + "/api/snippet/" + vSnippetSlug)
    '    Return json
    'End Function

    Private vUrl As String = Root_url
    Private vCurrentRoutingDetailSlug As String
    Private vCurrentRoutingDetailUrl As String
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

        'Return multiple Record
        'Current workOrder is WIP=True
        If objSns.length() = 0 Then
            MsgBox("Serial number " & vSn & " doesn't exits in system",
            MsgBoxStyle.Critical, "Not found Serial number")
            Exit Function
        End If

        Dim vSnSlug As String = ""
        Dim vSnUrl As String = ""
        Dim vWip As Boolean = False
        Dim vSerialNumber As String = ""

        'Dim objWorkOrder As Object = Nothing
        Dim vWorkOrder As String = ""

        Dim objRoute As Object = Nothing
        Dim vRoute As String = ""

        Dim objSn As Object = Nothing
        For Each objSn In objSns
            vSerialNumber = objSn("number")
            'objWorkOrder = objSn("workorder")
            vSnSlug = objSn("slug")
            vWip = objSn("wip")
            'objRoute = objSn("routing")
            'Looking for Record that WIP == True
            If vWip Then
                'vWorkOrder = objWorkOrder("name")
                vSnUrl = objSn("url")
                Exit For
            End If
        Next

        If Not vWip Then
            MsgBox("Serial number " & vSn & " is not in WIP",
            MsgBoxStyle.Critical, "Not in WIP")
            Exit Function
        End If

        'Get Serial number details for particular SN
        objSn = objApiService.getObjectByUrl(vSnUrl)

        'Get Routing for Serial number
        Dim objRouting As Object = Nothing
        objRouting = getSerialNumberRouteObject(vSn, vSnUrl)
        If IsNothing(objRouting) Then
            Exit Function
        End If

        'Start Check Routing process
        Dim vCurrentOpr As String
        Dim vSelectedOpr As String
        '1)Check if SN.curr == Selected.operation
        vCurrentOpr = objSn("current_operation")
        vSelectedOpr = cbOperation.SelectedValue


        'Get Routing Details (Route + Operation)
        Dim objRouteDetail As Object

        objRouteDetail = objApiService.getRoutingDetail(objRouting("name"), vSelectedOpr)
        If objRouteDetail.length() = 0 Then
            MsgBox("Operation :" & cbOperation.SelectedValue & " is not exist in routing :" & vRoute("name"),
                   MsgBoxStyle.Critical, "Operation not exist.")
            Exit Function
        Else
            'Final Route Detail Slug
            vCurrentRoutingDetailSlug = objRouteDetail(0)("slug")
            vCurrentRoutingDetailUrl = objRouteDetail(0)("url")
        End If

        'Get Routing Detail
        objRouteDetail = objApiService.getObjectByUrl(vCurrentRoutingDetailUrl)

        If vCurrentOpr = vSelectedOpr Then
            'In case Sn.curr.Operation = Selected.Operation  -- check only Reject Routing
            If IsNothing(objRouteDetail("reject_code")) Then
                'Operation Matched and No Reject condition.
                Return True
            Else
                'Check Reject Code
                If checkExceptance(objRouteDetail("reject_code")) Then
                    'exit return False
                    MsgBox("Not allow to operate on this operation",
                                MsgBoxStyle.Critical, "Exceptance check")
                    Exit Function
                End If
            End If
        Else
            'In case Sn.curr.Operation <> Selected.Operation  -- check only Accept Routing
            If IsNothing(objRouteDetail("accept_code")) Then
                'Operation Matched and No Reject condition.
                Return False
            Else
                'Check Accept Code
                If Not checkAcceptance(objRouteDetail("accept_code")) Then
                    'exit return False
                    MsgBox("Either Wrong operation or No acceptace condition accepted",
                            MsgBoxStyle.Critical, "Not accept to perform")
                    Exit Function
                End If
            End If


        End If




        ''Check Current Operation-
        'Dim vCurrentOpr As String
        'Dim vObjRoutingDetail As Object
        'vObjRoutingDetail = getRoutingDetail(vCurrentRoutingDetailSlug)

        'If Not IsNothing(objSn("current_operation")) Then
        '    vCurrentOpr = objSn("current_operation")("name")
        'Else
        '    vCurrentOpr = objSn("current_operation")
        'End If

        ''get Next pass and Next Fail.
        'If Not IsNothing(vObjRoutingDetail("next_pass")) Then
        '    vDefaultNextPassOperation = vObjRoutingDetail("next_pass")("name")
        'End If
        'If Not IsNothing(vObjRoutingDetail("next_fail")) Then
        '    vDefaultNextFailOperation = vObjRoutingDetail("next_fail")("name")
        'End If
        ''-----------------------------







        'If everything is Okay ,it will return True
        Return True
    End Function

    Function checkAcceptance(vAcceptSlugLists As Object) As Boolean
        'ANY True -- return True
        Dim objAccept As Object
        Dim vAcceptSlug As String

        Dim vSnippetSlug As String = ""
        Dim vCode As String
        checkAcceptance = False
        For Each vAcceptSlug In vAcceptSlugLists
            objAccept = objApiService.getObjectBySlug("routing-accept", vAcceptSlug)

            vSnippetSlug = objAccept("snippet")("slug")
            'objSnippet = getSnippetBySlug(vSnippetSlug)
            vCode = objAccept("snippet")("code")
            'Execute Script-----
            If vCode <> "" And objAccept("snippet")("status") = "A" Then
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
        Dim objReject As Object
        Dim vRejectSlug As String
        Dim vCode As String
        checkExceptance = False
        For Each vRejectSlug In vExceptObjs
            objReject = objApiService.getObjectBySlug("routing-reject", vRejectSlug)
            vCode = objReject("snippet")("code")
            'Execute Script-----
            If vCode <> "" And objReject("snippet")("status") = "A" Then
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

    Function getSerialNumberRouteObject(vSn As String, vSnUrl As String) As Object
        Dim objSn As Object = Nothing
        Dim objRouting As Object = Nothing
        Dim objWorkOrder As Object = Nothing

        'Get Serial number details for specific SN (by Url)
        objSn = objApiService.getObjectByUrl(vSnUrl)

        'If there is routing defined ,Get Routing for Serial number
        If objSn("routing") <> "" Then
            objRouting = objApiService.getObjectBySlug("routing", objSn("routing"))
            Return objRouting
        End If

        'Get WorkOrder's Routing
        objWorkOrder = objApiService.getObjectBySlug("workorder", objSn("workorder"))
        If Not IsNothing(objWorkOrder("routing")) Then
            Return objWorkOrder("routing")
        End If


        'Check Product's Routing
        'Dim vProduct As String
        Dim objProduct As Object
        objProduct = objApiService.getObjectBySlug("product", objWorkOrder("product"))
        'vtmpRoute = objProduct("routing")
        If Not IsNothing(objProduct("routing")) Then
            'Routing assigned
            Return objProduct("routing")
        End If
        MsgBox("Not found Routing setting for this serial number",
               MsgBoxStyle.Critical, "Not found routing setting")
        Return Nothing
Exit_Function:
        'Return Routing Detail
        'Return Nothing
    End Function


    Sub CreateObject(vRoutingDetailSlug As String)
        Dim objRoutingDetail As Object
        objRoutingDetail = objApiService.getObjectBySlug("routing-detail", vRoutingDetailSlug)

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
        For Each vParamSlug In objRoutingDetail("parameter")
            '-----get Parameter Object----
            objParam = objApiService.getObjectBySlug("parameter", vParamSlug)
            '-----------------------------
            Dim tabPage As New TabPage
            tabPage.Text = IIf(IsDBNull(objParam("title")), objParam("name"), objParam("title"))
            tabPage.Name = objParam("name")
            tabPage.AutoScroll = True
            tabPage.AutoSize = True
            ' AddHandler Text.Validating, AddressOf text_Validating
            tabControl.TabPages.Add(tabPage)
            ''------Add Parameter to Page---
            'vParamSlug = objParam("slug")
            addParameterToPage(vParamSlug, tabPage)
            ''------------------------------
        Next
        tabControl.Show()
        Me.Controls.Add(tabControl)
    End Sub


    Sub addParameterToPage(vParameterSlug As String, page As TabPage)
        Dim objItems As Object
        Dim objItem As Object
        'objItems = getItemBySlug(vParameterSlug)
        objItems = objApiService.getObjectBySlug("parameter", vParameterSlug)("items")

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
        For Each objItem In objItems
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

        'Root_url from Module mdlWMP
        objApiService.Url = Root_url

        tss1.Text = objApiService.Url
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
            operations = objApiService.getJsonObject(vUrl + "/api/users/" & vUserId & "/")("operations")
        Else
            operations = objApiService.getJsonObject(vUrl + "/api/operation/")
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
        objRoutingDetail = objApiService.getRoutingDetail(vCurrentRoutingDetailSlug, "")

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
            objSnippet = objApiService.getObjectBySlug("snippet", vSnippetSlug)
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