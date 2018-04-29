Public Class clsSerialNumber
    Private vSn As String
    Private vSlug As String
    Private vUrl As String 'Object url

    'local object
    Private objSn As Object
    Private objRouting As Object
    Private objWorkOrder As Object
    Private objProduct As Object
    Private objRoutingDetail As Object

    Public vCurrentForm As Form

    Private vAPIServiceUrl As String
    Private objApiService As clsAPIService

    Public Sub New()
        objApiService = New clsAPIService()
        objApiService.Url = vUrl
        objApiService.access_token = vAccessToken
    End Sub

    Public Property serialnumber() As String
        Get
            Return vSn
        End Get
        Set(ByVal Value As String)
            Me.vSn = Value
        End Set
    End Property

    Public Property slug() As String
        Get
            Return vSlug
        End Get
        Set(ByVal Value As String)
            Me.vSlug = Value
        End Set
    End Property

    Public Property url() As String
        Get
            Return vUrl
        End Get
        Set(ByVal Value As String)
            Me.vUrl = Value
            objApiService.Url = vUrl
        End Set
    End Property

    Public Property APIServiceUrl() As String
        Get
            Return vAPIServiceUrl
        End Get
        Set(ByVal Value As String)
            Me.vAPIServiceUrl = Value
            objApiService.Url = Value
        End Set
    End Property

    Dim vAccessToken As String
    Public Property access_token() As String
        Get
            Return vAccessToken
        End Get
        Set(ByVal Value As String)
            Me.vAccessToken = Value
            objApiService.access_token = Value
        End Set
    End Property


    Public Property object_Serialnumber() As Object
        Get
            Return objSn
        End Get
        Set(ByVal Value As Object)
            Me.objSn = Value
        End Set
    End Property

    Public Property object_Routing() As Object
        Get
            Return objRouting
        End Get
        Set(ByVal Value As Object)
            Me.objRouting = Value
        End Set
    End Property

    Public Property object_RoutingDetail() As Object
        Get
            Return objRoutingDetail
        End Get
        Set(ByVal Value As Object)
            Me.objRoutingDetail = Value
        End Set
    End Property

    Public Property object_WorkOrder() As Object
        Get
            Return objWorkOrder
        End Get
        Set(ByVal Value As Object)
            Me.objWorkOrder = Value
        End Set
    End Property

    Public Property object_Product() As Object
        Get
            Return objProduct
        End Get
        Set(ByVal Value As Object)
            Me.objProduct = Value
        End Set
    End Property

    Public Property Form() As Object
        Get
            ' Gets the property value.
            Return vCurrentForm
        End Get
        Set(ByVal vForm_ As Object)
            ' Sets the property value.
            vCurrentForm = vForm_
        End Set
    End Property

    Dim vSuccess As Boolean
    Public Property success As Boolean
        Get
            ' Gets the property value.
            Return vSuccess
        End Get
        Set(ByVal vSuccess_ As Boolean)
            ' Sets the property value.
            vSuccess = vSuccess_
        End Set
    End Property

    Dim vErrorMsg As String
    Public Property error_message As String
        Get
            ' Gets the property value.
            Return vErrorMsg
        End Get
        Set(ByVal vErrorMsg_ As String)
            ' Sets the property value.
            vErrorMsg = vErrorMsg_
        End Set
    End Property

    Dim vMsg As String
    Public Property message As String
        Get
            ' Gets the property value.
            Return vMsg
        End Get
        Set(ByVal vMsg_ As String)
            ' Sets the property value.
            vMsg = vMsg_
        End Set
    End Property



    'Public function
    Public Function getObject(Optional workOrder As String = "") As Object
        Dim objResponse As Object
        objResponse = objApiService.getJsonObject(vUrl + "/api/serialnumber/?q=" + vSn)
        '---Finding Serial number object --
        'Return multiple Record
        'Current workOrder is WIP=True
        If objResponse.length() = 0 Then
            vErrorMsg = "Serial number " & vSn & " doesn't exits in system"
            Return Nothing
        End If

        Dim vWip As Boolean
        Dim obj As Object = Nothing
        For Each obj In objResponse
            vSn = obj("number")
            'objWorkOrder = objSn("workorder")
            vSlug = obj("slug")
            vUrl = obj("url")
            vWip = obj("wip")
            'objRoute = objSn("routing")
            'Looking for Record that WIP == True
            If vWip Then
                Exit For
            End If
        Next

        If Not vWip Then
            'MsgBox("Serial number " & vSn & " is not in WIP",
            'MsgBoxStyle.Critical, "Not in WIP")
            vErrorMsg = "Serial number " & vSn & " is not in WIP"
            Return Nothing
        End If
        objSn = obj

        '--get Serial number more detail (current routing,workorder,Product)
        objRouting = getRouting()

        objSn = objApiService.getObjectByUrl(obj("url"))

        Return objSn
    End Function

    Public Function checkRouting(vOperation As String) As Boolean
        'Start Check Routing process
        Dim vCurrentOpr As String
        Dim vSelectedOpr As String
        '1)Check if SN.curr == Selected.operation
        vCurrentOpr = objSn("current_operation")
        vSelectedOpr = vOperation


        'Get Routing Details (Route + Operation)
        'Dim objRouteDetail As Object
        Dim vRouteDetailSlug As String
        Dim vRouteDetailUrl As String

        objRoutingDetail = objApiService.getRoutingDetail(objRouting("name"), vSelectedOpr)
        If objRoutingDetail.length() = 0 Then
            vErrorMsg = "Operation :" & vOperation & " is not exist in routing :" & objRouting("name")
            Exit Function
        Else
            vRouteDetailSlug = objRoutingDetail(0)("slug")
            vRouteDetailUrl = objRoutingDetail(0)("url")
        End If


        'Get Routing Detail

        objRoutingDetail = objApiService.getObjectByUrl(vRouteDetailUrl)

        If vCurrentOpr = vOperation Then
            'In case Sn.curr.Operation = Selected.Operation  -- check only Reject Routing
            'objRouteDetail("reject_code") is list of slug
            If objRoutingDetail("reject_code").length = 0 Then
                'Operation Matched and No Reject condition.
                Return True
            Else
                'Check Reject Code
                If checkExceptance(objRoutingDetail("reject_code")) Then
                    'exit return False
                    vErrorMsg = "Not allow to operate on this operation " & vbCrLf &
                         vErrorMsg

                    Exit Function
                Else
                    Return True
                End If
            End If
        Else
            'In case Sn.curr.Operation <> Selected.Operation  -- check only Accept Routing
            If objRoutingDetail("accept_code").length = 0 Then
                'Operation Matched and No Reject condition.
                vErrorMsg = "Wrong operation (there is no Accept code configured)"
                Return False
            Else
                'Check Accept Code
                If Not checkAcceptance(objRoutingDetail("accept_code")) Then
                    'exit return False
                    vErrorMsg = "Accept condition failed."
                    Exit Function
                Else
                    Return True
                End If
            End If
        End If

        Return False
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
                    vErrorMsg = "Accept on :" + objAccept("snippet")("name") & " -- " & objAccept("snippet")("title")
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
                    vErrorMsg = "Reject on :" + objReject("snippet")("name") & " -- " & objReject("snippet")("title")
                    checkExceptance = True
                    Exit For
                End If
            End If
            '-------------------
        Next
    End Function


    'Private Function executeScript(vCode As String) As Boolean
    Private Function executeScript(vCode As String) As Object
        Dim vCls As New clsMPFlex
        'initial
        vCls.Form = vCurrentForm
        vCls.Url = vUrl

        Dim vReturn As Object
        vReturn = vCls.executeScritp(vCode)

        With vCls
            vMsg = .message
            vErrorMsg = .error_message
            vSuccess = .success
        End With

        'If Not vCls.success Then
        '    MsgBox(vCls.error_message)
        'End If

        Return vReturn
    End Function


    Private Function getRouting() As Object
        'If there is routing defined ,Get Routing for Serial number
        objSn = objApiService.getObjectByUrl(vUrl)

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
        objProduct = objApiService.getObjectBySlug("product", objWorkOrder("product"))
        'vtmpRoute = objProduct("routing")
        If Not IsNothing(objProduct("routing")) Then
            'Routing assigned
            Return objProduct("routing")
        End If

        vMsg = "Not found Routing setting for this serial number"

        Return Nothing
Exit_Function:
    End Function


End Class
