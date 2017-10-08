Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.Web.Script.Serialization
Imports System.Dynamic

<Runtime.InteropServices.ComVisible(True)>
Public Class ucParamList
    'Start Property
    Private vTitleValue As String
    Private vCode As String

    Public Property title() As String
        Get
            Return vTitleValue
        End Get
        Set(ByVal value As String)
            vTitleValue = value
        End Set
    End Property

    Private vMessage As String
    Public Property message() As String
        Get
            Return vMessage
        End Get
        Set(ByVal value As String)
            vMessage = value
        End Set
    End Property

    Private vDefaultValue As String
    Public Property value() As String
        Get
            Return vDefaultValue
        End Get
        Set(ByVal value As String)
            vDefaultValue = value
        End Set
    End Property

    Private vRegEx As String
    Public Property regExpress() As String
        Get
            Return vRegEx
        End Get
        Set(ByVal value As String)
            vRegEx = value
        End Set
    End Property

    Private vSlug As String
    Public Property slug() As String
        Get
            Return vSlug
        End Get
        Set(ByVal value As String)
            vSlug = value
        End Set
    End Property

    Private vUrl As String
    Public Property url() As String
        Get
            Return vUrl
        End Get
        Set(ByVal value As String)
            vUrl = value
        End Set
    End Property

    Private Sub ucParamList_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.AutoSize = True
        Dim label As New Label
        Dim text As New TextBox
        Dim clist As New ComboBox
        Dim labelMsg As New Label

        With label
            .Name = "caption"
            .Text = vTitleValue
            '.Anchor = AnchorStyles.Left + AnchorStyles.Top
            .Dock = DockStyle.Left
        End With
        'With labelMsg
        '    .Name = "lblMsg"
        '    .Text = vMessage
        '    .Dock = DockStyle.Fill
        'End With

        With clist
            .Name = "list"
            '.Text = vDefaultValue
            .Size = New Size(200, 30)
            .Anchor = AnchorStyles.Left
            .Dock = DockStyle.Right
            .Margin = New Padding(5)

        End With
        AddHandler clist.SelectionChangeCommitted, AddressOf clist_SelectionChangeCommitted
        'AddHandler text.Validating, AddressOf text_Validating
        'AddHandler text.KeyPress, AddressOf text_KeyPress
        'AddHandler text.GotFocus, AddressOf text_GotFocus



        Me.Controls.Add(label)
        Me.Controls.Add(clist)

        'Get List------
        Dim vR As String
        vR = getItem(vSlug, clist)
        '--------------
    End Sub

    Friend Sub clist_SelectionChangeCommitted(sender As Object, e As EventArgs)
        Dim clist = DirectCast(sender, ComboBox)
        executeScript()


    End Sub

    'End Property
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

    Private Function getItemBySlug(vItemSlug As String) As Object
        Dim json As Object
        json = getJsonObject(vUrl + "/api/item/" + vItemSlug)
        getItemBySlug = json
    End Function

    Private Function getSnippetBySlug(vSnippetSlug As String) As Object
        Dim json As Object
        json = getJsonObject(vUrl + "/api/snippet/" + vSnippetSlug)
        Return json
    End Function


    Private Function executeScript() As Boolean
        Dim vCls As New clsMPFlex
        'initial
        vCls.Form = vCurrentFormIn
        vCls.Url = "http://127.0.0.1:8000"

        Dim vReturn As String
        vCode = getCode(vSlug)
        vReturn = vCls.executeScritp(vCode)
        'MsgBox(vCls.Url)
        'lblSuccess.Text = vCls.success
        'lblMsg.Text = vCls.message
        If Not vCls.success Then
            MsgBox(vCls.error_message)
        End If

        Return False
    End Function

    Private Function getCode(vSlug_ As String) As String
        Dim iObject As Object
        iObject = getItemBySlug(vSlug_)
        getCode = getSnippetBySlug(iObject("snippet"))("code")
    End Function

    Private Function getItem(vSlug_ As String, ucList As ComboBox) As String
        Dim iObject As Object
        Dim iListObj As Object
        Dim vName As String
        Dim vTile As String
        Dim vValue As String
        Dim vDefault As Boolean
        Dim vOrdered As Integer
        Dim vStatus As String

        iObject = getItemBySlug(vSlug_)
        iListObj = iObject("lists")
        Dim comboSource As New Dictionary(Of String, String)()
        Dim vDefaultValue As String = ""
        For Each aa In iListObj
            vName = aa("name")
            vTile = aa("title")
            vValue = aa("value")
            vDefault = aa("default")
            If vDefault Then
                vDefaultValue = vValue
            End If
            vOrdered = aa("ordered")
            vStatus = aa("status")
            comboSource.Add(vName & ":" & vTile, vValue)
        Next
        With ucList
            .DropDownStyle = ComboBoxStyle.DropDownList
            .DataSource = New BindingSource(comboSource, Nothing)
            .DisplayMember = "Key"
            .ValueMember = "Value"
            .SelectedValue = vDefaultValue
        End With

        Return "s"
        'getCode = getSnippetBySlug(iObject("snippet"))("code")
    End Function

    '--------Mandatory Funtion for all User Controls-------------
    '--------DO NOT change---------------------------------------
    Dim toolTip1 As New ToolTip()
    Dim vParentObjectName As String
    Private vCurrentFormIn As Form

    Public Property ParentObjectName() As String
        Get
            ParentObjectName = vParentObjectName
        End Get
        Set(vValue As String)
            vParentObjectName = vValue
        End Set
    End Property

    Public Property getLocalObjectValue(ByVal vObjectName As String) As String
        Get
            Dim vStep() As String
            vStep = Split(vObjectName, ".")
            If UBound(vStep) > 0 Then
                Dim nControl As Control
                nControl = Controls(vStep(0))
                'getLocalObjectValue = nControl.getLocalObjectValue(vStep(1)).text
                getLocalObjectValue = CallByName(nControl, "", Microsoft.VisualBasic.CallType.Get, vStep(1))
            Else
                getLocalObjectValue = Controls(vObjectName).Text
            End If
        End Get
        Set(value As String)

        End Set
    End Property

    Public Property localControls(vObjName As String) As Object
        Get
            localControls = Me.Controls(vObjName)
        End Get
        Set(value As Object)

        End Set
    End Property

    Public Property CurrentForm() As Form
        Get
            CurrentForm = vCurrentFormIn
        End Get
        Set(vCurrForm As Form)
            vCurrentFormIn = vCurrForm
        End Set
    End Property
    '    Public Property Get localControls(vObjName As String) As Object
    '        Set localControls = UserControl.Controls(vObjName)
    'End Property

    '    Public Property Let CurrentForm(ByVal vCurrForm As Form)
    '        Set vCurrentFormIn = vCurrForm
    'End Property

    '    Public Function initODCSEscript()
    '        objOdcsys.setCustomForm = vCurrentFormIn
    '        objOdcsys.setAllowUI = True
    '        objOdcsys.setOnCustomControl
    '    End Function

    Private Sub showObjectName()
        Dim nControl As Control
        'Comment by Chutchai S on March 2,2009
        'To fix program Crash , because below this command.
        '    For Each nControl In UserControl.Controls
        '        If Not TypeOf UserControl.ParentControls.Item(0) Is Form Then
        '            nControl.ToolTipText = UserControl.ParentControls.Item(0).Name & "." & Extender.Name & "." & nControl.Name
        '        Else
        '            nControl.ToolTipText = Extender.Name & "." & nControl.Name
        '        End If
        '    Next

        'Added by Chutchai S on March 2,2009
        Dim strTooltrip As String
        For Each nControl In Me.Controls
            'nControl.ToolTipText = Extender.Name & "." & nControl.Name

            'If vParentObjectName = "" Then
            '    strTooltrip = Me.Name & "." & nControl.Name
            'Else
            '    strTooltrip = vParentObjectName & "." & Me.Name & "." & nControl.Name
            'End If

            Dim vParentName As String

            If TypeOf Me.Parent Is Form Then
                strTooltrip = Me.Name & "." & nControl.Name
            Else

                vParentName = Me.Parent.Name
                strTooltrip = vParentName & "." & Me.Name & "." & nControl.Name
            End If
            toolTip1.SetToolTip(nControl, strTooltrip)

        Next

    End Sub
    Private Function getExtenderName(vObjName As String) As String
        'Added by Tuk on July 4,2008
        'To protect "send error to MS" box
        'Still not sure.

        On Error GoTo HasError

        'Comment by Chutchai S on March 2,2009
        'To fix program Crash , because below this command.
        '    If Not TypeOf UserControl.ParentControls.Item(0) Is Form Then
        '        getExtenderName = UserControl.ParentControls.Item(0).Name & "." & Extender.Name & "." & vObjName
        '    Else
        '        getExtenderName = Extender.Name & "." & vObjName
        '    End If


        'Added by Chutchai S on March 2,2009
        If vParentObjectName = "" Then
            getExtenderName = Me.Name & "." & vObjName
        Else
            getExtenderName = vParentObjectName & "." & Me.Name & "." & vObjName
        End If

HasError:
    End Function
    Public Sub showOpject()
        showObjectName()

    End Sub
End Class
