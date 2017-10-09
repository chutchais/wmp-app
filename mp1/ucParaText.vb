Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.Web.Script.Serialization
Imports System.Dynamic

<Runtime.InteropServices.ComVisible(True)>
Public Class ucParaText
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

    Private vRequired As Boolean
    Public Property required() As Boolean
        Get
            Return vRequired
        End Get
        Set(ByVal value As Boolean)
            vRequired = value
        End Set
    End Property
    'End Property

    Private Sub ucParaText_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.AutoSize = True


        Dim label As New Label
        Dim text As New TextBox
        Dim labelMsg As New Label

        With label
            .Name = "caption"
            .Text = If(vRequired, "*" + vTitleValue, vTitleValue)
            '.Anchor = AnchorStyles.Left + AnchorStyles.Top
            .Dock = DockStyle.Left
        End With
        'With labelMsg
        '    .Name = "lblMsg"
        '    .Text = vMessage
        '    .Dock = DockStyle.Fill
        'End With

        With text
            .Name = "text"
            .Text = vDefaultValue
            .Size = New Size(200, 30)
            '.Anchor = AnchorStyles.Left + AnchorStyles.Top
            .Dock = DockStyle.Right
            .Margin = New Padding(5)

        End With
        AddHandler text.TextChanged, AddressOf text_Changed
        AddHandler text.Validating, AddressOf text_Validating
        AddHandler text.KeyPress, AddressOf text_KeyPress
        AddHandler text.GotFocus, AddressOf text_GotFocus



        Me.Controls.Add(label)
        Me.Controls.Add(text)

        'TableLayout.Controls.Add(label, 0, 0)
        'TableLayout.Controls.Add(text, 1, 0)
        'TableLayout.Controls.Add(labelMsg, 0, 1)
        'TableLayout.SetColumnSpan(labelMsg, 3)


    End Sub

    Private Sub text_KeyPress(sender As Object, e As KeyPressEventArgs)
        Dim txt = DirectCast(sender, TextBox)
        'Dim lbl As New Label
        'lbl = CType(Me.Controls("lblMsg"), Label)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then

            If Not IsValid(txt.Text) Then
                'e.Cancel = True
                txt.Select(0, txt.Text.Length)
                'lbl.Text = "Invalid value"
            Else
                executeScript()
                SendKeys.Send("{TAB}")
            End If
            e.Handled = True
        End If

    End Sub

    Friend Sub text_GotFocus(sender As Object, e As EventArgs)
        Dim txt = DirectCast(sender, TextBox)
        txt.Select(0, txt.Text.Length)
    End Sub
    Friend Sub text_Changed(sender As Object, e As EventArgs)
        'Write code here.
    End Sub

    Private Sub text_Validating(sender As Object, e As CancelEventArgs)
        'If Not IsValid(Me.Text) Then
        '    e.Cancel = True
        '    sender.Select(0, Me.Text.Length)
        'End If
    End Sub

    Function IsValid(ByRef value As String) As Boolean
        Return Regex.IsMatch(value, vRegEx)
    End Function


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
        If iObject("snippet") Is Nothing Then
            getCode = ""
        Else
            getCode = getSnippetBySlug(iObject("snippet"))("code")
        End If

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

    'Private Sub btnShow_Click(sender As Object, e As EventArgs) Handles btnShow.Click
    '    showObjectName()
    'End Sub

    Public Sub showOpject()
        showObjectName()

    End Sub

    '--------End Mandatory Funtion for all User Controls-------------
End Class
