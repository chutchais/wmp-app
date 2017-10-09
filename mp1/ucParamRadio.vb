Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.Web.Script.Serialization
Imports System.Dynamic



Public Class BaseControl
    Protected Overridable Sub OnLoad(ByVal e As EventArgs)
        'MessageBox.Show("BaseControl Click")
        'base.OnLoad(e)
    End Sub
End Class

<Runtime.InteropServices.ComVisible(True)>
Public Class ucParamRadio
    'Start Property
    Private vFinalHeight As Integer
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

    Private vSelectedItem As String
    Public Property selectedItem() As String
        Get
            Return vSelectedItem
        End Get
        Set(ByVal value As String)
            vSelectedItem = value
        End Set
    End Property

    Private vSelectedValue As String
    Public Property selectedValue() As String
        Get
            Return vSelectedValue
        End Get
        Set(ByVal value As String)
            vSelectedValue = value
        End Set
    End Property

    Private vNewHeight As String
    Public Property Height_New() As String
        Get
            Return vFinalHeight
        End Get
        Set(ByVal value As String)
            vNewHeight = value
        End Set
    End Property

    Protected Overridable Sub OnLoad(e As EventArgs)

    End Sub

    Private Sub ucParamList_Load(sender As Object, e As EventArgs) Handles Me.Load


        Dim label As New Label
        Dim text As New TextBox
        Dim cradio As New RadioButton
        Dim labelMsg As New Label

        With label
            .Name = "caption"
            .Text = vTitleValue
            .AutoSize = True
            '.Anchor = AnchorStyles.Left + AnchorStyles.Top
            .Dock = DockStyle.Left
        End With
        'With labelMsg
        '    .Name = "lblMsg"
        '    .Text = vMessage
        '    .Dock = DockStyle.Fill
        'End With

        With cradio
            .Name = "radio"

            '.Size = New Size(200, 30)
            .Anchor = AnchorStyles.Left
            .Dock = DockStyle.Right
            .Margin = New Padding(5)

        End With
        'AddHandler clist.SelectionChangeCommitted, AddressOf clist_SelectionChangeCommitted
        'AddHandler text.Validating, AddressOf text_Validating
        'AddHandler text.KeyPress, AddressOf text_KeyPress
        'AddHandler text.GotFocus, AddressOf text_GotFocus



        Me.Controls.Add(label)
        'Me.Controls.Add(cradio)

        'Get List------
        Dim vR As String
        vR = getItem(vSlug, cradio)
        '--------------
        Me.AutoSize = True
    End Sub

    Friend Sub clist_SelectionChangeCommitted(sender As Object, e As EventArgs)
        Dim clist = DirectCast(sender, ComboBox)
        vSelectedItem = clist.Text
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
        If vCode <> "" Then
            vReturn = vCls.executeScritp(vCode)
            If Not vCls.success Then
                MsgBox(vCls.error_message)
            End If
        End If
        Return False
    End Function

    Private Function getCode(vSlug_ As String) As String
        Dim iObject As Object
        iObject = getItemBySlug(vSlug_)
        getCode = getSnippetBySlug(iObject("snippet"))("code")
    End Function

    Private Function getItem(vSlug_ As String, ucRadio As RadioButton) As String
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
        Dim vPosBottom As Integer = 0
        Dim vRadioIndex As Integer = 0
        Dim vXFirstCol As Integer = 160
        Dim vXSecondCol As Integer = 250
        Dim ca As Boolean
        Dim x, y As Integer
        Dim radBtn As New RadioButton
        For Each aa In iListObj
            vName = aa("name")
            vTile = aa("title")
            vValue = aa("value")
            vDefault = aa("default")

            vOrdered = aa("ordered")
            vStatus = aa("status")
            If vStatus = "A" Then
                ca = vRadioIndex Mod 2
                If ca = 0 Then
                    x = vXFirstCol
                    y = vPosBottom
                Else
                    x = vXSecondCol
                    y = vPosBottom
                End If
                radBtn = New RadioButton With {
                                .Text = vTile,
                                .Name = "radio" + vRadioIndex.ToString,
                                .Location = New Point(x, y),
                                .Tag = vRadioIndex,
                                .AutoSize = True,
                                .Checked = vDefault
                                }
                Me.Controls.Add(radBtn)

                AddHandler radBtn.Click, AddressOf radio_Click ' add an event
                If ca Then
                    vPosBottom = vPosBottom + 20
                End If
                vFinalHeight = vPosBottom + radBtn.Size.Height
                'Me.Size = New Size(Me.Size.Width, vPosBottom + radBtn.Size.Height)
                vRadioIndex = vRadioIndex + 1
            End If

        Next
        'Dim ccc As Integer


        Return "s"

    End Function

    Private Sub radio_Click(sender As Object, e As EventArgs)
        Dim radBtn = DirectCast(sender, RadioButton)
        vSelectedItem = radBtn.Text
        vSelectedValue = radBtn.Text
        executeScript()
    End Sub

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
