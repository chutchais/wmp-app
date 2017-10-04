Imports System.Net
Imports System.Web.Script.Serialization
Imports System.Dynamic




<Runtime.InteropServices.ComVisible(True)>
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        showObjectName()



        Dim vCls As New clsMPFlex
        'initial
        vCls.Form = Me

        Dim vTestScript As String
        Dim xx As String

        vTestScript = tbScript.Text
        txtReturn.Text = vCls.executeScritp(vTestScript)



    End Sub

    '--------Mandatory Funtion for all User Controls-------------
    '--------DO NOT change---------------------------------------
    Public Property getLocalObjectValue(ByVal vObjectName As String) As String

        Get
            Dim vStep() As String
            vStep = Split(vObjectName, ".")
            If UBound(vStep) = 1 Then
                Dim nControl As Control
                nControl = Controls(vStep(0))
                getLocalObjectValue = ""
                'getLocalObjectValue = nControl.getLocalObjectValue(vStep(1))
                getLocalObjectValue = CallByName(nControl, "getLocalObjectValue", Microsoft.VisualBasic.CallType.Get, vStep(1))

            Else
                'getLocalObjectValue = UserControl.Controls(vObjectName)
                getLocalObjectValue = Controls(vObjectName).Text
            End If
        End Get
        Set(ByVal vName As String)
            ' Sets the property value.
            'currentForm = vForm_
        End Set


    End Property

    'Public Property CurrentForm() As Form
    '    Get
    '    End Get
    '    Set(ByVal vCurrForm As Form)
    '        vCurrentFormIn = vCurrForm
    '    End Set
    'End Property

    'Sub initODCSEscript()
    '    objOdcsys.setCustomForm = Me
    '    objOdcsys.setAllowUI = True
    '    objOdcsys.setOnCustomControl

    '    objPrint.setCustomForm = Me
    '    objPrint.setAllowUI = True
    '    objPrint.setOnCustomControl
    'End Sub

    Private Sub showObjectName()
        'On Error Resume Next
        Dim nControl As Control
        Dim toolTip1 As New ToolTip()
        Dim strObjName As String



        For Each nControl In Me.Controls
            strObjName = TypeName(nControl)

            If TypeOf nControl Is UserControl Then
                'Call Show Object for usercontrol
                'nControl.showObjectName()
                'cccc = 0
            Else
                'nControl.Text = nControl.Name
                toolTip1.SetToolTip(nControl, nControl.Name)
            End If
            'If Not TypeOf Me.ParentForm Is Form Then
            '    nControl.set = UserControl.ParentControls.Item(0).Name & "." & Extender.Name & "." & nControl.Name
            'Else
            '    nControl.ToolTipText = Extender.Name & "." & nControl.Name
            'End If
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'MsgBox(getLocalObjectValue("button2"))

        Dim vvvv As ucTest
        vvvv = Controls("uctest1")
        MsgBox(vvvv.Controls("textbox2").Text)

        ' MsgBox(CallByName(vvvv, "getLocalObjectValue", Microsoft.VisualBasic.CallType.Get, "textbox2"))

    End Sub


    'Private Sub GetResponse(uri As Uri, callback As Action(Of Response))
    '    Dim wc As New WebClient()
    '    AddHandler wc.OpenReadCompleted,
    '        Function(o, a)
    '            If callback IsNot Nothing Then
    '                Dim ser As New DataContractJsonSerializer(GetType(Response))
    '                callback(TryCast(ser.ReadObject(a.Result), Response))
    '            End If
    '            Return 0
    '        End Function
    '    wc.OpenReadAsync(uri)
    'End Sub
    Public Shared Function getJsonString(ByVal address As String) As String

        Dim client As WebClient = New WebClient()
        Dim reply As String = client.DownloadString(address)
        Return reply
    End Function

    Public Shared Function getJsonObject(ByVal address As String) As Object

        Dim client As WebClient = New WebClient()
        Dim json As String = client.DownloadString(address)
        Dim jss = New JavaScriptSerializer()
        Dim data = jss.Deserialize(Of Object)(json)
        Return data
    End Function



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim json As Object
        json = getJsonObject("http://127.0.0.1:8000/api/snippet/" + txtSnippetSlug.Text)
        tbScript.Text = json("items")

        'json = getJsonObject("http://127.0.0.1:8000/api/item/" + txtSnippetSlug.Text)
        'Dim ccc As String = json("lists")(0)("title")
        'Dim x As Object
        'For Each x In json("lists")
        '    ccc = x("title")
        'Next
    End Sub
    '--------End Mandatory Funtion for all User Controls-------------

End Class

Public Class MyModel
    Public Property name() As String
    Public Property title() As String
End Class
