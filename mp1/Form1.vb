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
    '--------End Mandatory Funtion for all User Controls-------------
End Class
