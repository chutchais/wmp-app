﻿<Runtime.InteropServices.ComVisible(True)>
Public Class clsMPFlex
    Dim currentForm As Form
    Public ObjectName As String 'argument of local function
    Public vObjName As String 'For access object name
    Public vReturn As String 'argument for Return

    'Dim vScriptControl As Object
    Public vScriptControl As MSScriptControl.ScriptControlClass

    'Property
    Public Property Form() As Object
        Get
            ' Gets the property value.
            Return currentForm
        End Get
        Set(ByVal vForm_ As Object)
            ' Sets the property value.
            currentForm = vForm_
        End Set
    End Property
    'ReadOnly Property Length() As Integer
    '    Get
    '        Return mstrLine.Length
    '    End Get
    'End Property

    'Method
    'Public Sub Capitalize()
    '    ' Capitalize the value of the property.
    '    userNameValue = UCase(userNameValue)
    'End Sub
    Public Function executeScritp(vScritp As String) As String
        On Error GoTo HasError
        initialScript()
        vScriptControl.ExecuteStatement(vScritp)
        executeScritp = vScriptControl.Eval("returndata("""")")
        vScriptControl.Reset()
        Exit Function
HasError:
        destroyObject()
    End Function

    Private Sub destroyObject()
        vScriptControl = Nothing
    End Sub

    Public Function returnData(vOutput As String) As String
        'getValueByParameter
        returnData = vReturn
    End Function


    Private Sub initialScript()


        'Set clsODC = New clsODCSEscript 'Added by Chuthcai S on June 18,2009
        Dim clsFlex As New clsMPFlex

        vScriptControl = New MSScriptControl.ScriptControlClass

        With vScriptControl
            .Language = "VBScript"
            .AllowUI = True

            'Core Objects
            .AddObject("flex", clsFlex, True)
            '    vScriptControl.AddObject "dbCon", vDBconnection ', True
            '    vScriptControl.AddObject "clsODCSEscript", clsODC ', True
            .AddObject("clsMain", vScriptControl, True) 'For includelib
            If Not currentForm Is Nothing Then
                .AddObject("clsForm", currentForm, True)
            End If

        End With


        'New function to get local object

        Dim vgetObj As String
        vgetObj = "Function localObject(vObjectName,vNewName)" & vbCrLf &
                "flex.ObjectName =  vObjectName " & vbCrLf &
                "flex.Form = clsForm " & vbCrLf &
                "clsMain.AddObject vNewName,flex.getObject("""") " & vbCrLf &
                "End Function"

        vScriptControl.AddCode(vgetObj)


        'Return
        Dim vReturn_ As String
        vReturn_ = "Function return(vReturn_) " & vbCrLf &
                " dim output_" & vbCrLf &
                "    flex.vReturn = vReturn_ " & vbCrLf &
                "End Function"
        vScriptControl.AddCode(vReturn_)

        vReturn_ = "Function returndata(vReturn_) " & vbCrLf &
                " dim output_" & vbCrLf &
                "    output_ = flex.returnData("""") " & vbCrLf &
                "   returndata = output_ " & vbCrLf &
                "End Function"
        vScriptControl.AddCode(vReturn_)







        '    'Added by Chutchai S on Dec 19,2007
        '    'To add result XML in to class
        '    Dim objXml As Object
        'Set objXml = CreateObject("MSXML.DOMDocument")
        'If vXmlResult <> "" Then
        '        objXml.loadXML vXmlResult
        '    vScriptControl.AddObject "objXml_tmp", objXml, True
        'End If
        '    '---------------------------------------------



        '    'Funtions
        '    'New function for get parameter (DES version 3.8.14)
        '    'Added by Chutchai S on Oct 16,2009
        '    vXmlRequest = "Function request(vRequestIn) " & vbCrLf &
        '                " dim vQueryString " & vbCrLf &
        '                " vQueryString =""/request/"" & vRequestIn  " & vbCrLf &
        '                " if not objXml_tmp.selectSingleNode(vQueryString) is nothing then " & vbCrLf &
        '                "       request = objXml_tmp.selectSingleNode(vQueryString).text " & vbCrLf &
        '                " end if " & vbCrLf &
        '                "End Function"
        '    vScriptControl.AddCode vXmlRequest
        ''--------------------------------------------------------


        '    'New function for get current result XML data
        '    'Added by Chutchai S on Dec 19,2007
        '    vXml = "Function getResultXML() " & vbCrLf &
        '            " getResultXML = clsODCSEscript.getXml(objXml_tmp.xml) " & vbCrLf &
        '            "End Function"
        '    If vXmlResult <> "" Then vScriptControl.AddCode vXml

        ''New function to include library (vFlexLibName)
        '    vInclude = "Function includeLib(vLibrary)" & vbCrLf &
        '            "clsODCSEscript.vFlexLibName = vLibrary " & vbCrLf &
        '            "clsMain.AddCode clsODCSEscript.getLibrary("""",dbCon) " & vbCrLf &
        '            "End Function"

        '    vScriptControl.AddCode vInclude

        'New function to get local object
        'vgetObj = "Function localObject(vObjectName,vNewName)" & vbCrLf &
        '        "clsODCSEscript.vObjName=  vObjectName " & vbCrLf &
        '        "clsODCSEscript.setCurrentForm clsForm " & vbCrLf &
        '        "clsMain.AddObject vNewName,clsODCSEscript.getObject("""") " & vbCrLf &
        '        "End Function"

        'vScriptControl.AddCode vgetObj

        ''Parameter definetion (par)
        '    '    fpara = "Function par(param)" & vbCrLf & _
        '    '                "Dim Result                    " & vbCrLf & _
        '    '                "Dim Sn                    " & vbCrLf & _
        '    '                "Dim So                    " & vbCrLf & _
        '    '                "Dim Product                    " & vbCrLf & _
        '    '                "clsODCSEscript.vScript = param " & vbCrLf & _
        '    '                "Sn =""" & vSN & """ " & vbCrLf & _
        '    '                "So =""" & vSO & """ " & vbCrLf & _
        '    '                "Product =""" & vProduct & """ " & vbCrLf & _
        '    '                "    Result = clsODCSEscript.getParameter("""",""" & vSN & """,""" & vSO & """,""" & vProduct & """,dbCon) " & vbCrLf & _
        '    '                "    par = Result    " & vbCrLf & _
        '    '                "End Function"

        '    fpara = "Function par(param)" & vbCrLf &
        '            "Dim Result                    " & vbCrLf &
        '            "Dim Sn                    " & vbCrLf &
        '            "Dim So                    " & vbCrLf &
        '            "Dim Product                    " & vbCrLf &
        '            "    clsODCSEscript.DTS_location = """ & vDtsLocation & """ " & vbCrLf &
        '            "    clsODCSEscript.DTS_mode = " & vDtsMode & " " & vbCrLf &
        '            "clsODCSEscript.vScript = param " & vbCrLf &
        '            "Sn =""" & vSN & """ " & vbCrLf &
        '            "So =""" & vSO & """ " & vbCrLf &
        '            "Product =""" & vProduct & """ " & vbCrLf &
        '            "    Result = clsODCSEscript.getParameter("""",""" & vSN & """,""" & vSO & """,""" & vProduct & """,dbCon) " & vbCrLf &
        '            "    par = Result    " & vbCrLf &
        '            "End Function"

        '    vScriptControl.AddCode fpara



        '    floc = "Function local(param)" & vbCrLf &
        '            "Dim Result                    " & vbCrLf &
        '            "clsODCSEscript.vObjectName = param " & vbCrLf &
        '            "  if not " & IIf(vRunAsService, "true", "false") & " then " & vbCrLf &
        '            "    clsODCSEscript.setCurrentForm clsForm " & vbCrLf &
        '            "    Result = clsODCSEscript.getLocalObject("""") " & vbCrLf &
        '            "  else " & vbCrLf &
        '            "     Result = clsODCSEscript.getLocalObjectFromXML(objXml_tmp.xml)  " & vbCrLf &
        '            "  end if " & vbCrLf &
        '            "    local = Result    " & vbCrLf &
        '            "End Function"

        '    vScriptControl.AddCode floc


        ''Database Object
        ''vExecuteQuery = "Function ExecuteQuery(vSql)" & vbCrLf & _
        '            "   dim rst " & vbCrLf & _
        '            "    clsODCSEscript.vSql = vSql " & vbCrLf & _
        '            "    set rst = clsODCSEscript.ExecuteQuery("""",dbCon) " & vbCrLf & _
        '            "    set ExecuteQuery = rst " & vbCrLf & _
        '            "End Function"

        ' vExecuteQuery = "Function ExecuteQuery(vSql)" & vbCrLf &
        '            "   dim rst " & vbCrLf &
        '            "    clsODCSEscript.vSql = vSql " & vbCrLf &
        '            "    clsODCSEscript.DTS_location = """ & vDtsLocation & """ " & vbCrLf &
        '            "    clsODCSEscript.DTS_mode = " & vDtsMode & " " & vbCrLf &
        '            "    set rst = clsODCSEscript.ExecuteQuery("""",dbCon) " & vbCrLf &
        '            "    set ExecuteQuery = rst " & vbCrLf &
        '            "End Function"

        '    vScriptControl.AddCode vExecuteQuery

        ''vExecuteNonQuery = "Function ExecuteNonQuery(vSql)" & vbCrLf & _
        '            "    clsODCSEscript.vSql = vSql " & vbCrLf & _
        '            "    clsODCSEscript.ExecuteNonQuery """",dbCon " & vbCrLf & _
        '            "    ExecuteNonQuery=true " & vbCrLf & _
        '            "End Function"

        'vExecuteNonQuery = "Function ExecuteNonQuery(vSql)" & vbCrLf &
        '            "    clsODCSEscript.vSql = vSql " & vbCrLf &
        '            "    clsODCSEscript.DTS_location = """ & vDtsLocation & """ " & vbCrLf &
        '            "    clsODCSEscript.DTS_mode = " & vDtsMode & " " & vbCrLf &
        '            "    clsODCSEscript.ExecuteNonQuery """",dbCon " & vbCrLf &
        '            "    ExecuteNonQuery=true " & vbCrLf &
        '            "End Function"

        '    vScriptControl.AddCode vExecuteNonQuery


        '    vBeginTrans = "Function BeginTrans()" & vbCrLf &
        '            "    clsODCSEscript.BeginTrans dbCon " & vbCrLf &
        '            "    clsODCSEscript.DTS_location = """ & vDtsLocation & """ " & vbCrLf &
        '            "    clsODCSEscript.DTS_mode = " & vDtsMode & " " & vbCrLf &
        '            "    BeginTrans=true " & vbCrLf &
        '            "End Function"
        '    vScriptControl.AddCode vBeginTrans

        'vCommit = "Function Commit()" & vbCrLf &
        '            "    clsODCSEscript.commit dbCon " & vbCrLf &
        '            "    clsODCSEscript.DTS_location = """ & vDtsLocation & """ " & vbCrLf &
        '            "    clsODCSEscript.DTS_mode = " & vDtsMode & " " & vbCrLf &
        '            "    Commit=true " & vbCrLf &
        '            "End Function"
        '    vScriptControl.AddCode vCommit

        'vRollback = "Function rollback()" & vbCrLf &
        '            "    clsODCSEscript.rollback dbCon " & vbCrLf &
        '            "    clsODCSEscript.DTS_location = """ & vDtsLocation & """ " & vbCrLf &
        '            "    clsODCSEscript.DTS_mode = " & vDtsMode & " " & vbCrLf &
        '            "    rollback=true " & vbCrLf &
        '            "End Function"
        '    vScriptControl.AddCode vRollback

        ''Return
        '    vReturn_ = "Function return(vReturn_) " & vbCrLf &
        '            " dim output_" & vbCrLf &
        '            "    clsODCSEscript.vReturn = vReturn_ " & vbCrLf &
        '            "End Function"
        '    vScriptControl.AddCode vReturn_

        'vReturn_ = "Function returndata(vReturn_) " & vbCrLf &
        '            " dim output_" & vbCrLf &
        '            "    output_ = clsODCSEscript.returnData("""") " & vbCrLf &
        '            "   returndata = output_ " & vbCrLf &
        '            "End Function"
        '    vScriptControl.AddCode vReturn_
        ''Parameter
        '    vparameter_ = "Function parametervalue() " & vbCrLf &
        '            " dim value_" & vbCrLf &
        '            "    value_ clsODCSEscript.para   " & vbCrLf &
        '            "   parametervalue = value_ " & vbCrLf &
        '            "End Function"
        '    vScriptControl.AddCode vparameter_
        ''"    clsODCSEscript.vReturn = clsODCSEscript.returnData("""") " & vbCrLf


        '    vShall_ = "Function shell(filename_)" & vbCrLf &
        '            "   clsODCSEscript.vExeFileName = filename_ " & vbCrLf &
        '            "    clsODCSEscript.ShellAndWait("""") " & vbCrLf &
        '            "    shell= true " & vbCrLf &
        '            "End Function"
        '    vScriptControl.AddCode vShall_

        '    If vAllowUI Then
        '        vDSP_ = "Function dsp(dspProfilename_,service_,sn_,so_,product_,station_,pc_,line_)" & vbCrLf &
        '            "   clsODCSEscript.vDSPProfileName = dspProfilename_ " & vbCrLf &
        '            "   clsODCSEscript.vDSPServiceName = service_ " & vbCrLf &
        '            "   clsODCSEscript.vDSPSnName = sn_ " & vbCrLf &
        '            "   clsODCSEscript.vDSPSoName = so_ " & vbCrLf &
        '            "   clsODCSEscript.vDSPStationName = station_ " & vbCrLf &
        '            "   clsODCSEscript.vDSPProductName = product_ " & vbCrLf &
        '            "   clsODCSEscript.vDSPPcName = pc_ " & vbCrLf &
        '            "   clsODCSEscript.vDSPLineName = line_ " & vbCrLf &
        '            "   clsODCSEscript.setCurrentForm clsForm " & vbCrLf &
        '            "    dsp = clsODCSEscript.dspProcess("""","""","""","""",dbcon)  " & vbCrLf &
        '            "End Function"
        '        vScriptControl.AddCode vDSP_
        'Else
        '        vDSP_ = "Function dsp(dspProfilename_,service_,sn_,so_,product_,station_,pc_,line_)" & vbCrLf &
        '            "   clsODCSEscript.vDSPProfileName = dspProfilename_ " & vbCrLf &
        '            "   clsODCSEscript.vDSPServiceName = service_ " & vbCrLf &
        '            "   clsODCSEscript.vDSPSnName = sn_ " & vbCrLf &
        '            "   clsODCSEscript.vDSPSoName = so_ " & vbCrLf &
        '            "   clsODCSEscript.vDSPStationName = station_ " & vbCrLf &
        '            "   clsODCSEscript.vDSPProductName = product_ " & vbCrLf &
        '            "   clsODCSEscript.vDSPPcName = pc_ " & vbCrLf &
        '            "   clsODCSEscript.vDSPLineName = line_ " & vbCrLf &
        '            "    dsp = clsODCSEscript.dspProcess("""","""","""","""",dbcon)  " & vbCrLf &
        '            "End Function"
        '        vScriptControl.AddCode vDSP_
        'End If


        '    If vAllowUI Then
        '        vWS_ = "Function ws(dspProfilename_)" & vbCrLf &
        '                    "   clsODCSEscript.vDSPProfileName = dspProfilename_ " & vbCrLf &
        '                    "   clsODCSEscript.setCurrentForm clsForm " & vbCrLf &
        '                    "    ws = clsODCSEscript.wsProcess(""" & vSN & """,""" & vSO & """,""" & vProduct & ""","""",dbcon)  " & vbCrLf &
        '                    "End Function"
        '    Else
        '        vWS_ = "Function ws(dspProfilename_)" & vbCrLf &
        '                   "   clsODCSEscript.vDSPProfileName = dspProfilename_ " & vbCrLf &
        '                   "    ws = clsODCSEscript.wsProcess(""" & vSN & """,""" & vSO & """,""" & vProduct & ""","""",dbcon)  " & vbCrLf &
        '                   "End Function"
        '    End If

        '    vScriptControl.AddCode vWS_

    End Sub

    Public Function getObject(vControlName As String) As Object  'vControlName=1.text1 (StepName.Control)
        On Error GoTo HasError
        Dim ddd As Control = Nothing
        'ObjectName = vObjName
        Dim vStep() As String

        vStep = Split(ObjectName, ".")

        If UBound(vStep) = 0 Then 'One level
            If Not Form Is Nothing Then
                ddd = Form.Controls(vStep(0))
            End If
        End If

        If UBound(vStep) = 1 Then 'One level
            If Not Form Is Nothing Then
                'ddd = Form.Controls(vStep(0)).object.localControls(vStep(1))
                ddd = Form.Controls(vStep(0)).Controls(vStep(1))
            End If
        End If

        If UBound(vStep) = 2 Then 'One level'Two level
            If Not Form Is Nothing Then
                'ddd = Form.Controls(vStep(0)).object.localControls(vStep(1)).object.localControls(vStep(2))
                ddd = Form.Controls(vStep(0)).Controls(vStep(1)).Controls(vStep(2))
            End If
        End If
        getObject = ddd
        Exit Function
HasError:
        getObject = Nothing

    End Function



End Class