Imports SAPbobsCOM
Imports SAPbouiCOM
Public Class EXO_COBROTF

    Private objGlobal As EXO_UIAPI.EXO_UIAPI

    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            'Apaño por un error que da EXO_Basic.dll al consultar infoEvento.FormTypeEx
            Try
                If infoEvento.FormTypeEx <> "" Then

                End If
            Catch ex As Exception
                Return False
            End Try

            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_COBROTF"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_VALIDATE_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_COBROTF"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_Before(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_COBROTF"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_Form_Visible(objGlobal, infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_COBROTF"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                End If
            End If

            Return True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_Form_Visible(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        EventHandler_Form_Visible = False
        Dim dFecha As Date = New Date(Now.Year, Now.Month, Now.Day)
        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                oForm.DataSources.UserDataSources.Item("UDFECHA").Value = Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00")
                Dim sTipo As String = oForm.DataSources.UserDataSources.Item("UDTIPO").Value.ToString
                Select Case sTipo
                    Case "C"
                    Case "V"
                        oForm.DataSources.UserDataSources.Item("UDCLI").ValueEx = oForm.DataSources.UserDataSources.Item("UDIMP").ValueEx
                End Select
            End If

            EventHandler_Form_Visible = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "1"
                    Return Generar_Cobro(oForm)
                    objGlobal.SBOApp.ActivateMenuItem("1044")
                    objGlobal.SBOApp.ActivateMenuItem("1304")
            End Select

            EventHandler_ItemPressed_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_VALIDATE_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_VALIDATE_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "txtCLI"
                    Dim dUDCLI As Double = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oForm.DataSources.UserDataSources.Item("UDCLI").Value)
                    Dim dUDIMP As Double = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oForm.DataSources.UserDataSources.Item("UDIMP").Value)
                    oForm.DataSources.UserDataSources.Item("UDCAM").ValueEx = EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dUDCLI - dUDIMP, EXO_GLOBALES.FuenteInformacion.Otros)

            End Select

            EventHandler_VALIDATE_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function Generar_Cobro(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim ORCT As SAPbobsCOM.Payments = Nothing
        Dim sDocEntryORCT As String = ""
        Dim sDocNumORCT As String = ""
        Dim sSQL As String = ""
        Dim sTIPO As String = ""
        Dim sAccount As String = ""
        Dim sLinVto As String = ""
        Generar_Cobro = False
        Try
            sTIPO = oForm.DataSources.UserDataSources.Item("UDTIPO").Value.ToString

            ORCT = CType(objGlobal.compañia.GetBusinessObject(BoObjectTypes.oIncomingPayments), SAPbobsCOM.Payments)
            ORCT.CardCode = oForm.DataSources.UserDataSources.Item("UDCARDCODE").Value.ToString
            ORCT.DocType = BoRcptTypes.rCustomer
            ORCT.CashSum = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oForm.DataSources.UserDataSources.Item("UDIMP").Value.ToString)
            Select Case sTIPO
                Case "C" : sSQL = "SELECT ""OUDG"".""CashAcct"" FROM ""OUSR"" INNER JOIN ""OUDG"" ON ""OUDG"".""Code""=""OUSR"".""DfltsGroup"" WHERE ""USER_CODE""='" & objGlobal.compañia.UserName & "' "
                Case "V" : sSQL = "SELECT ""OUDG"".""CheckAcct"" FROM ""OUSR"" INNER JOIN ""OUDG"" ON ""OUDG"".""Code""=""OUSR"".""DfltsGroup"" WHERE ""USER_CODE""='" & objGlobal.compañia.UserName & "' "
            End Select
            sAccount = objGlobal.refDi.SQL.sqlStringB1(sSQL)
            ORCT.CashAccount = sAccount
            Dim sFecha As String = oForm.DataSources.UserDataSources.Item("UDFECHA").ValueEx.ToString
            Dim dFecha As Date = New Date(CType(Left(sFecha, 4), Integer), CType(Mid(sFecha, 5, 2), Integer), CType(Right(sFecha, 2), Integer))
            Select Case sTIPO
                Case "C"
                    ORCT.CounterReference = "CAJA"
                Case "V"
                    ORCT.CounterReference = "VISA"
            End Select
            ORCT.DocDate = dFecha
            ORCT.TransferDate = dFecha
            ORCT.Remarks = "Factura Nº" & oForm.DataSources.UserDataSources.Item("UDDOCNUM").Value.ToString
            ORCT.UserFields.Fields.Item("U_EXO_SERIE").Value = oForm.DataSources.UserDataSources.Item("UDSERIE").Value.ToString
            ORCT.Invoices.SumApplied = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oForm.DataSources.UserDataSources.Item("UDIMP").Value.ToString)
            ORCT.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
            ORCT.Invoices.DocEntry = CInt(oForm.DataSources.UserDataSources.Item("UDDOCENTRY").Value.ToString)
            sSQL = "SELECT TOP 1 ""InstlmntID""  FROM INV6 T0 WHERE ""DocEntry""=" & oForm.DataSources.UserDataSources.Item("UDDOCENTRY").Value.ToString & " And ""ObjType"" ='13' and ""Status""='O' "
            sLinVto = objGlobal.refDi.SQL.sqlStringB1(sSQL)
            ORCT.Invoices.InstallmentId = CInt(sLinVto)
            ORCT.Invoices.Add()
            If sAccount <> "" Then
                If ORCT.Add() = 0 Then
                    objGlobal.compañia.GetNewObjectCode(sDocEntryORCT)
                    objGlobal.SBOApp.StatusBar.SetText("Creado cobro. Se procede a actualizar la factura...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    sSQL = "Select ""DocNum"" FROM ""ORCT"" WHERE ""DocEntry""=" & sDocEntryORCT
                    sDocNumORCT = objGlobal.refDi.SQL.sqlStringB1(sSQL)

                    sSQL = "UPDATE OINV Set "
                    sSQL &= " ""U_EXO_CDOCENTRY""='" & sDocEntryORCT & "', "
                    sSQL &= " ""U_EXO_CDOCNUM""='" & sDocNumORCT & "', "
                    sSQL &= " ""U_EXO_CTIPO""='" & sTIPO & "' "
                    sSQL &= " WHERE ""DocEntry""= " & oForm.DataSources.UserDataSources.Item("UDDOCENTRY").Value.ToString
                    If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                        objGlobal.SBOApp.StatusBar.SetText("Actualizada la factura Nº" & oForm.DataSources.UserDataSources.Item("UDDOCNUM").Value.ToString, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("No se ha podido actualizar la factura Nº" & oForm.DataSources.UserDataSources.Item("UDDOCNUM").Value.ToString, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                    End If
                    Generar_Cobro = True
                Else
                    objGlobal.SBOApp.MessageBox("Error generando cobro. Por favor realícelo de forma manual: " + objGlobal.compañia.GetLastErrorDescription)
                    Generar_Cobro = False
                End If
            Else
                objGlobal.SBOApp.StatusBar.SetText("No ha definido una cuenta para esta operación. Por favor, revise la parametrización", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                Generar_Cobro = False
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If ORCT IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ORCT)
            ORCT = Nothing
        End Try
    End Function
End Class
