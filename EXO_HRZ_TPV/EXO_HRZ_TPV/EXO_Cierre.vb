Imports SAPbobsCOM
Imports SAPbouiCOM
Public Class EXO_Cierre
    Private objGlobal As EXO_UIAPI.EXO_UIAPI

    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sLineNum As String = "0"

        Try
            oForm = objGlobal.SBOApp.Forms.ActiveForm
            If infoEvento.BeforeAction = True Then
                Select Case oForm.TypeEx
                    Case "EXOCIERRE"
                        Select Case infoEvento.MenuUID
                            Case ""

                        End Select
                End Select
            Else
                Select Case infoEvento.MenuUID
                    Case "mCierreCaja"
#Region "CargoScreen"
                        Dim oParametrosCreacion As SAPbouiCOM.FormCreationParams = CType((objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)), FormCreationParams)

                        Dim strXML As String = objGlobal.leerEmbebido(Me.GetType(), "sEXO_Cierre.srf")
                        oParametrosCreacion.XmlData = strXML
                        oParametrosCreacion.UniqueID = ""

                        Try
                            oForm = objGlobal.SBOApp.Forms.AddEx(oParametrosCreacion)
                        Catch ex As Exception
                            objGlobal.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "")
                        End Try
#End Region
                        oForm.DataBrowser.BrowseBy = "txtDoc"

                        oForm.Items.Item("txtDoc").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)

                        oForm.EnableMenu("1282", False)
                        oForm.Visible = True

                        objGlobal.SBOApp.ActivateMenuItem("1291")
                        oForm.ActiveItem = CType((IIf(oForm.Items.Item("txtFecCie").Enabled, "txtFecCie", "edEdit")), String)
                End Select
            End If

            Return True

        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
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
                        Case "EXOCIERRE"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_Validate_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXOCIERRE"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXOCIERRE"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXOCIERRE"
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
    Private Function EventHandler_Validate_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_Validate_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            oForm.Freeze(True)

            If pVal.ItemUID = "txtFecCie" Then
                If pVal.ItemChanged = True Then
                    PintoValores(oForm)
                End If
            End If

            EventHandler_Validate_After = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)

            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Sub PintoValores(ByRef oForm As SAPbouiCOM.Form)

#Region "Valores de venta y efectivo "
        Dim cFechApertura As String = oForm.DataSources.DBDataSources.Item("@EXO_APERCIERRE").GetValue("U_EXO_FechaAper", 0).Trim()
        Dim cFechCierre As String = oForm.DataSources.DBDataSources.Item("@EXO_APERCIERRE").GetValue("U_EXO_FechaCierre", 0).Trim()

        Dim cAux As String = oForm.DataSources.DBDataSources.Item("@EXO_APERCIERRE").GetValue("U_EXO_SaldoIncial", 0).Trim()
        Dim nSaldoInicial As Double = 0
        If cAux = "" Then
            nSaldoInicial = 0
        Else
            nSaldoInicial = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, cAux)
        End If
        If (cFechCierre <> "") Then
            Dim Sql As String = "SELECT SUM(T0.Total)FROM dbo.EXO_VentasTPV(convert(DATETIME, '##DESDEFECHA', 112), convert(DATETIME, '##HASTAFECHA', 112)) T0"
            Sql = Sql.Replace("##DESDEFECHA", cFechApertura).Replace("##HASTAFECHA", cFechCierre)
            oForm.DataSources.UserDataSources.Item("dsVentas").ValueEx = EXO_GLOBALES.DblNumberToText(objGlobal.compañia, objGlobal.refDi.SQL.sqlNumericaB1(Sql), EXO_GLOBALES.FuenteInformacion.Otros)

            Sql = "SELECT SUM(T0.Efectivo) FROM dbo.EXO_CobrosEfectivoTPV(CONVERT(DATETIME, '##DESDEFECHA', 112), CONVERT(DATETIME, '##HASTAFECHA', 112)) T0"
            Sql = Sql.Replace("##DESDEFECHA", cFechApertura).Replace("##HASTAFECHA", cFechCierre)
            Dim nEfectivo As Double = objGlobal.refDi.SQL.sqlNumericaB1(Sql)
            oForm.DataSources.UserDataSources.Item("dsEfect").ValueEx = EXO_GLOBALES.DblNumberToText(objGlobal.compañia, nEfectivo, EXO_GLOBALES.FuenteInformacion.Otros)

            Dim nSaldoTeorico As Double = nSaldoInicial + nEfectivo
            oForm.DataSources.UserDataSources.Item("dsTeor").ValueEx = EXO_GLOBALES.DblNumberToText(objGlobal.compañia, nSaldoTeorico, EXO_GLOBALES.FuenteInformacion.Otros)
        Else
            oForm.DataSources.UserDataSources.Item("dsVentas").Value = EXO_GLOBALES.DblNumberToText(objGlobal.compañia, 0, EXO_GLOBALES.FuenteInformacion.Otros)
            oForm.DataSources.UserDataSources.Item("dsEfect").Value = EXO_GLOBALES.DblNumberToText(objGlobal.compañia, 0, EXO_GLOBALES.FuenteInformacion.Otros)
            oForm.DataSources.UserDataSources.Item("dsTeor").Value = EXO_GLOBALES.DblNumberToText(objGlobal.compañia, 0, EXO_GLOBALES.FuenteInformacion.Otros)
        End If
#End Region

    End Sub
    Public Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument

        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "EXOCIERRE"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If Validar(oForm) = False Then
                                    Return False
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If Validar(oForm) = False Then
                                    Return False
                                End If
                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "EXOCIERRE"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If infoEvento.ActionSuccess = True Then

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If Habilitar_Deshabilitar_Campos_Cierre(oForm) = False Then
                                    Return False
                                End If
                        End Select
                End Select
            End If

            Return True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function Habilitar_Deshabilitar_Campos_Cierre(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim sSQL As String = ""
        Habilitar_Deshabilitar_Campos_Cierre = False

        Try
#Region "Habilito / Deshabilito campos de cierre"
            If (oForm.Visible = True) Then
                Dim cFechCierre As String = oForm.DataSources.DBDataSources.Item("@EXO_APERCIERRE").GetValue("U_EXO_FechaCierre", 0).Trim()

                oForm.Items.Item("txtFecCie").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, CType(IIf(cFechCierre = "", BoModeVisualBehavior.mvb_True, BoModeVisualBehavior.mvb_False), BoModeVisualBehavior))
                oForm.Items.Item("txtSalFin").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, CType(IIf(cFechCierre = "", BoModeVisualBehavior.mvb_True, BoModeVisualBehavior.mvb_False), BoModeVisualBehavior))

                PintoValores(oForm)
            End If
#End Region
            Habilitar_Deshabilitar_Campos_Cierre = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
            Return False
        Catch ex As Exception
            Throw ex
            Return False
        Finally

        End Try
    End Function
    Private Function Validar(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim sSQL As String = ""
        Validar = False

        Try

#Region "Valido fechas"
            Dim cFechCierre As String = oForm.DataSources.DBDataSources.Item("@EXO_APERCIERRE").GetValue("U_EXO_FechaCierre", 0).Trim()
            Dim cFechApertura As String = oForm.DataSources.DBDataSources.Item("@EXO_APERCIERRE").GetValue("U_EXO_FechaAper", 0).Trim()

            If cFechCierre = "" Then
                objGlobal.SBOApp.MessageBox("Ha de introducir Fecha de Cierre", 1, "Si", "", "")
                Return False
            End If

            Dim dFechaCierre As DateTime : Dim dFechApertura As DateTime
            DateTime.TryParseExact(cFechCierre, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, dFechaCierre)
            DateTime.TryParseExact(cFechApertura, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, dFechApertura)

            If (dFechaCierre < dFechApertura) Then
                objGlobal.SBOApp.MessageBox("La Fecha de Cierre ha de ser mayor o igual que la Fecha de Apertura", 1, "Si", "", "")
                Return False
            End If
#End Region

#Region "Aviso si el dinero final es 0"
            Dim cAux As String = oForm.DataSources.DBDataSources.Item("@EXO_APERCIERRE").GetValue("U_EXO_SaldoFinal", 0).Trim()
            If EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, cAux) = 0 Then
                If (objGlobal.SBOApp.MessageBox("!! ATENCION !!\nEl dinero final es 0\n ¿ Continuar ?", 2, "Si", "No", "") <> 1) Then
                    Return False
                End If
            End If
#End Region
            Validar = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.SBOApp.MessageBox("Error en validacion\n" + exCOM.Message, 1, "Ok", "", "")
            Return False
        Catch ex As Exception
            objGlobal.SBOApp.MessageBox("Error en validacion\n" + ex.Message, 1, "Ok", "", "")
            Return False
        Finally

        End Try
    End Function
    Private Function Desactivar_pantalla(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Desactivar_pantalla = False
        Try
#Region "Para desactivar la pantalla si hace falta"
            Dim cFechCierre As String = oForm.DataSources.DBDataSources.Item("@EXO_APERCIERRE").GetValue("U_EXO_FechaCierre", 0).Trim()

            oForm.Items.Item("txtFecCie").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, CType(IIf(cFechCierre = "", BoModeVisualBehavior.mvb_True, BoModeVisualBehavior.mvb_False), BoModeVisualBehavior))
            oForm.Items.Item("txtSalFin").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, CType(IIf(cFechCierre = "", BoModeVisualBehavior.mvb_True, BoModeVisualBehavior.mvb_False), BoModeVisualBehavior))
#End Region
            Desactivar_pantalla = True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
End Class
