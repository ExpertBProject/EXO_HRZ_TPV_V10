Imports SAPbobsCOM
Imports SAPbouiCOM
Public Class EXO_Apertura
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
                    Case "EXOAPERTURA"
                        Select Case infoEvento.MenuUID
                            Case "1283"
                                'Por defecto el dia de hoy al añadir
#Region "Si ya se ha cerrado, no se puede eliminar"
                                Dim cFechCierre As String = oForm.DataSources.DBDataSources.Item("@EXO_APERCIERRE").GetValue("U_EXO_FechaCierre", 0).Trim()
                                If cFechCierre <> "" Then
                                    objGlobal.SBOApp.MessageBox("No puede eliminar este registro\nYa ha realizado el cierre de caja", 1, "Si", "", "")
                                    Return False
                                End If
#End Region
                                If objGlobal.SBOApp.MessageBox("¿ Este seguro de eliminar este registro ?", 1, "Si", "No", "") <> 1 Then
                                    Return False
                                End If
                        End Select
                End Select
            Else
                Select Case oForm.TypeEx
                    Case "EXOAPERTURA"
                        Select Case infoEvento.MenuUID
                            Case "1282"
                                'Por defecto el dia de hoy al añadir
                                CType(oForm.Items.Item("txtFecAp").Specific, SAPbouiCOM.EditText).Value = DateTime.Today.ToString("yyyyMMdd")

                        End Select
                End Select
                Select Case infoEvento.MenuUID
                    Case "mAperCaja"
#Region "CargoScreen"
                        Dim oParametrosCreacion As SAPbouiCOM.FormCreationParams = CType((objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)), FormCreationParams)

                        Dim strXML As String = objGlobal.leerEmbebido(Me.GetType(), "sEXO_Apertura.srf")
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
                        oForm.Items.Item("txtFecAp").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
                        oForm.Items.Item("txtSalIni").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)

                        oForm.Visible = True
                        INICIO._lgHayqueSalir = False
                        If oForm.Mode = BoFormMode.fm_ADD_MODE Then
                            CType(oForm.Items.Item("txtFecAp").Specific, SAPbouiCOM.EditText).Value = DateTime.Today.ToString("yyyyMMdd")
                        End If
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
                        Case "EXOAPERTURA"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXOAPERTURA"
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
                        Case "EXOAPERTURA"
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
                        Case "EXOAPERTURA"
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
        Dim sSQL As String = ""
        EventHandler_Form_Visible = False
        Dim dFecha As Date = New Date(Now.Year, Now.Month, Now.Day)
        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                'Cargamos combo de almacén
                If CargaComboAlm(oForm) = False Then
                    Return False
                Else
                    'Ponemos el valor por defecto del usuario.
                    sSQL = "SELECT OUDG.""Warehouse"" FROM OUSR "
                    sSQL &= " INNER JOIN OUDG ON OUDG.""Code""=OUSR.""DfltsGroup"" "
                    sSQL &= " WHERE ""USERID""='" & objGlobal.compañia.UserSignature & "'" 'Buscamos el valor que tienen las opciones de usuario
                    Dim sALM As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    If sALM.Trim <> "" Then
                        CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Select(sALM.Trim, BoSearchKey.psk_ByValue)
                        oForm.Items.Item("cbALM").DisplayDesc = True
                    End If
                End If
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
    Private Function CargaComboAlm(ByRef oForm As SAPbouiCOM.Form) As Boolean
        CargaComboAlm = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try

            sSQL = "SELECT ""WhsCode"" ""Código"", ""WhsName"" ""Almacén"" FROM OWHS Order BY ""WhsName"" "
            oRs.DoQuery(sSQL)

            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

            CargaComboAlm = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "1" Then
                If pVal.ActionSuccess = True Then
                    If INICIO._lgHayqueSalir = True Then
                        INICIO._lgHayqueSalir = False
                        oForm.Close()
                    End If
                End If
            End If
            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument

        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "EXOAPERTURA"
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
                    Case "EXOAPERTURA"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess = True Then
                                    INICIO._lgHayqueSalir = True
                                    objGlobal.SBOApp.MessageBox("Apertura realizada", 1, "Ok", "", "")
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

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
    Private Function Validar(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim sSQL As String = ""
        Validar = False
        Dim cFechApertura As String = ""
        Dim cAlmacen As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
#Region "Obligatorio fecha apertura"
            cFechApertura = oForm.DataSources.DBDataSources.Item("@EXO_APERCIERRE").GetValue("U_EXO_FechaAper", 0).Trim()
            If cFechApertura = "" Then
                objGlobal.SBOApp.MessageBox("Ha de introducir Fecha de Apertura", 2, "Si", "No", "")
                Return False
            End If
#End Region
#Region "Control si ya está abierto en el mismo almacén y mismo día"
            cFechApertura = oForm.DataSources.DBDataSources.Item("@EXO_APERCIERRE").GetValue("U_EXO_FechaAper", 0).Trim()
            cAlmacen = oForm.DataSources.DBDataSources.Item("@EXO_APERCIERRE").GetValue("U_EXO_ALM", 0).Trim()
            If cFechApertura <> "" And cAlmacen <> "" Then
                sSQL = "SELECT * FROM ""@EXO_APERCIERRE"" WHERE ""U_EXO_ALM""='" & cAlmacen & "' "
                sSQL &= " And ifnull(""U_EXO_FechaCierre"",'19000101')='19000101' and ""U_EXO_FechaAper""='" & cFechApertura & "'"
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    objGlobal.SBOApp.MessageBox("No se puede abrir dos veces la caja el mismo día.",)
                    oForm.Mode = BoFormMode.fm_FIND_MODE
                    CType(oForm.Items.Item("txtDoc").Specific, SAPbouiCOM.EditText).Value = oRs.Fields.Item("DocEntry").Value.ToString
                    oForm.Items.Item("1").Click()
                    Exit Function
                End If

            End If
#End Region

#Region "Aviso si el dinero inicial es 0"
            Dim cAux As String = oForm.DataSources.DBDataSources.Item("@EXO_APERCIERRE").GetValue("U_EXO_SaldoIncial", 0).Trim()
            If EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, cAux) = 0 Then
                If objGlobal.SBOApp.MessageBox("!! ATENCION !!\nEl dinero inicial es 0\n ¿ Continuar ?", 2, "Si", "No", "") <> 1 Then
                    Exit Function
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
End Class
