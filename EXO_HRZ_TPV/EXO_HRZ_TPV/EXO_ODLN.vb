Imports SAPbobsCOM
Imports SAPbouiCOM
Public Class EXO_ODLN
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
                        Case "140"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_COMBO_SELECT_After(infoEvento) = False Then
                                        Return False
                                    End If
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
                        Case "140"
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
                        Case "140"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "140"
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
    Private Function EventHandler_Form_Load(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oItem As SAPbouiCOM.Item

        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            oForm.Visible = False

            'Buscar XML de update
            objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
#Region "Campos identificación del cobro"
            oItem = oForm.Items.Add("lblCobro", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("89").Top
            oItem.Left = oForm.Items.Item("230").Left
            oItem.Height = oForm.Items.Item("230").Height
            oItem.Width = oForm.Items.Item("230").Width
            oItem.LinkTo = "222"
            oItem.FromPane = 0
            oItem.ToPane = 0
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Referencia Cobro: "
            oItem.TextStyle = 1


            oItem = oForm.Items.Add("txtCDEntry", BoFormItemTypes.it_EXTEDIT)
            oItem.Top = oForm.Items.Item("103").Top
            oItem.Left = oForm.Items.Item("222").Left
            oItem.Height = oForm.Items.Item("222").Height
            oItem.Width = oForm.Items.Item("222").Width
            oItem.LinkTo = "lblCobro"
            oItem.FromPane = 0
            oItem.ToPane = 0
            CType(oItem.Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "ODLN", "U_EXO_CDOCENTRY")
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            oItem = oForm.Items.Add("IrPago", BoFormItemTypes.it_LINKED_BUTTON)
            oItem.Top = oForm.Items.Item("txtCDEntry").Top  'Incidencia
            oItem.Left = oForm.Items.Item("229").Left
            oItem.Height = oForm.Items.Item("87").Height
            oItem.Width = oForm.Items.Item("87").Width
            oItem.LinkTo = "txtCDEntry"
            oItem.FromPane = 0
            oItem.ToPane = 0
            CType(oItem.Specific, SAPbouiCOM.LinkedButton).LinkedObjectType = CType(BoObjectTypes.oIncomingPayments, String)
            CType(oItem.Specific, SAPbouiCOM.LinkedButton).Item.LinkTo = "txtCDEntry"


            oItem = oForm.Items.Add("lblCDEntry", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("txtCDEntry").Top
            oItem.Left = oForm.Items.Item("230").Left
            oItem.Height = oForm.Items.Item("230").Height
            oItem.Width = oForm.Items.Item("230").Width
            oItem.LinkTo = "txtCDEntry"
            oItem.FromPane = 0
            oItem.ToPane = 0
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Nº Interno"

            oItem = oForm.Items.Add("txtCDNum", BoFormItemTypes.it_EDIT)
            oItem.Top = oForm.Items.Item("27").Top
            oItem.Left = oForm.Items.Item("222").Left
            oItem.Height = oForm.Items.Item("222").Height
            oItem.Width = oForm.Items.Item("222").Width
            oItem.LinkTo = "lblCobro"
            oItem.FromPane = 0
            oItem.ToPane = 0
            oItem.Enabled = False
            CType(oItem.Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "ODLN", "U_EXO_CDOCNUM")
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            oItem = oForm.Items.Add("lblCDNum", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("txtCDNum").Top
            oItem.Left = oForm.Items.Item("230").Left
            oItem.Height = oForm.Items.Item("230").Height
            oItem.Width = oForm.Items.Item("230").Width
            oItem.LinkTo = "txtCDNum"
            oItem.FromPane = 0
            oItem.ToPane = 0
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Nº Cobro"

            oItem = oForm.Items.Add("txtCTipo", BoFormItemTypes.it_COMBO_BOX)
            oItem.Top = oForm.Items.Item("29").Top
            oItem.Left = oForm.Items.Item("222").Left
            oItem.Height = oForm.Items.Item("222").Height
            oItem.Width = oForm.Items.Item("222").Width
            oItem.LinkTo = "lblCobro"
            oItem.FromPane = 0
            oItem.ToPane = 0
            oItem.DisplayDesc = True
            oItem.Enabled = False
            CType(oItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "ODLN", "U_EXO_CTIPO")
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            oItem = oForm.Items.Add("lblCTipo", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("txtCTipo").Top
            oItem.Left = oForm.Items.Item("230").Left
            oItem.Height = oForm.Items.Item("230").Height
            oItem.Width = oForm.Items.Item("230").Width
            oItem.LinkTo = "txtCTipo"
            oItem.FromPane = 0
            oItem.ToPane = 0
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Tipo"
#End Region
#Region "Botones"
            oItem = oForm.Items.Add("btnCOBROT", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oForm.Items.Item("10000330").Left - (oForm.Items.Item("10000330").Width * 2) + 50
            oItem.Width = (oForm.Items.Item("10000330").Width * 2) - 30
            oItem.Top = oForm.Items.Item("46").Top + 25
            oItem.Height = oForm.Items.Item("2").Height
            oItem.Enabled = False
            Dim oBtnAct As SAPbouiCOM.Button
            oBtnAct = CType(oItem.Specific, Button)
            oBtnAct.Caption = "Pago Total Tarjeta"
            oItem.TextStyle = 1
            oItem.LinkTo = "46"
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            oItem = oForm.Items.Add("btnCOBROC", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oForm.Items.Item("btnCOBROT").Left - oForm.Items.Item("btnCOBROT").Width - 2
            oItem.Width = oForm.Items.Item("btnCOBROT").Width
            oItem.Top = oForm.Items.Item("btnCOBROT").Top
            oItem.Height = oForm.Items.Item("btnCOBROT").Height
            oItem.Enabled = False
            oBtnAct = CType(oItem.Specific, Button)
            oBtnAct.Caption = "Pago Total Caja"
            oItem.TextStyle = 1
            oItem.LinkTo = "btnCOBROT"
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            oItem = oForm.Items.Add("btnCOBROCN", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oForm.Items.Item("btnCOBROT").Left + oForm.Items.Item("btnCOBROT").Width + 2
            oItem.Width = oForm.Items.Item("btnCOBROT").Width
            oItem.Top = oForm.Items.Item("btnCOBROT").Top
            oItem.Height = oForm.Items.Item("btnCOBROT").Height
            oItem.Enabled = False
            oBtnAct = CType(oItem.Specific, Button)
            oBtnAct.Caption = "Cancelar Pago Asociado"
            oItem.TextStyle = 1
            oItem.LinkTo = "btnCOBROT"
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
#End Region
#Region "TEXTO TIPO ALBARAN"
            oItem = oForm.Items.Add("lblTIPOA", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("btnCOBROT").Top
            oItem.Left = oForm.Items.Item("15").Left
            oItem.Height = (oForm.Items.Item("230").Height * 2)
            oItem.Width = (oForm.Items.Item("230").Width * 2) - 20
            oItem.LinkTo = "15"
            oItem.FromPane = 0
            oItem.ToPane = 0
            oItem.AffectsFormMode = False
            Dim sSerie As String = ""
            If CType(oForm.Items.Item("88").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                sSerie = CType(oForm.Items.Item("88").Specific, SAPbouiCOM.ComboBox).Selected.Description.ToString
            End If

            Select Case Left(sSerie, 2)
                Case "TQ" : CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "TICKET VENTA"
                Case Else : CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "ALBARÁN VENTA"
            End Select

            oItem.TextStyle = 1 : oItem.FontSize = 18
#End Region
            oForm.Visible = True

            EventHandler_Form_Load = True

        Catch ex As Exception
            oForm.Visible = True
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btnCOBROT"
                    If pVal.ActionSuccess = True Then
                        If CargarFormCOBROT(oForm, "V") = False Then
                            Exit Function
                        End If
                    End If
                Case "btnCOBROC"
                    If pVal.ActionSuccess = True Then
                        If CargarFormCOBROT(oForm, "C") = False Then
                            Exit Function
                        End If
                    End If
                Case "btnCOBROCN"
                    If pVal.ActionSuccess = True Then
                        If Cancelar_Cobro(oForm) = False Then
                            Exit Function
                        End If
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Function CargarFormCOBROT(ByRef oFormODLN As SAPbouiCOM.Form, ByVal sTipo As String) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing

        CargarFormCOBROT = False

        Try
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_COBROT.srf")
            oFP.XmlData = oFP.XmlData.Replace("modality=""0""", "modality=""1""")
            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try
            oForm.Left = oFormODLN.Left
            'Damos valores a los campos para generar el pago-cobro
            oForm.DataSources.UserDataSources.Item("UDTIPO").ValueEx = sTipo
            oForm.DataSources.UserDataSources.Item("UDCARDCODE").ValueEx = oFormODLN.DataSources.DBDataSources.Item("ODLN").GetValue("CardCode", 0).ToString
            oForm.DataSources.UserDataSources.Item("UDDOCENTRY").ValueEx = oFormODLN.DataSources.DBDataSources.Item("ODLN").GetValue("DocEntry", 0).ToString
            oForm.DataSources.UserDataSources.Item("UDDOCNUM").ValueEx = oFormODLN.DataSources.DBDataSources.Item("ODLN").GetValue("DocNum", 0).ToString
            oForm.DataSources.UserDataSources.Item("UDIMP").ValueEx = oFormODLN.DataSources.DBDataSources.Item("ODLN").GetValue("DocTotal", 0).ToString
            CType(oForm.Items.Item("lblDOCNUM").Specific, SAPbouiCOM.StaticText).Item.TextStyle = 1
            CType(oForm.Items.Item("txtDOCNUM").Specific, SAPbouiCOM.EditText).Item.TextStyle = 1
            CType(oForm.Items.Item("txtDOCNUM").Specific, SAPbouiCOM.EditText).Item.AffectsFormMode = False
            CType(oForm.Items.Item("lblTIPO").Specific, SAPbouiCOM.StaticText).Item.TextStyle = 1
            CType(oForm.Items.Item("cbTIPO").Specific, SAPbouiCOM.ComboBox).Item.TextStyle = 1
            CType(oForm.Items.Item("cbTIPO").Specific, SAPbouiCOM.ComboBox).Item.AffectsFormMode = False

            CType(oForm.Items.Item("txtIMP").Specific, SAPbouiCOM.EditText).Item.TextStyle = 1
            CType(oForm.Items.Item("txtIMP").Specific, SAPbouiCOM.EditText).Item.AffectsFormMode = False
            CType(oForm.Items.Item("txtCAM").Specific, SAPbouiCOM.EditText).Item.TextStyle = 1
            CType(oForm.Items.Item("txtCAM").Specific, SAPbouiCOM.EditText).Item.AffectsFormMode = False
            CType(oForm.Items.Item("txtCLI").Specific, SAPbouiCOM.EditText).Item.AffectsFormMode = False


            CargarFormCOBROT = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Function Cancelar_Cobro(ByRef oFormODLN As SAPbouiCOM.Form) As Boolean
        Dim sDocEntryCobro As String = oFormODLN.DataSources.DBDataSources.Item("ODLN").GetValue("U_EXO_CDOCENTRY", 0).ToString
        Dim sMensaje As String = ""
        Dim ORCT As SAPbobsCOM.Payments = Nothing
        Cancelar_Cobro = False

        Try
            If sDocEntryCobro <> "" Then
                If objGlobal.SBOApp.MessageBox("¿Está seguro que quiere cancelar el cobro asociado?", 1, "Sí", "No") = 1 Then
                    ORCT = CType(objGlobal.compañia.GetBusinessObject(BoObjectTypes.oIncomingPayments), SAPbobsCOM.Payments)
                    If ORCT.GetByKey(CType(sDocEntryCobro, Integer)) = True Then
                        ORCT.CancelbyCurrentSystemDate()
                        If ORCT.Update() <> 0 Then
                            objGlobal.SBOApp.StatusBar.SetText("No se ha podico cancelar el cobro asociado - " & objGlobal.compañia.GetLastErrorCode & " / " & objGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                            Exit Function
                        Else
                            sMensaje = "Se ha cancelado correctamente el cobro asociado."
                            objGlobal.SBOApp.StatusBar.SetText(sMensaje, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                            objGlobal.SBOApp.MessageBox(sMensaje)
                        End If
                    Else
                        sMensaje = "Error grave. No se ha encontrado el cobro asociado."
                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                        objGlobal.SBOApp.MessageBox(sMensaje)
                    End If
                End If
            Else
                sMensaje = "No existe ningún cobro asociado."
                objGlobal.SBOApp.StatusBar.SetText(sMensaje, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox(sMensaje)
            End If


            Cancelar_Cobro = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            If ORCT IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ORCT)
            ORCT = Nothing
        End Try
    End Function
    Private Function EventHandler_COMBO_SELECT_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_COMBO_SELECT_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "88"
                    If oForm.Visible = True Then
                        Dim sSerie As String = "" : Dim oItem As SAPbouiCOM.Item
                        If CType(oForm.Items.Item("88").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                            sSerie = CType(oForm.Items.Item("88").Specific, SAPbouiCOM.ComboBox).Selected.Description.ToString
                        End If
                        oItem = oForm.Items.Item("lblTIPOA")
                        Select Case Left(sSerie, 2)
                            Case "TQ" : CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "TICKET VENTA"
                            Case Else : CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "ALBARÁN VENTA"
                        End Select
                    End If
            End Select

            EventHandler_COMBO_SELECT_After = True

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
        Dim bEstado As String = ""
        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "140"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                'Revisamos la serie
                                Return Revisar_Serie(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "140"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If oForm.Visible = True Then
                                    Dim sSerie As String = "" : Dim oItem As SAPbouiCOM.Item
                                    If CType(oForm.Items.Item("88").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                                        sSerie = CType(oForm.Items.Item("88").Specific, SAPbouiCOM.ComboBox).Selected.Description.ToString
                                    End If
                                    oItem = oForm.Items.Item("lblTIPOA")
                                    Select Case Left(sSerie, 2)
                                        Case "TQ" : CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "TICKET VENTA"
                                        Case Else : CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "ALBARÁN VENTA"
                                    End Select
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select
                    Case "140"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD


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
    Private Function Revisar_Serie(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim sSQL As String = ""
        Revisar_Serie = False

        Try
            Dim sCardCode As String = "" : Dim sProp As String = "" : Dim sSerie As String = ""
            sCardCode = oForm.DataSources.DBDataSources.Item("ODLN").GetValue("CardCode", 0).ToString.Trim
            sSQL = "SELECT ""QryGroup10"" FROM ""OCRD"" WHERE ""CardCode""='" & sCardCode & "' "
            sProp = objGlobal.refDi.SQL.sqlStringB1(sSQL)

            If CType(oForm.Items.Item("88").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                sSerie = CType(oForm.Items.Item("88").Specific, SAPbouiCOM.ComboBox).Selected.Description.ToString
            End If

            If sProp = "" Then
                objGlobal.SBOApp.StatusBar.SetText("Error grave, no se encuentra el Interlocutor.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objGlobal.SBOApp.MessageBox("Error grave, no se encuentra el Interlocutor.")
                Exit Function
            Else
                Select Case sProp
                    Case "N"
                        If Left(sSerie, 2) <> "TQ" Then
                            Return True
                        Else
                            objGlobal.SBOApp.StatusBar.SetText("Por favor, revise la Serie. No es correcta la actual.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    Case "Y"
                        If Left(sSerie, 2) = "TQ" Then
                            Return True
                        Else
                            objGlobal.SBOApp.StatusBar.SetText("Por favor, revise la Serie. No es correcta la actual.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                End Select
            End If

            Revisar_Serie = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Function
End Class
