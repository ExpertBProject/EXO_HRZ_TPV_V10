Imports SAPbouiCOM
Imports System.Xml
Imports EXO_UIAPI.EXO_UIAPI

Public Class INICIO
    Inherits EXO_UIAPI.EXO_DLLBase
#Region "Variables globales"
    Public Shared _lgHayqueSalir As Boolean = False
#End Region
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaDatos()
            CargaFuncion()
        End If
        cargamenu()
        InsertarReport()
    End Sub
    Private Sub InsertarReport()
#Region "Variables"
        Dim sArchivo As String = ""
        Dim sMenuId As String = "12800" ' Informes de Ventas
        Dim sReport As String = "" : Dim sNomReport As String = ""
#End Region
        Try
            sArchivo = objGlobal.path & "\05.Rpt\TPV\"
#Region "Report Ventas Diarias"
            sReport = "Ventas Diarias.rpt" : sNomReport = "TPV Ventas Diarias"
            'Si no existe lo importamos
            If IO.File.Exists(sArchivo & sReport) = False Then
                If IO.Directory.Exists(sArchivo) = False Then
                    IO.Directory.CreateDirectory(sArchivo)
                End If
                EXO_GLOBALES.CopiarRecurso(Reflection.Assembly.GetExecutingAssembly(), sReport, sArchivo)
                EXO_GLOBALES.Import_Report(objGlobal, sArchivo & sReport, sMenuId, sNomReport)
                objGlobal.SBOApp.StatusBar.SetText("Importado en Menú: " & sArchivo & sReport & ". Fin Sincronización.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            End If
#End Region
#Region "Report Ventas Diarias Pedidos"
            'sReport = "Ventas Diarias Pedidos.rpt" : sNomReport = "TPV Ventas Diarias Pedidos"
            ''Si no existe lo importamos
            'If IO.File.Exists(sArchivo & sReport) = False Then
            '    EXO_GLOBALES.CopiarRecurso(Reflection.Assembly.GetExecutingAssembly(), sReport, sArchivo)
            '    EXO_GLOBALES.Import_Report(objGlobal, sArchivo & sReport, sMenuId, sNomReport)
            '    objGlobal.SBOApp.StatusBar.SetText("Importado en Menú: " & sArchivo & sReport & ". Fin Sincronización.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            'End If
#End Region
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub cargaDatos()
        Dim sXML As String = ""
        Dim res As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_ODLN.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_ODLN", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_APERCIERRE.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_APERCIERRE", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            res = objGlobal.SBOApp.GetLastBatchResults
        End If
    End Sub
    Private Sub CargaFuncion()
        Dim sSQL As String = ""
        Dim sBBDD As String = ""
        Dim bResultado As Boolean = False

        If objGlobal.refDi.comunes.esAdministrador Then
            sBBDD = objGlobal.compañia.CompanyDB
            If InStr(sBBDD, "TEST") <> -1 Then
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "HANA_copbrosefectivo.sql")
            Else
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "HANA_TEST_copbrosefectivo.sql")
            End If
            bResultado = objGlobal.refDi.SQL.executeNonQuery(sSQL)
            objGlobal.SBOApp.StatusBar.SetText("HANA_copbrosefectivo.sql", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
        End If
    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "xMenuAma.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
    End Sub
    Public Overrides Function filtros() As Global.SAPbouiCOM.EventFilters
        Dim fXML As String = ""
        Try
            fXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml")
            Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
            filtro.LoadFromXML(fXML)
            Return filtro
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        Finally

        End Try
    End Function

    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Dim res As Boolean = True
        Dim Clase As Object = Nothing

        Try
            Select Case infoEvento.FormTypeEx
                Case "140"
                    Clase = New EXO_ODLN(objGlobal)
                    Return CType(Clase, EXO_ODLN).SBOApp_ItemEvent(infoEvento)
                Case "EXO_COBROT"
                    Clase = New EXO_COBROT(objGlobal)
                    Return CType(Clase, EXO_COBROT).SBOApp_ItemEvent(infoEvento)
                Case "EXOAPERTURA"
                    Clase = New EXO_Apertura(objGlobal)
                    Return CType(Clase, EXO_Apertura).SBOApp_ItemEvent(infoEvento)
                Case "EXOCIERRE"
                    Clase = New EXO_Cierre(objGlobal)
                    Return CType(Clase, EXO_Cierre).SBOApp_ItemEvent(infoEvento)
            End Select

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Finally
            Clase = Nothing
        End Try
    End Function

    Public Overrides Function SBOApp_FormDataEvent(infoEvento As BusinessObjectInfo) As Boolean
        Dim Res As Boolean = True
        Dim Clase As Object = Nothing
        Try
            Select Case infoEvento.FormTypeEx
                Case "140"
                    Clase = New EXO_ODLN(objGlobal)
                    Return CType(Clase, EXO_ODLN).SBOApp_FormDataEvent(infoEvento)
                Case "EXOAPERTURA"
                    Clase = New EXO_Apertura(objGlobal)
                    Return CType(Clase, EXO_Apertura).SBOApp_FormDataEvent(infoEvento)
                Case "EXOCIERRE"
                    Clase = New EXO_Cierre(objGlobal)
                    Return CType(Clase, EXO_Cierre).SBOApp_FormDataEvent(infoEvento)
            End Select

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Finally
            Clase = Nothing
        End Try

    End Function

    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim Clase As Object = Nothing

        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.MenuUID
                    Case "1282"
                        Select Case objGlobal.SBOApp.Forms.ActiveForm.TypeEx
                            Case "EXOAPERTURA"
                                Clase = New EXO_Apertura(objGlobal)
                                Return CType(Clase, EXO_Apertura).SBOApp_MenuEvent(infoEvento)
                        End Select
                    Case "1283" ' Eliminar
                        Select Case objGlobal.SBOApp.Forms.ActiveForm.TypeEx
                            Case "EXOAPERTURA"
                                Clase = New EXO_Apertura(objGlobal)
                                Return CType(Clase, EXO_Apertura).SBOApp_MenuEvent(infoEvento)
                            Case "EXOCIERRE"
                                objGlobal.SBOApp.StatusBar.SetText("No se puede eliminar el registro desde el cierre de caja", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                objGlobal.SBOApp.MessageBox("No se puede eliminar el registro desde el cierre de caja", 1, "Ok", "", "")
                                Exit Function
                        End Select
                End Select

                Select Case objGlobal.SBOApp.Forms.ActiveForm.TypeEx
                    Case ""
                        Return True
                End Select
            Else
                Select Case infoEvento.MenuUID
                    Case "mAperCaja"
                        Clase = New EXO_Apertura(objGlobal)
                        Return CType(Clase, EXO_Apertura).SBOApp_MenuEvent(infoEvento)
                    Case "mCierreCaja"
                        Clase = New EXO_Cierre(objGlobal)
                        Return CType(Clase, EXO_Cierre).SBOApp_MenuEvent(infoEvento)
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            Clase = Nothing
        End Try
    End Function
End Class
