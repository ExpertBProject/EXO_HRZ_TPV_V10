Imports System.IO
Imports SAPbouiCOM
Public Class EXO_GLOBALES
    Public Enum FuenteInformacion
        Visual = 1
        Otros = 2
    End Enum
#Region "Funciones formateos datos"
    Public Shared Function DblTextToNumber(ByRef oCompany As SAPbobsCOM.Company, ByVal sValor As String) As Double
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim cValor As Double = 0
        Dim sValorAux As String = "0"
        Dim sSeparadorMillarB1 As String = "."
        Dim sSeparadorDecimalB1 As String = ","
        Dim sSeparadorDecimalSO As String = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator

        DblTextToNumber = 0

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sSQL = "SELECT COALESCE(""DecSep"", ',') ""DecSep"", COALESCE(""ThousSep"", '.') ""ThousSep"" " &
                   "FROM ""OADM"" " &
                   "WHERE ""Code"" = 1"

            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                sSeparadorMillarB1 = oRs.Fields.Item("ThousSep").Value.ToString
                sSeparadorDecimalB1 = oRs.Fields.Item("DecSep").Value.ToString
            End If

            sValorAux = sValor

            If sSeparadorDecimalSO = "," Then
                If sValorAux <> "" Then
                    If Left(sValorAux, 1) = "." Then sValorAux = "0" & sValorAux

                    If sSeparadorMillarB1 = "." AndAlso sSeparadorDecimalB1 = "," Then 'Decimales ES
                        If sValorAux.IndexOf(".") > 0 AndAlso sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ""))
                        ElseIf sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ","))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    Else 'Decimales USA
                        If sValorAux.IndexOf(".") > 0 AndAlso sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "").Replace(".", ","))
                        ElseIf sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ","))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    End If
                End If
            Else
                If sValorAux <> "" Then
                    If Left(sValorAux, 1) = "," Then sValorAux = "0" & sValorAux

                    If sSeparadorMillarB1 = "." AndAlso sSeparadorDecimalB1 = "," Then 'Decimales ES
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", "").Replace(",", "."))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    Else 'Decimales USA
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", ""))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    End If
                End If
            End If

            DblTextToNumber = cValor

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Shared Function DblNumberToText(ByRef oCompany As SAPbobsCOM.Company, ByVal cValor As Double, ByVal oDestino As FuenteInformacion) As String
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim sNumberDouble As String = "0"
        Dim sSeparadorMillarB1 As String = "."
        Dim sSeparadorDecimalB1 As String = ","
        Dim sSeparadorDecimalSO As String = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator

        DblNumberToText = "0"

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sSQL = "SELECT COALESCE(""DecSep"", ',') ""DecSep"", COALESCE(""ThousSep"", '.') ""ThousSep"" " &
                   "FROM ""OADM"" " &
                   "WHERE ""Code"" = 1"

            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                sSeparadorMillarB1 = oRs.Fields.Item("ThousSep").Value.ToString
                sSeparadorDecimalB1 = oRs.Fields.Item("DecSep").Value.ToString
            End If

            If cValor.ToString <> "" Then
                If sSeparadorMillarB1 = "." AndAlso sSeparadorDecimalB1 = "," Then 'Decimales ES
                    sNumberDouble = cValor.ToString
                Else 'Decimales USA
                    sNumberDouble = cValor.ToString.Replace(",", ".")
                End If
            End If

            If oDestino = FuenteInformacion.Visual Then
                If sSeparadorDecimalSO = "," Then
                    DblNumberToText = sNumberDouble
                Else
                    DblNumberToText = sNumberDouble.Replace(".", ",")
                End If
            Else
                DblNumberToText = sNumberDouble.Replace(",", ".")
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
#End Region
    Public Shared Sub CopiarRecurso(ByVal pAssembly As Reflection.Assembly, ByVal pNombreRecurso As String, ByVal pRuta As String)
        Dim s As Stream = pAssembly.GetManifestResourceStream(pAssembly.GetName().Name + "." + pNombreRecurso)
        If s.Length = 0 Then
            Throw New Exception("No se puede encontrar el recurso '" + pNombreRecurso + "'")
        Else
            Dim buffer(CInt(s.Length() - 1)) As Byte
            s.Read(buffer, 0, buffer.Length)

            Dim sw As BinaryWriter = New BinaryWriter(File.Open(pRuta & pNombreRecurso, FileMode.Create))
            sw.Write(buffer)
            sw.Close()
        End If
    End Sub
    Public Shared Function Import_Report(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal sNOMFICH As String, ByVal sMENUID As String, ByVal sNomReport As String) As Boolean
#Region "Varibales"
        Dim oLayoutService As SAPbobsCOM.ReportLayoutsService = Nothing
        Dim oReport As SAPbobsCOM.ReportLayout = Nothing
        Dim sTypeCode As String = ""
        Dim oCompanyService As SAPbobsCOM.CompanyService = Nothing
        Dim oBlobParams As SAPbobsCOM.BlobParams = Nothing
        Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment = Nothing
        Dim oBlob As SAPbobsCOM.Blob = Nothing
        Dim sReportExiste As String = ""
        Dim sSQL As String = ""
#End Region
        Import_Report = False
        Try
            oLayoutService = CType(oObjGlobal.compañia.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService), SAPbobsCOM.ReportLayoutsService)
            oReport = CType(oLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout), SAPbobsCOM.ReportLayout)

            'Initialize critical properties 
            ' Use TypeCode "RCRI" to specify a Crystal Report. 
            ' Use other TypeCode to specify a layout for a document type. 
            ' List of TypeCode types are in table RTYP. 

            sTypeCode = "RCRI"

            oReport.Name = sNomReport
            oReport.TypeCode = sTypeCode
            oReport.Author = oObjGlobal.compañia.UserName
            oReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal
            oReport.Localization = "ES"

            Dim newReportCode As String = ""
            Try
                ' Add New object 
                oReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal

                'Comprobamos si Existe
                sSQL = "SELECT ""DocCode"" FROM  """ & oObjGlobal.compañia.CompanyDB & """.""RDOC"" WHERE ""DocName""='" & oReport.Name & "' "
                sReportExiste = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                If sReportExiste <> "" Then
                    Dim oExisteReportParams As SAPbobsCOM.ReportLayoutParams = CType(oLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutParams), SAPbobsCOM.ReportLayoutParams)
                    oExisteReportParams.LayoutCode = sReportExiste
                    oLayoutService.DeleteReportLayout(oExisteReportParams)
                    oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - Se borra Report / Layaout existente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If


                Dim oNewReportParams As SAPbobsCOM.ReportLayoutParams = CType(oLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutParams), SAPbobsCOM.ReportLayoutParams)
                Select Case sTypeCode
                    Case "RCRI" : oNewReportParams = oLayoutService.AddReportLayoutToMenu(oReport, sMENUID)
                    Case Else : oNewReportParams = oLayoutService.AddReportLayout(oReport)
                End Select

                'Get code of the added ReportLayout object 
                newReportCode = oNewReportParams.LayoutCode

            Catch ex As Exception
                Dim sError As String = Err.Description
                oObjGlobal.SBOApp.StatusBar.SetText(sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

            ' Wpload .rpt file using SetBlob interface 
            Dim rptFilePath As String = sNOMFICH

            oCompanyService = oObjGlobal.compañia.GetCompanyService()
            'Specify the table And field to update 
            oBlobParams = CType(oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams), SAPbobsCOM.BlobParams)
            oBlobParams.Table = "RDOC"
            oBlobParams.Field = "Template"

            ' Specify the record whose blob field Is to be set 
            oKeySegment = oBlobParams.BlobTableKeySegments.Add()
            oKeySegment.Name = "DocCode"
            oKeySegment.Value = newReportCode

            oBlob = CType(oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob), SAPbobsCOM.Blob)

            ' Put the rpt file into buffer 
            Dim oFile As FileStream = New FileStream(rptFilePath, System.IO.FileMode.Open)
            Dim fileSize As Integer = CType(oFile.Length, Integer)
            Dim buf(CInt(oFile.Length() - 1)) As Byte
            oFile.Read(buf, 0, fileSize)
            oFile.Close()


            ' Convert memory buffer to Base64 string 
            oBlob.Content = Convert.ToBase64String(buf, 0, fileSize)

            Try
                'Upload Blob to database 
                oCompanyService.SetBlob(oBlobParams, oBlob)
            Catch ex As Exception
                Dim sError As String = Err.Description
                oObjGlobal.SBOApp.StatusBar.SetText(sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

            Import_Report = True
        Catch ex As Exception
            Throw ex
        Finally
#Region "Liberar"
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oReport, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oLayoutService, Object))

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oBlob, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oKeySegment, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oBlobParams, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCompanyService, Object))
#End Region
        End Try
    End Function
End Class
