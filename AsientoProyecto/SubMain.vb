Module SubMain

    Public SBOCompany As SAPbobsCOM.Company

    Sub Main()

        Conectar()
        UpdateOJDT()

    End Sub


    Public Function Conectar()

        Try

            SBOCompany = New SAPbobsCOM.Company

            SBOCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            SBOCompany.Server = My.Settings.Server
            SBOCompany.LicenseServer = My.Settings.LicenseServer
            SBOCompany.DbUserName = My.Settings.DbUserName
            SBOCompany.DbPassword = My.Settings.DbPassword

            SBOCompany.CompanyDB = My.Settings.CompanyDB

            SBOCompany.UserName = My.Settings.UserName
            SBOCompany.Password = My.Settings.Password

            SBOCompany.Connect()

        Catch ex As Exception

            MsgBox("Error al Conectar: " & ex.Message)

        End Try

    End Function


    Public Function UpdateOJDT()

        Dim oOJDT As SAPbobsCOM.JournalEntries

        Dim oRecSettxb, oRecSettxb2 As SAPbobsCOM.Recordset
        Dim stQuerytxb, stQuerytxb2 As String
        Dim TransId, Project, LineID, LineProject, LineProfit As String

        Try

            oRecSettxb = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb = "call ""Asiento_Titulo_Proyect"""

            oRecSettxb.DoQuery(stQuerytxb)

            If oRecSettxb.RecordCount > 0 Then

                oRecSettxb.MoveFirst()

                For cont As Integer = 0 To oRecSettxb.RecordCount - 1

                    TransId = oRecSettxb.Fields.Item("TransId").Value
                    Project = oRecSettxb.Fields.Item("Project").Value

                    oOJDT = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                    '///// Encabezado del Asiento
                    oOJDT.GetByKey(TransId)
                    oOJDT.ProjectCode = Project

                    oOJDT.Update()

                    oRecSettxb2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    stQuerytxb2 = "call ""Asiento_Linea_Proyect""('" & TransId & "')"

                    oRecSettxb2.DoQuery(stQuerytxb2)

                    If oRecSettxb2.RecordCount > 0 Then

                        oRecSettxb2.MoveFirst()

                        For cont2 As Integer = 0 To oRecSettxb2.RecordCount - 1

                            LineID = oRecSettxb2.Fields.Item("Line_ID").Value
                            LineProject = oRecSettxb2.Fields.Item("Project").Value
                            LineProfit = oRecSettxb2.Fields.Item("ProfitCode").Value

                            oOJDT.Lines.SetCurrentLine(LineID)

                            If LineProject = Nothing Then
                                oOJDT.Lines.ProjectCode = Project
                            End If

                            If LineProfit = Nothing Then
                                oOJDT.Lines.CostingCode = Project
                            End If

                            oRecSettxb2.MoveNext()

                        Next

                        oOJDT.Update()

                    End If

                    oRecSettxb.MoveNext()

                Next

            End If

        Catch ex As Exception

            MsgBox("Error al Actualizar: " & ex.Message)

        End Try

    End Function

End Module
