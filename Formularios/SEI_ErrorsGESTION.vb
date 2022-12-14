
Imports SAPbouiCOM
Imports SAPbouiCOM.BoEventTypes
Imports SEI.SEI_IAC.SEI_AddOnEnum

Public Class SEI_ErrorsGESTION
    Inherits SEI_Form

#Region "Contructor"
    Public Sub New(ByRef ParentAddon As SEI_Addon)
        'since the form is not a system form, it does not already exist 
        'therefore do not pass a uid (you don't have one anyway)
        MyBase.New(ParentAddon, enSBO_LoadFormTypes.XmlFile, enAddOnFormType.f_Errors_GESTION, "SEI_Errors_GESTION")

        'if creating controls via code, use initialize
        Initialize()
    End Sub

    Private Sub Initialize()

        Dim sSQL As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim oCol As SAPbouiCOM.ComboBox
        Dim sCode As String
        Dim sName As String

        Me.Form.SupportedModes = -1

        CarregarGrid()

    End Sub

#End Region

#Region "Funcions"

    Private Sub CarregarGrid()
        Dim sSQL As String
        Dim oDataTable0 As SAPbouiCOM.DataTable = Nothing

        Try

            Me.Form.Freeze(True)

            sSQL = " --> Articles
                    Select 'Articles' as Tipus, a.ID, a.Codi as Codi,  a.Descripcio as Descripcio, Null as Data, a.Error, a.Processat
                    From " & sBBDD_GESTION & ".dbo.ARTICLES a
                    Where a.Importat in ('E')  
                    UNION ALL
                    --> Moviments
                    Select'Moviments' as Tipus, m.ID, m.Article as Codi, Observacions as Descripcio, m.Data, m.Error, m.Processat 
                    From " & sBBDD_GESTION & ".dbo.MOVIMENTS m  
                    Where m.Importat in ('E')  
                    UNION ALL
                    --> Comandes (CAPÇALERA)
                    Select 'Comandes' as Tipus, c.ID, c.NumComanda as Codi, c.Comentaris as Descripcio, c.DataComanda, c.Error, c.Processat  
                    From " & sBBDD_GESTION & ".dbo.COMANDES_CAP c 
                    Where c.Importat in ('E')  
                    Order by Tipus, ID Desc "

            SBO_Application.StatusBar.SetText("Cargando datos...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oDataTable0 = Me.Form.DataSources.DataTables.Item("DT_0")
            oDataTable0.ExecuteQuery(sSQL)

            Configurar_Columnes_Grid()

            Me.Form.Freeze(False)

            SBO_Application.StatusBar.SetText("Datos cargados.", 2, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message)
        Finally
            LiberarObjCOM(oDataTable0)
            Me.Form.Freeze(False)
        End Try
    End Sub

    Private Sub Configurar_Columnes_Grid()
        Dim oGrid As SAPbouiCOM.Grid

        oGrid = Me.Form.Items.Item("3").Specific

        oGrid.SelectionMode = BoMatrixSelect.ms_Single

        Me.Form.Items.Item("3").Enabled = False

        oGrid.CollapseLevel = 1

        oGrid.Columns.Item("Tipus").Width = 100
        oGrid.Columns.Item("Tipus").TitleObject.Caption = "Tipus Importació"

        oGrid.Columns.Item("ID").Width = 80
        oGrid.Columns.Item("ID").TitleObject.Caption = "Identificador"

        oGrid.Columns.Item("Codi").Width = 80
        oGrid.Columns.Item("Codi").TitleObject.Caption = "Codi"

        oGrid.Columns.Item("Descripcio").Width = 150
        oGrid.Columns.Item("Descripcio").TitleObject.Caption = "Descripció"

        oGrid.Columns.Item("Data").Width = 90
        oGrid.Columns.Item("Data").TitleObject.Caption = "Data"

        oGrid.Columns.Item("Error").Width = 150
        oGrid.Columns.Item("Error").TitleObject.Caption = "Error"

        oGrid.Columns.Item("Processat").Width = 110
        oGrid.Columns.Item("Processat").TitleObject.Caption = "Processat"

        '''oGrid.Rows.CollapseAll()

        LiberarObjCOM(oGrid)

    End Sub


    Private Sub Reprocessar()
        Dim oGrid As SAPbouiCOM.Grid
        Dim oQuery As SAPbobsCOM.Recordset
        Dim sQuery As String

        Dim iFilaSel As Integer

        Dim bEsTipus As Boolean

        Dim sId As String
        Dim sTipus As String
        Dim sTaula As String

        Try

            oGrid = Me.m_SBO_Form.Items.Item("3").Specific

            If oGrid.Rows.SelectedRows.Count = 0 Then
                SBO_Application.MessageBox("Has de seleccionar una fila")
                Exit Sub
            End If

            iFilaSel = oGrid.Rows.SelectedRows.Item(0, BoOrderType.ot_SelectionOrder)

            If oGrid.Rows.IsLeaf(iFilaSel) Then
                bEsTipus = False

                sId = oGrid.DataTable.GetValue("ID", oGrid.GetDataTableRowIndex(iFilaSel))
                sTipus = oGrid.DataTable.GetValue("Tipus", oGrid.GetDataTableRowIndex(iFilaSel))

            Else
                bEsTipus = True

                sTipus = oGrid.DataTable.GetValue("Tipus", oGrid.GetDataTableRowIndex(iFilaSel + 1))
            End If

            If sTipus = "Comandes" Then
                SBO_Application.SetStatusBarMessage("Les Comandes s'han de Reprocessar des del GESTION.", BoMessageTime.bmt_Short, True)
                Exit Sub
            End If

            If bEsTipus Then
                '-> Tots de cop
                If SBO_Application.MessageBox("Està segur que vol reprocessar TOTS els Identificadors de '" & sTipus & "' que estan com a ERROR?", 2, "Sí", "No") = 2 Then
                    SBO_Application.SetStatusBarMessage("Operació cancel·lada.", BoMessageTime.bmt_Short, True)
                    Exit Sub
                End If
            Else
                '-> Només un
                If SBO_Application.MessageBox("Està segur que vol reprocessar el Identificador '" & sId & "' de '" & sTipus & "'?", 2, "Sí", "No") = 2 Then
                    SBO_Application.SetStatusBarMessage("Operació cancel·lada.", BoMessageTime.bmt_Short, True)
                    Exit Sub
                End If
            End If

            Select Case sTipus
                Case "Articles"
                    sTaula = "ARTICLES"

                Case "Moviments"
                    sTaula = "MOVIMENTS"

                Case "Comandes"  '-> no arribarà mai
                    sTaula = "COMANDES_CAP"

            End Select

            '--> Reprocesar 
            sQuery = " "
            sQuery = sQuery & " Update " & sBBDD_GESTION & ".[dbo].[" & sTaula & "] "
            sQuery = sQuery & " Set Importat = 'N' "
            sQuery = sQuery & " Where Importat = 'E' "
            If Not bEsTipus Then
                '-> Només un
                sQuery = sQuery & "     And ID ='" & sId & "' "
            End If

            oQuery = Me.SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oQuery.DoQuery(sQuery)

            SBO_Application.SetStatusBarMessage("Operació realitzada.", BoMessageTime.bmt_Short, False)

            CarregarGrid()


        Catch ex As Exception
            SBO_Application.MessageBox("Hi ha hagut algun problema al intentar Reprocessar. Causa: " & ex.ToString)
        Finally
            LiberarObjCOM(oGrid)
            LiberarObjCOM(oQuery)
        End Try


    End Sub

    Private Sub Eliminar()
        Dim oGrid As SAPbouiCOM.Grid
        Dim oQuery As SAPbobsCOM.Recordset
        Dim sQuery As String

        Dim iFilaSel As Integer

        Dim sID As String
        Dim sTipus As String
        Dim sTaula As String

        Try

            oGrid = Me.m_SBO_Form.Items.Item("3").Specific

            If oGrid.Rows.SelectedRows.Count = 0 Then
                SBO_Application.MessageBox("Has de seleccionar una fila")
                Exit Sub
            End If

            iFilaSel = oGrid.Rows.SelectedRows.Item(0, BoOrderType.ot_SelectionOrder)

            If Not oGrid.Rows.IsLeaf(iFilaSel) Then
                sTipus = oGrid.DataTable.GetValue("Tipus", oGrid.GetDataTableRowIndex(iFilaSel + 1))
                SBO_Application.SetStatusBarMessage("No es poden eliminar de cop de la llista d'errors els Identificadors de '" & sTipus & "'.", BoMessageTime.bmt_Short, True)
                Exit Sub
            End If

            sID = oGrid.DataTable.GetValue("ID", oGrid.GetDataTableRowIndex(iFilaSel))
            sTipus = oGrid.DataTable.GetValue("Tipus", oGrid.GetDataTableRowIndex(iFilaSel))

            If SBO_Application.MessageBox("Està segur que vol Eliminar de la llista d'errors el Identificador '" & sID & "' de '" & sTipus & "'?", 2, "Sí", "No") = 2 Then
                SBO_Application.SetStatusBarMessage("Operació cancel·lada.", BoMessageTime.bmt_Short, True)
                Exit Sub
            End If

            Select Case sTipus
                Case "Articles"
                    sTaula = "ARTICLES"

                Case "Moviments"
                    sTaula = "MOVIMENTS"

                Case "Comandes"  '-> no arribarà mai
                    sTaula = "COMANDES_CAP"

            End Select

            '--> "Eliminar" de la pantalla d'errors
            sQuery = " "
            sQuery = sQuery & " Update " & sBBDD_GESTION & ".[dbo].[" & sTaula & "] "
            sQuery = sQuery & " Set Importat = 'G' " '-> Error Gestionat
            sQuery = sQuery & " Where Importat = 'E' "
            sQuery = sQuery & "     And ID ='" & sID & "' "


            oQuery = Me.SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oQuery.DoQuery(sQuery)

            SBO_Application.SetStatusBarMessage("Operació realitzada.", BoMessageTime.bmt_Short, False)

            CarregarGrid()

        Catch ex As Exception
            SBO_Application.MessageBox("Hi ha hagut algun problema al intentar Eliminar dels errors. Causa: " & ex.ToString)
        Finally
            LiberarObjCOM(oGrid)
            LiberarObjCOM(oQuery)
        End Try
    End Sub

#End Region

    Public Overrides Sub HANDLE_PRINT_EVENT(ByRef eventInfo As PrintEventInfo, ByRef BubbleEvent As Boolean)

    End Sub

    Public Overrides Sub HANDLE_REPORT_DATA_EVENT(ByRef eventInfo As ReportDataInfo, ByRef BubbleEvent As Boolean)

    End Sub

    Public Overrides Sub HANDLE_FORM_EVENTS(FormUID As String, ByRef pVal As ItemEvent, ByRef BubbleEvent As Boolean)
        If Trim(pVal.ItemUID) <> "" Then

            '---- Botons ----
            If pVal.EventType = et_ITEM_PRESSED And Not pVal.BeforeAction Then
                Select Case pVal.ItemUID
                    Case "6" '-> Mostrar
                        CarregarGrid()

                    Case "4" '-> Eliminar
                        Eliminar()

                    Case "5" '-> Reprocessar
                        Reprocessar

                End Select
            End If
        Else

            Select Case pVal.EventType
                Case et_FORM_ACTIVATE

                Case et_FORM_DEACTIVATE

                Case et_FORM_LOAD

                Case et_FORM_UNLOAD

                Case et_FORM_CLOSE

            End Select

        End If
    End Sub

    Public Overrides Sub HANDLE_MENU_EVENTS(ByRef pVal As MenuEvent, ByRef BubbleEvent As Boolean)

    End Sub

    Public Overrides Sub HANDLE_RIGHT_CLICK_EVENTS(ByRef eventInfo As ContextMenuInfo, ByRef BubbleEvent As Boolean)

    End Sub

    Public Overrides Function HANDLE_DATA_EVENT(ByRef BusinessObjectInfo As BusinessObjectInfo, ByRef BubbleEvent As Boolean) As Integer

    End Function

    Public Overrides Function ModalFormEventAllowed(ByRef pVal As MenuEvent) As Boolean

    End Function
End Class
