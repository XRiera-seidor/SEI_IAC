'
Option Explicit On
'
Imports SEI.SEI_IAC.SEI_AddOnEnum
Imports SEI.SEI_IAC.SEI_AddOnEnum.enAddOnFormType
Imports System.Reflection
'
Public Class SEI_SBOAddon
    '
    Inherits SEI_Addon
    '
#Region "Contructor"
    Public Sub New(ByVal AddonName As String, ByRef pbo_RunApplication As Boolean)

        MyBase.New(AddonName)
        If IsNothing(Me.SBO_Application) Or IsNothing(Me.SBO_Company) Then
            ' Starting the Application
            pbo_RunApplication = False
            Exit Sub
        Else
            pbo_RunApplication = True
        End If
        initialize()
    End Sub

    Public Sub initialize()
        ' 
        'IMPORTANTE  Todos los objectos COM liberar memoria con "ReleaseComObject"
        '''''    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)
        '''''    Me.SBO_Application.MetadataAutoRefresh = True
        '
        Me.SetEventFilters()
        UpdateDatabase()
        BuildMenus()
        ' If Me.SBO_Company.UserName = "manager" Then
        AddUserDefinedObjects()
        FormattedQueries()

        sBBDD_GESTION = RecuperarValores(Me.SBO_Company, "U_SEI_BD_GESTION", "OADM", "'1'".Split, "1".Split)
        '
        ' End If
    End Sub
    '
    Private Sub FormattedQueries()
        Try
            Dim oFQ As SEI_AddingFormatedQueries
            oFQ = New SEI_AddingFormatedQueries(Me)
            oFQ.Initialize()

        Catch e As System.Exception
            Me.m_SBO_Application.MessageBox(e.Message.ToString, 1)
        End Try
    End Sub
    '
    Public Sub BuildMenus()

        Try
            Dim oMenus As SEI_AddingMenuItems
            oMenus = New SEI_AddingMenuItems(Me)

            oMenus.AddMenus()

        Catch e As System.Exception
            Me.m_SBO_Application.MessageBox(e.Message.ToString, 1)
        End Try
    End Sub

    Public Sub AddUserDefinedObjects()
        Try
            Dim oUDO As SEI_CreateUDOs

            oUDO = New SEI_CreateUDOs(Me)
            oUDO.AddUserDefinedObjects()

        Catch err As System.Exception
            Me.m_SBO_Application.MessageBox(err.Message.ToString, 1)
        End Try
    End Sub

    Public Sub UpdateDatabase()
        Try
            Dim oCreateTables As SEI_CreateTables

            oCreateTables = New SEI_CreateTables(Me)
            oCreateTables.AddUserDefinedData()

        Catch err As System.Exception
            Me.m_SBO_Application.MessageBox(err.Message.ToString, 1)
        End Try
    End Sub

#End Region

#Region "Handles Events"
    Public Overrides Sub Handle_SBO_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) ' Handles m_SBO_Application.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                SBO_Company.Disconnect()
                MyBase.InicializarIcono()
                If Not MyBase.ConnectToSBO Then
                    System.Windows.Forms.Application.Exit()
                End If
                initialize()

            Case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged
                initialize()

            Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                MsgBox("El servidor de UI se ha parado y es esencial para el funcionamiento del Add-On" & vbCrLf & _
                ".El Add-onn " & m_Name & " se ha cerrado." & vbCrLf & _
                "Por favor reinicie SAP Business One.", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Warning")
                System.Windows.Forms.Application.Exit()

            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                MyBase.InicializarIcono()
                System.Windows.Forms.Application.Exit()

        End Select

    End Sub

    Public Overrides Sub Handle_SBO_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) 'Handles m_SBO_Application.ItemEvent
        '
        Console.WriteLine("Formulari: " & FormUID & " Event: " & pVal.EventType & " Item:" & pVal.ItemUID & " BeforeAction:" & pVal.Before_Action)
        '-------------------------------------------------------
        ' Necesario para liberar Memoria del Add-on
        GC.Collect()
        GC.WaitForPendingFinalizers()
        '-------------------------------------------------------
        '
        Try
            Dim oForm As SEI_Form
            '-----------------------------------------------------
            ' Validar Formulario Modal 
            '-----------------------------------------------------
            If Not IsNothing(mst_FormUIDModal) Then
                If FormUID <> mst_FormUIDModal Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD, SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST, _
                             SAPbouiCOM.BoEventTypes.et_FORM_CLOSE, SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE
                            '----> No deshabilitar
                        Case Else
                            Dim bFormFound As Boolean = False
                            'if a son form is son
                            Dim oSBOForm As SAPbouiCOM.Form
                            For Each oSBOForm In Me.SBO_Application.Forms
                                'look in application forms collection to check if it's still open
                                If oSBOForm.UniqueID = mst_FormUIDModal Then
                                    'Son form found, select it
                                    If FormUID <> mst_FormUIDModal Then
                                        If Not Me.SBO_Application.Forms.Item(mst_FormUIDModal).Selected() Then
                                            Me.SBO_Application.Forms.Item(mst_FormUIDModal).Select()
                                        End If
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                    oSBOForm = Nothing
                                    bFormFound = True
                                End If
                            Next
                            'Form not found, so it means it has been closed                
                            If Not bFormFound Then
                                Me.mst_FormUIDModal = Nothing
                            End If

                    End Select

                End If
            End If
            ' Fin Validar Formulario Modal
            '-------------------------------------------------------------------
            '
            If pVal.BeforeAction = False Then
                '// AFTER ACTION     
                Select Case pVal.EventType

                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        ' 
                        ' Formularios Propietarios de Sap
                        ' en el New se crea un registro en la coleccion de Formularios "col_SBOFormsOpened"
                        ' 
                        '[PROGRAMAR]
                        'Dim oForm As SAPbouiCOM.Form
                        Select Case pVal.FormTypeEx

                            ''Case f_CFL_Articulos
                            ''    oForm = New SEI_CFL_Articulos(Me, pVal.FormUID)

                            Case f_OfertasVentas, f_PedidosVentas, f_EntregaVentas, f_DevolucionVentas, f_FacturaAnticipoVentas, f_FacturaVentas, f_AbonoVentas,
                              f_SolicitudCompras, f_SolicitudPedido, f_PedidosCompras, f_EntradaMercanciasCompras, f_DevolucionCompras, f_FacturaCompras, f_AbonoCompras
                                oForm = New SEI_Documentos(Me, pVal.FormUID)


                            ''Case f_ConfMail_D
                            ''    oForm = New SEI_ConfMail(Me, pVal.FormUID)
                            ''    'Case f_Encriptar_Password
                            ''    'oForm = New SEI_Encriptar_Password(Me)
                            ''Case f_ConfMail
                            ''    oForm = New SEI_ConfiguracionM(Me, pVal.FormUID)

                            Case enAddOnFormType.f_ChooseFromList_ITEMS
                                '--> Pels MODALS que criden a ChooseFromList
                                If Not Me.mst_FormUIDModal Is Nothing Then
                                    mst_FormUIDModal_PARE = mst_FormUIDModal '-> Ens guardem el Modal Pare
                                    mst_FormUIDModal = pVal.FormUID '-> El Modal passa a ser el ChooseFromList cridat
                                End If

                        End Select

                    Case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD
                        '--> Pels MODALS que criden a ChooseFromList
                        Select Case pVal.FormType
                            Case enAddOnFormType.f_ChooseFromList_ITEMS
                                If Not Me.mst_FormUIDModal Is Nothing Then
                                    mst_FormUIDModal = mst_FormUIDModal_PARE  '-> Tornem a posar com a Modal el Pare
                                    mst_FormUIDModal_PARE = "" '-> Ja no té pare
                                End If
                        End Select
                        'Dim oForm As SEI_Form
                        ' 
                        ' Eliminar el Formulario Activo de la Coleccion
                        '
                        For Each oForm In col_SBOFormsOpened
                            If oForm.UniqueID = FormUID Then
                                col_SBOFormsOpened.Remove(FormUID)
                                'TODO Falta eliminar el objeto??
                                Exit For
                            End If
                        Next

                    Case Else
                        '  Enviar Manejador de Eventos "AFTER ACTION "  al formulario Activo
                        'Dim oForm As SEI_Form

                        For Each oForm In col_SBOFormsOpened
                            If oForm.UniqueID = FormUID Then
                                oForm.HANDLE_FORM_EVENTS(FormUID, pVal, BubbleEvent)
                            End If

                        Next

                End Select
            Else
                '  Enviar Manejador de Eventos "BEFORE_ACTION"  al formulario Activo
                'Dim oForm As SEI_Form

                For Each oForm In col_SBOFormsOpened
                    If oForm.UniqueID = FormUID Then
                        oForm.HANDLE_FORM_EVENTS(FormUID, pVal, BubbleEvent)
                    End If
                Next

            End If

        Catch ExcE As Exception
            Me.SBO_Application.MessageBox(ExcE.ToString)
        End Try
        '
    End Sub

    Public Overrides Sub Handle_SBO_DataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) 'Handles m_SBO_Application.FormDataEvent
        '
        'Console.WriteLine("Formulari: " & FormUID & " Event: " & pVal.EventType & " Item:" & pVal.ItemUID & " BeforeAction:" & pVal.Before_Action)
        '-------------------------------------------------------
        ' Necesario para liberar Memoria del Add-on
        GC.Collect()
        GC.WaitForPendingFinalizers()
        '-------------------------------------------------------
        '
        Dim oForm As SEI_Form
        '
        Try
            '
            '  Enviar Manejador de Eventos   al formulario Activo
            For Each oForm In col_SBOFormsOpened
                If oForm.UniqueID = BusinessObjectInfo.FormUID Then
                    oForm.HANDLE_DATA_EVENT(BusinessObjectInfo, BubbleEvent)
                End If

            Next

        Catch ExcE As Exception
            Me.SBO_Application.StatusBar.SetText(ExcE.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        '
    End Sub

    Public Overrides Sub Handle_SBO_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) 'Handles m_SBO_Application.MenuEvent

        Console.WriteLine("MENU_EVENT Item:" & pVal.MenuUID & " BeforeAction:" & pVal.BeforeAction)

        '[PROGRAMAR] 
        Try
            '-----------------------------------------------------
            ' Validar Formulario Modal 
            '-----------------------------------------------------
            If Not IsNothing(mst_FormUIDModal) Then
                '
                Dim bFormFound As Boolean = False
                Dim oForm As SEI_Form
                '
                For Each oForm In Me.col_SBOFormsOpened
                    'look in application forms collection to check if it's still open
                    If oForm.UniqueID = mst_FormUIDModal Then
                        'Son form found, select it
                        bFormFound = True
                        If Not oForm.ModalFormEventAllowed(pVal) Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                Next
                '
                'Form not found, so it means it has been closed                
                If Not bFormFound Then
                    Me.mst_FormUIDModal = Nothing
                End If
            End If
            '
            ' Comprobar si el formulario Activo
            ' se encuentra en la coleccion de Formularios Abiertos
            ' para enviar los eventos
            ' 
            If pVal.BeforeAction = True Then
                ' BEFORE MENU ACTION
                If Me.col_SBOFormsOpened.Count <> 0 Then
                    Dim oForm As SEI_Form
                    For Each oForm In Me.col_SBOFormsOpened
                        If oForm.UniqueID = SBO_Application.Forms.ActiveForm.UniqueID Then
                            oForm.HANDLE_MENU_EVENTS(pVal, BubbleEvent)
                        End If
                    Next
                End If

            Else
                '----------------------------------------------
                '[PROGRAMAR] MENUS y FORMULARIOS DE USUARIO
                '----------------------------------------------
                '// AFTER MENU ACTION     
                Select Case pVal.MenuUID.ToUpper

                    Case enMenuUID.MNU_SEI_ErrorsGESTION.ToUpper
                        Dim oForm As SEI_Form
                        oForm = New SEI_ErrorsGESTION(Me)

                        ''Case enMenuUID.MNU_SEI_ModifyHist.ToUpper
                        ''Dim oForm As SEI_Form
                        ''oForm = New SEI_ModifyHist(Me)


                    Case Else
                        ' AFTER MENU ACTION
                        If Me.col_SBOFormsOpened.Count <> 0 Then
                            Dim oForm As SEI_Form
                            For Each oForm In Me.col_SBOFormsOpened
                                If oForm.UniqueID = SBO_Application.Forms.ActiveForm.UniqueID Then
                                    oForm.HANDLE_MENU_EVENTS(pVal, BubbleEvent)
                                End If
                            Next
                        End If
                End Select

            End If

        Catch excE As Exception
            SBO_Application.MessageBox(excE.Message.ToString)
            Me.BlockEvents = False
        End Try
    End Sub
    '
    Public Overrides Sub Handle_ReportDataEvent(ByRef eventInfo As SAPbouiCOM.ReportDataInfo, ByRef BubbleEvent As Boolean)
        '
        'Console.WriteLine("Formulari: " & FormUID & " Event: " & pVal.EventType & " Item:" & pVal.ItemUID & " BeforeAction:" & pVal.Before_Action)
        '-------------------------------------------------------
        ' Necesario para liberar Memoria del Add-on
        GC.Collect()
        GC.WaitForPendingFinalizers()
        '-------------------------------------------------------
        '
        Dim oForm As SEI_Form
        '
        Try
            '
            '  Enviar Manejador de Eventos   al formulario Activo
            For Each oForm In col_SBOFormsOpened
                If oForm.UniqueID = eventInfo.FormUID Then
                    oForm.HANDLE_REPORT_DATA_EVENT(eventInfo, BubbleEvent)
                End If

            Next

        Catch ExcE As Exception
            Me.SBO_Application.StatusBar.SetText(ExcE.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        '

    End Sub

    Public Overrides Sub Handle_SBO_PrintEvent(ByRef eventInfo As SAPbouiCOM.PrintEventInfo, ByRef BubbleEvent As Boolean)
        '
        'Console.WriteLine("Formulari: " & FormUID & " Event: " & pVal.EventType & " Item:" & pVal.ItemUID & " BeforeAction:" & pVal.Before_Action)
        '-------------------------------------------------------
        ' Necesario para liberar Memoria del Add-on
        GC.Collect()
        GC.WaitForPendingFinalizers()
        '-------------------------------------------------------
        '
        Dim oForm As SEI_Form
        '
        Try
            '
            '  Enviar Manejador de Eventos   al formulario Activo
            For Each oForm In col_SBOFormsOpened
                If oForm.UniqueID = eventInfo.FormUID Then
                    oForm.HANDLE_PRINT_EVENT(eventInfo, BubbleEvent)
                End If

            Next

        Catch ExcE As Exception
            Me.SBO_Application.StatusBar.SetText(ExcE.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        '

    End Sub

    Public Overrides Sub Handle_SBO_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        '
        'Console.WriteLine("Formulari: " & FormUID & " Event: " & pVal.EventType & " Item:" & pVal.ItemUID & " BeforeAction:" & pVal.Before_Action)
        '-------------------------------------------------------
        ' Necesario para liberar Memoria del Add-on
        GC.Collect()
        GC.WaitForPendingFinalizers()
        '-------------------------------------------------------
        '
        Dim oForm As SEI_Form
        '
        Try
            '
            '  Enviar Manejador de Eventos   al formulario Activo
            For Each oForm In col_SBOFormsOpened
                If oForm.UniqueID = eventInfo.FormUID Then
                    oForm.HANDLE_RIGHT_CLICK_EVENTS(eventInfo, BubbleEvent)
                End If

            Next

        Catch ExcE As Exception
            Me.SBO_Application.StatusBar.SetText(ExcE.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        '
    End Sub

#End Region

End Class
