Option Explicit On
'
Imports System.Windows.forms
Imports System.Threading
Imports System.Timers
Imports System.Drawing
Imports System.Xml

Imports System.Diagnostics
'
Public MustInherit Class SEI_Addon
    Inherits Object
    Implements SEI_IAddOn

#Region "variables Icono Sap"
    Public bMostrarIcono As Boolean
    Private NotifyIcon1 As NotifyIcon
    Private oTimer As System.Timers.Timer
    Private bSemafor As Boolean
    Private oIconRed As Icon
    Private oIconBlue As Icon

#End Region

#Region "variables"
    '
    Protected WithEvents m_SBO_Application As SAPbouiCOM.Application
    Protected m_SBO_Company As SAPbobsCOM.Company
    '
    Const DEVELOPERSCONNECTIONSTRING As String = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
    Private Const EXTENSION = ".srf"
    '
    Protected m_StartUpPath As String
    Protected m_CompanyName As String
    Protected m_Name As String
    Protected m_AddOnHomePath As String
    Public col_SBOFormsOpened As Collection
    Public mst_FormUIDModal As String
    Public mst_FormUIDModal_PARE As String
    Protected m_Connected As Boolean
    '
    Protected m_EventsBlocked As Boolean
    Protected m_ErrCode As Long
    Protected m_ErrMsg As String
    Protected m_RutaFitxer As String
    Protected m_FormParentID As String  'ID Formulario Padre
    Protected m_NomFitxer As String
    Protected m_TipoFichero As String
    Protected m_TituloShowDialogo As String
    Protected m_DirectorioInicial As String
    Protected m_Mensaje As String
    '
#End Region

#Region "constructor"
    Public Sub New(ByVal AddOnName As String)
        m_Name = AddOnName
        bMostrarIcono = True
        Initialize()
    End Sub
    Private Sub Initialize()
        col_SBOFormsOpened = New Collection
        m_EventsBlocked = False
        If ConnectToSBO() = False Then
            Application.Exit()
        End If
    End Sub

#End Region

#Region "MustOverrides"
    Public MustOverride Sub Handle_SBO_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles m_SBO_Application.ItemEvent
    Public MustOverride Sub Handle_SBO_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles m_SBO_Application.AppEvent
    Public MustOverride Sub Handle_SBO_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles m_SBO_Application.MenuEvent
    Public MustOverride Sub Handle_SBO_DataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles m_SBO_Application.FormDataEvent
    Public MustOverride Sub Handle_SBO_PrintEvent(ByRef eventInfo As SAPbouiCOM.PrintEventInfo, ByRef BubbleEvent As Boolean) Handles m_SBO_Application.PrintEvent
    Public MustOverride Sub Handle_SBO_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles m_SBO_Application.RightClickEvent
    Public MustOverride Sub Handle_ReportDataEvent(ByRef eventInfo As SAPbouiCOM.ReportDataInfo, ByRef BubbleEvent As Boolean) Handles m_SBO_Application.ReportDataEvent

#End Region

#Region "properties"
    Public ReadOnly Property Connected() As Boolean Implements SEI_IAddOn.Connected
        Get
            Return m_Connected
        End Get
    End Property

    Public ReadOnly Property Name() As String Implements SEI_IAddOn.Name
        Get
            Return m_Name
        End Get
    End Property

    Public ReadOnly Property HomePath() As String Implements SEI_IAddOn.HomePath
        Get
            Return m_AddOnHomePath
        End Get
    End Property

    Public ReadOnly Property SBO_Application() As SAPbouiCOM.Application Implements SEI_IAddOn.SBO_Application
        Get
            Return m_SBO_Application
        End Get
    End Property

    Public ReadOnly Property SBO_Company() As SAPbobsCOM.Company Implements SEI_IAddOn.SBO_Company
        Get
            Return m_SBO_Company
        End Get
    End Property

    Public ReadOnly Property SBO_Forms() As Collection Implements SEI_IAddOn.SBO_Forms
        Get
            Return col_SBOFormsOpened
        End Get
    End Property

    Public ReadOnly Property StartupPath() As String Implements SEI_IAddOn.StartupPath
        Get
            Return m_StartUpPath
        End Get
    End Property

    Public Property BlockEvents() As Boolean Implements SEI_IAddOn.BlockEvents
        Get
            Return m_EventsBlocked
        End Get
        Set(ByVal Value As Boolean)
            m_EventsBlocked = Value
        End Set
    End Property

    Property FormParentID() As String
        Get
            ' Formlario Padre 
            FormParentID = m_FormParentID
        End Get
        Set(ByVal value As String)
            m_FormParentID = value
        End Set
    End Property
    Property ErrCode() As Long
        ' Codigo de error de GetLastError de la Company 
        Get
            ErrCode = m_ErrCode
        End Get
        Set(ByVal value As Long)
            m_ErrCode = value
        End Set
    End Property

    Property ErrMsg() As String
        ' Descripcion de error de GetLastError de la Company 
        Get
            ErrMsg = m_ErrMsg
        End Get
        Set(ByVal value As String)
            m_ErrMsg = value
        End Set
    End Property

    Public Property RutaFitxer() As String
        Get
            RutaFitxer = m_RutaFitxer
        End Get
        Set(ByVal value As String)
            m_RutaFitxer = value
        End Set
    End Property

    Public Property NomFitxer() As String
        Get
            NomFitxer = m_NomFitxer
        End Get
        Set(ByVal value As String)
            m_NomFitxer = value
        End Set
    End Property

    Public Property TipoFichero() As String
        Get
            TipoFichero = m_TipoFichero
        End Get
        Set(ByVal value As String)
            m_TipoFichero = value
        End Set
    End Property
    '
    Public Property TituloShowDialogo() As String
        Get
            TituloShowDialogo = m_TituloShowDialogo
        End Get
        Set(ByVal value As String)
            m_TituloShowDialogo = value
        End Set
    End Property

    Public Property DirectorioInicial() As String
        Get
            DirectorioInicial = m_DirectorioInicial
        End Get
        Set(ByVal value As String)
            m_DirectorioInicial = value
        End Set
    End Property
    '
    Public Property Mensaje() As String
        Get
            Mensaje = m_Mensaje
        End Get
        Set(ByVal value As String)
            m_Mensaje = value
        End Set
    End Property


    Public ReadOnly Property BitMapPath() As String
        Get
            BitMapPath = SBO_Company.BitMapPath & "ADDON\"
        End Get
    End Property

#End Region

#Region "functions"
    '************************************************************************************
    Private Sub m_SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles m_SBO_Application.AppEvent

        If m_EventsBlocked = False Then

            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                    'If Me.SBO_Company.Connected Then
                    '    Me.SBO_Company.Disconnect()
                    'End If
                    'bMostrarIcono = False
                    'Initialize()

                Case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged

                Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                    MsgBox("The UIServer, which is essential for addon functionality" & vbCrLf & _
                    " has been stopped. The addon " & m_Name & " is closing down." & vbCrLf & _
                    "Please restart SAP Business One.", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "warning")
                    Application.Exit()

                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                    Application.Exit()
            End Select
        End If
    End Sub

    Private Function GetNewFormUID() As String
        Return Me.Name.ToString & SEI_Form.Count
    End Function
    '
    Public Function CreateSBO_Form(ByVal pst_XMLDocumentName As String, ByVal pst_FormType As String, Optional ByVal pst_UDOCode As String = "") As SAPbouiCOM.Form Implements SEI_IAddOn.CreateSBO_Form
        'create a new form via xml
        '-----------------------------------------------------------------------------
        ' DESCRIPTION  : Create a form
        ' Entry       : XMLDocumentName (string)   : name of the XML File (without extension)
        '                FormType (string)          : Form Type (MRO_ENUM.enAeroOneFormType.cst_Counter As String = "MRO_0001"
        '                UDOName (string)           : Name of the UDO if there is one attached   
        '                   = ""  : Create a normal form
        '                   <> "" : Create a UDO type form
        ' Exit       : object SAPbouiCOM.Form
        ' MODIFICATION : 
        '-----------------------------------------------------------------------------
        Dim dsa_CreationPackage As SAPbouiCOM.FormCreationParams
        Dim dxl_XLMDocument As System.Xml.XmlDocument
        Try
            dsa_CreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)

            dsa_CreationPackage.UniqueID = GetNewFormUID()
            dsa_CreationPackage.FormType = pst_FormType
            If pst_UDOCode <> "" Then
                dsa_CreationPackage.ObjectType = pst_UDOCode
            End If
            dxl_XLMDocument = GetFormDefinitions(pst_XMLDocumentName)

            'set form position (X and Y)
            dsa_CreationPackage.XmlData = dxl_XLMDocument.InnerXml
            Return SBO_Application.Forms.AddEx(dsa_CreationPackage)

        Catch ex As Exception
            Throw ex
        End Try

    End Function
    '
    Private Function GetFormDefinitions(ByVal NomFormulari As String) As XmlDocument
        Dim oXMLDocument As XmlDocument = New XmlDocument
        Dim oFilename As String
        '
        oFilename = Application.StartupPath() & "\Formularios_srf\" & NomFormulari & ".srf"
        oXMLDocument.Load(oFilename)

        'SetFormPosition(oXMLDocument)
        Return oXMLDocument
    End Function

#End Region

#Region "connect to SBO"
    Public Function ConnectToSBO() As Boolean
        m_Connected = False

        'If m_AppConfig.AppMode = "DEBUG" Then
        '    MsgBox("Method ConnectToSBO called")
        'End If

        'Try to connect to SAP Business One UIAPI
        If ConnectToUIAPI() = True Then
            'Try to connect to SAP Business One DIAPI
            If ConnectToDIAPI() = True Then

                'Both connections were successful
                m_Connected = True
                Dim strMessage As String

                strMessage = "SBO_AddOn connected."
                Me.m_SBO_Application.StatusBar.SetText(strMessage, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                '
                If bMostrarIcono Then
                    MostrarIconoSap()
                End If

                'If m_AppConfig.AppMode = "DEBUG" Then
                '    m_SBO_Application.MessageBox(strMessage)
                'Else
                '    Me.m_SBO_Application.StatusBar.SetText(strMessage, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                'End If
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If

        'If m_AppConfig.AppMode = "DEBUG" Then
        '    MsgBox("Method ConnectToSBO ended")
        'End If
    End Function
    Public Sub InicializarIcono()
        NotifyIcon1.Dispose()
        oTimer.Stop()
        oTimer.Close()
        oTimer.Dispose()
    End Sub

    Private Function ConnectToUIAPI() As Boolean
        Try
            '*******************************************************************
            '// Use an SboGuiApi object to establish connection
            '// with the SAP Business One application and return an
            '// initialized appliction object
            '*******************************************************************

            Dim SboGuiApi As SAPbouiCOM.SboGuiApi
            Dim sConnectionString As String

            'If m_AppConfig.AppMode = "DEBUG" Then
            '    MsgBox("Method ConnectToUIAPI called")
            'End If

            SboGuiApi = New SAPbouiCOM.SboGuiApi

            '// by following the steps specified above, the following
            '// statment should be suficient for either development or run mode
#If DEBUG Then
            sConnectionString = DEVELOPERSCONNECTIONSTRING
#Else
            sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            'sConnectionString = DEVELOPERSCONNECTIONSTRING
#End If

            'If m_AppConfig.AppMode = "DEBUG" Then
            '    MsgBox("Connectionstring: " & sConnectionString)
            'End If

            '// connect to a running SBO Application
            SboGuiApi.Connect(sConnectionString)

            '// get an initialized application object

            m_SBO_Application = SboGuiApi.GetApplication()

            'If m_AppConfig.AppMode = "DEBUG" Then
            '    MsgBox("Method ConnectToUIAPI ended")
            'End If

            '' m_SBO_Application.StatusBar.SetText("Add-on Conectado", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Return True

        Catch excE As Exception

            'If m_AppConfig.AppMode = "DEBUG" Then
            '    MsgBox("Method ConnectToUIAPI failed to connect")
            '    MsgBox("Method ConnectToUIAPI ended")
            '    MsgBox(excE.ToString)
            'End If

            MsgBox("El Addon " & m_Name & " no puede conectar con SAP Business One." _
                   & vbCrLf & "Tiene que reiniciar SAP Business One o" & vbCrLf & _
                   "ponerse en contacto con su Administrador.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Addon " & m_Name)
            Application.Exit()
        End Try
    End Function

    Private Function ConnectToDIAPI() As Boolean
        'catch error
        Dim ErrCode As Integer
        Dim ErrMessage As String

        Dim strCookie As String
        Dim strConnectionContext As String

        ErrCode = 0
        ErrMessage = ""

        Try
            'If m_AppConfig.AppMode = "DEBUG" Then
            '    MsgBox("Method ConnectToDIAPI called")
            'End If

            Me.m_EventsBlocked = True

            'Create company object
            m_SBO_Company = New SAPbobsCOM.Company

            'get connection context
            strCookie = m_SBO_Company.GetContextCookie

            'retrieve connection context string via cookie
            strConnectionContext = m_SBO_Application.Company.GetConnectionContext(strCookie)

            m_SBO_Company.SetSboLoginContext(strConnectionContext)

            If m_SBO_Company.Connect() <> 0 Then

                m_SBO_Company.GetLastError(ErrCode, ErrMessage)

                m_SBO_Application.MessageBox("Error al conectar: " & ErrMessage & _
                vbCrLf & ErrMessage)
                'If m_AppConfig.AppMode = "DEBUG" Then
                '    MsgBox("Could not connect to DIAPI")
                'End If


                Me.m_EventsBlocked = False

                Me.m_Connected = False

                Return False
            Else
                m_SBO_Application.StatusBar.SetText("Add-on conectado", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                'If m_AppConfig.AppMode = "DEBUG" Then
                '    MsgBox("Connected to DIAPI")
                'End If

                Me.m_EventsBlocked = False
                Return True
            End If

        Catch excE As Exception
            MsgBox(excE.Message.ToString)
            Application.Exit()
        End Try
    End Function

    Private Sub MostrarIconoSap()
        '----------------------------------------------------------------
        ' Definir Iconos
        'oIconRed = New Icon(Me.GetType, My.Resources.ResourceManager.GetObject("sap48.ico"))
        oIconRed = New Icon(Me.GetType(), "sap48Rojo.ico")
        oIconBlue = New Icon(Me.GetType(), "sap48.ico")
        '
        NotifyIcon1 = New NotifyIcon
        NotifyIcon1.Icon = oIconRed
        NotifyIcon1.Visible = True
        NotifyIcon1.Text = Me.Name & " (Versión " & Application.ProductVersion & ")"
        '
        '----------------------------------------------------------------
        ' Definir Temporizador 
        oTimer = New System.Timers.Timer
        '
        AddHandler oTimer.Elapsed, AddressOf OnTimer
        ' 
        oTimer.Interval = 1000
        oTimer.Enabled = True
        '----------------------------------------------------------------

    End Sub
    '
    Public Sub OnTimer(ByVal source As Object, ByVal e As ElapsedEventArgs)

        If bSemafor Then
            NotifyIcon1.Icon = oIconBlue
            bSemafor = False
        Else
            NotifyIcon1.Icon = oIconRed
            bSemafor = True
        End If

        '''If Me.SBO_Application Is Nothing Then
        '''    Try
        '''        Dim s


        '''        s = s
        '''        ConnectToUIAPI()
        '''    Catch ex As Exception
        '''        SBO_Application.MessageBox(ex.Message)
        '''    End Try

        '''    ''    End
        '''End If

    End Sub

    Private Sub OpenFile()
        Try
            Dim MyDialog As New OpenFileDialog()

            Dim oWindowsSbo As SEI_WindowsSbo
            oWindowsSbo = New SEI_WindowsSbo(Me)
            MyDialog.ShowHelp = True

            MyDialog.ShowDialog(oWindowsSbo)
            Me.RutaFitxer = MyDialog.FileName

        Catch ExcE As Exception
            Me.SBO_Application.MessageBox(ExcE.Message.ToString)
        End Try

    End Sub

    Public Function ObrirFitxer__OLD() As String

        Dim myThread As New System.Threading.Thread(AddressOf OpenFile)
        myThread.SetApartmentState(Threading.ApartmentState.STA)
        myThread.Start()
        myThread.Join()
        ObrirFitxer__OLD = Me.RutaFitxer
        myThread.Abort()

    End Function

    Public Function ObrirFitxer(Optional ByRef sNombreFichero As String = "", _
                                Optional ByVal sTipoFichero As String = "", _
                                Optional ByVal sDirectorioInicial As String = "", _
                                Optional ByVal sTituloShowDialogo As String = "") As String

        Me.TipoFichero = sTipoFichero
        Me.TituloShowDialogo = sTituloShowDialogo

        If sDirectorioInicial = "" Then
            Me.DirectorioInicial = Me.SBO_Company.AttachMentPath
        Else
            Me.DirectorioInicial = sDirectorioInicial
        End If

        Dim myThread As New System.Threading.Thread(AddressOf OpenFile)
        myThread.SetApartmentState(Threading.ApartmentState.STA)
        myThread.Start()
        myThread.Join()
        ObrirFitxer = Me.RutaFitxer
        sNombreFichero = Me.NomFitxer
        myThread.Abort()

    End Function

    Private Sub SaveFile()

        Try
            Dim MyDialog As New SaveFileDialog()

            Dim oWindowsSbo As SEI_WindowsSbo
            oWindowsSbo = New SEI_WindowsSbo(Me)
            MyDialog.ShowHelp = True
            MyDialog.InitialDirectory = Me.DirectorioInicial
            If Me.TipoFichero <> "" Then
                MyDialog.Filter = Me.TipoFichero
                MyDialog.FilterIndex = 0
            End If
            If Me.TituloShowDialogo <> "" Then
                MyDialog.Title = Me.TituloShowDialogo
            End If
            If Me.NomFitxer <> "" Then
                MyDialog.FileName = Me.NomFitxer
            End If

            MyDialog.OverwritePrompt = False
            MyDialog.ShowDialog(oWindowsSbo)
            Me.RutaFitxer = MyDialog.FileName
            Me.NomFitxer = ObtenerNombre(MyDialog.FileName)

        Catch ExcE As Exception
            Me.SBO_Application.MessageBox(ExcE.ToString)
        End Try

    End Sub

    Public Function GuardarFitxer(Optional ByRef sNombreFichero As String = "", _
                                Optional ByVal sTipoFichero As String = "", _
                                Optional ByVal sDirectorioInicial As String = "", _
                                Optional ByVal sTituloShowDialogo As String = "") As String

        Me.TipoFichero = sTipoFichero
        Me.TituloShowDialogo = sTituloShowDialogo
        Me.NomFitxer = sNombreFichero

        If sDirectorioInicial = "" Then
            Me.DirectorioInicial = Me.SBO_Company.AttachMentPath
        Else
            Me.DirectorioInicial = sDirectorioInicial
        End If

        Dim myThread As New System.Threading.Thread(AddressOf SaveFile)
        myThread.SetApartmentState(Threading.ApartmentState.STA)
        myThread.Start()
        myThread.Join()
        GuardarFitxer = Me.RutaFitxer
        sNombreFichero = Me.NomFitxer
        myThread.Abort()

    End Function

    Private Function ObtenerNombre(ByVal sRuta As String) As String

        Dim sValor() As String

        sValor = sRuta.Split("\")

        ObtenerNombre = sValor(sValor.Length - 1)

    End Function


#End Region

#Region "SetFilters"
    Public Function SetEventFilters() As Long
        '''''
        ''''Dim oFilters As New SAPbouiCOM.EventFilters
        ''''Dim oFilter As SAPbouiCOM.EventFilter

        '''''Perquè capturi els events de botons s'ha d'afegir l'event sol sense cap formulari
        ''''oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
        '''''
        '''''[PROGRAMAR]
        ''''oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ALL_EVENTS)
        '''''-> Ventes
        ''''oFilter.AddEx(SEI_AddOnEnum.enAddOnFormType.f_OfertasVentas)
        ''''oFilter.AddEx(SEI_AddOnEnum.enAddOnFormType.f_PedidosVentas)
        ''''oFilter.AddEx(SEI_AddOnEnum.enAddOnFormType.f_EntregaVentas)
        ''''oFilter.AddEx(SEI_AddOnEnum.enAddOnFormType.f_DevolucionVentas)
        ''''oFilter.AddEx(SEI_AddOnEnum.enAddOnFormType.f_FacturaVentas)
        ''''oFilter.AddEx(SEI_AddOnEnum.enAddOnFormType.f_FacturaAnticipoVentas)
        ''''oFilter.AddEx(SEI_AddOnEnum.enAddOnFormType.f_AbonoVentas)
        '''''-> Compres        
        ''''oFilter.AddEx(SEI_AddOnEnum.enAddOnFormType.f_PedidosCompras)
        ''''oFilter.AddEx(SEI_AddOnEnum.enAddOnFormType.f_EntradaMercanciasCompras)
        ''''oFilter.AddEx(SEI_AddOnEnum.enAddOnFormType.f_DevolucionCompras)
        ''''oFilter.AddEx(SEI_AddOnEnum.enAddOnFormType.f_FacturaCompras)
        ''''oFilter.AddEx(SEI_AddOnEnum.enAddOnFormType.f_AbonoCompras)
        '''''-> Altres
        ''''oFilter.AddEx(SEI_AddOnEnum.enAddOnFormType.f_Articulos)
        ''''oFilter.AddEx(SEI_AddOnEnum.enAddOnFormType.f_InterlocutoresComerciales)
        ''''oFilter.AddEx(SEI_AddOnEnum.enAddOnFormType.f_Alertas)

        '''''
        ''''Me.SBO_Application.SetFilter(oFilters)
        '''''
    End Function

#End Region



    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
