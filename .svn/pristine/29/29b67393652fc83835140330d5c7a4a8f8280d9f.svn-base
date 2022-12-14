Option Strict Off
Option Explicit On
'
Imports SEI.SEI_IAC.SEI_AddOnEnum
Imports SAPbobsCOM.BoObjectTypes
Imports SAPbobsCOM.BoFormattedSearchActionEnum
Imports SAPbouiCOM.BoFormMode
Imports SAPbouiCOM.BoEventTypes
Imports SAPbouiCOM.BoDataType
Imports System.Collections
Imports SAPbobsCOM.BoYesNoEnum
'
Public Class SEI_AddingFormatedQueries
    '

#Region "Variables"

    Private _ParentAddon As SEI_Addon
    Private _Form As SAPbouiCOM.Form
    Private sErrNumber As String
    Private iErrMsg As Integer

    'Aquest mòdul afegeix les búsquedes formatejades necessàries per a l'Add-On
    '
    'Taules implicades
    'OQCN -> Es crearà una categoria amb el nom SDK on s'hi afegiran les consultes formatejades
    'OUQR -> Consultes necessàries per al Gestor de consultes
    'CSHS -> Taula on es relaciona la consulta formatejada amb el item
    'CUVV -> Si la consulta formatejada es basa amb valors existents, aquí és on es guarden aquests valors
    '
    Const c_Categoria As String = "SDK"
    Private Enum ACCIO
        AccioValors = 1
        AccioConsulta = 2
    End Enum

#End Region

#Region "Contructor"
    Private _sEI_SBOAddon As SEI_SBOAddon

    '
    Sub New(ByRef o_ParentAddon As SEI_Addon) ', ByRef o_Form As SAPbouiCOM.Form)
        _ParentAddon = o_ParentAddon
        '_Form = o_Form
    End Sub

#End Region

#Region "Funciones"

    Public Sub Initialize()
        'EliminarConsultesFormatejades()
        AddFormatedQueries()
    End Sub


    'Sub New(sei_sboaddon As SEI_SBOAddon)
    '    ' todo: complete member initialization 
    '    _sei_sboaddon = sei_sboaddon
    'End Sub

    '
    Public Sub AddFormatedQueries()
        '
        'Afegeixo la categoria "SDK"
        AfegirCategoria_SDK()
        'Afegeixo les consultes necessàries dintre la categoria "SDK"
        AfegirGestorConsultes()
        'Afegeixo les consultes formatejades necessàries
        AfegirBusquedesFormatejades()
        '
    End Sub
    '
    Private Sub AfegirGestorConsultes()
        '
        Dim lCategoria As Long
        Dim sNomConsulta As String
        Dim sSQL As String
        Dim sHANA As String
        '''
        ''lCategoria = ContadorCategoriaSDK()
        '''
        ''If lCategoria = -1 Then
        ''    Me._ParentAddon.SBO_Application.MessageBox(("No existe la Categoria SDK"))
        ''    Exit Sub
        ''End If
        '''
        '''
        ''sNomConsulta = "Tipos Actividad"
        ''sSQL = ""
        ''sSQL &= "SELECT Code, Name FROM OCLT"
        ''sHANA = ""
        ''sHANA &= "SELECT ""Code"", ""Name"" FROM OCLT"

        ''If Not ExisteixConsulta("Tipos Actividad") Then
        ''    AddQuery(lCategoria, sNomConsulta, CheckIfHana(sSQL, sHANA))
        ''    Application.DoEvents()
        ''End If
        '''
    End Sub
    '
    Private Sub AfegirBusquedesFormatejades()
        '
        ' Formulario Lista de Materiales
        'AddFormQuery(enAddOnFormType.f_ListaMateriales, "Artículos Ventas", "grdgrid", "articulo")
        '-----------------------------------------------------------------------------------------------
    End Sub
    '
    Public Sub EliminarConsultesFormatejades()
        ' EliminarConsultas()
        ' EliminarBusquedas()
    End Sub
    ''
    Private Sub EliminarConsultas()

    End Sub

#End Region
    '
#Region "Funciones Auxiliares"
    Private Function ExistsBusqueda(ByVal sForm As String, ByVal sItem As String, ByVal sCol As String) As Boolean

        If sCol = "" Then sCol = "-1"

        Dim sSQL As String
        Dim sHANA As String
        sSQL = "SELECT IndexID FROM CSHS " & _
               " WHERE FormID = '" & sForm & "' " & _
               " AND ItemID = '" & sItem & "' " & _
               " AND ColID = '" & sCol & "'"

        sHANA = "SELECT ""IndexID"" FROM ""CSHS"" " & _
               " WHERE ""FormID"" = '" & sForm & "' " & _
               " AND ""ItemID"" = '" & sItem & "' " & _
               " AND ""ColID"" = '" & sCol & "'"
        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = _ParentAddon.SBO_Company.GetBusinessObject(BoRecordset)
        oRecordSet.DoQuery(CheckIfHana(sSQL, sHANA))

        If oRecordSet.EoF Then
            ExistsBusqueda = False
        Else
            ExistsBusqueda = True
        End If
        LiberarObjCOM(oRecordSet)
    End Function
    '
    Private Sub InserirBusqueda(ByVal sForm As String, _
                                ByVal sItem As String, _
                                ByVal sCol As String, _
                                ByVal lQueryID As Long, _
                                ByVal ActualizarAutomaticamente As SAPbobsCOM.BoYesNoEnum, _
                                ByVal SiCampoModifica As SAPbobsCOM.BoYesNoEnum, _
                                ByVal CampoQueSeModifica As String, _
                                ByVal ActualizarRegularmente As SAPbobsCOM.BoYesNoEnum)
        '
        Dim oBusqueda As SAPbobsCOM.FormattedSearches
        '
        If sCol = "" Then sCol = "-1"
        '
        oBusqueda = _ParentAddon.SBO_Company.GetBusinessObject(oFormattedSearches)
        With oBusqueda
            .Action = bofsaQuery
            .ByField = SiCampoModifica
            .ColumnID = sCol
            .FieldID = CampoQueSeModifica
            .ForceRefresh = ActualizarRegularmente
            .FormID = sForm
            .ItemID = sItem
            .QueryID = lQueryID
            .Refresh = ActualizarAutomaticamente
        End With
        If oBusqueda.Add <> 0 Then
            Me._ParentAddon.SBO_Application.MessageBox("Ha habido algun problema al intentar asociar una Consulta Formateada (" & lQueryID & ", " & sForm & ", " & sItem & ", " & sCol & "). Causa: " & _ParentAddon.SBO_Company.GetLastErrorDescription)
        End If
        '
    End Sub

    '
    '
    Private Sub InserirConsulta(ByVal lCategoria As Long, ByVal sConsulta As String, ByVal sSQL As String)
        'mana
        Dim oConsulta As SAPbobsCOM.UserQueries
        '
        oConsulta = _ParentAddon.SBO_Company.GetBusinessObject(oUserQueries)
        With oConsulta
            .Query = sSQL
            .QueryCategory = lCategoria
            .QueryDescription = sConsulta
            .Add()
        End With
    End Sub
    '
    Private Function ExisteixConsulta(ByVal sConsulta As String) As Boolean
        Dim sSQL As String
        Dim sHANA As String
        sSQL = "SELECT IntrnalKey FROM OUQR WHERE QName = '" & sConsulta & "'"
        sHANA = "SELECT ""IntrnalKey"" FROM ""OUQR"" WHERE ""QName"" = '" & sConsulta & "'"
        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = _ParentAddon.SBO_Company.GetBusinessObject(BoRecordset)
        oRecordSet.DoQuery(CheckIfHana(sSQL, sHANA))

        If oRecordSet.EoF Then
            ExisteixConsulta = False
        Else
            ExisteixConsulta = True
        End If
        LiberarObjCOM(oRecordSet)
    End Function
    '
    Private Function EliminarConsulta(ByVal sConsulta As String) As Boolean
        '
        Dim sSQL As String
        Dim sHANA As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        '
        sSQL = "DELETE OUQR WHERE QName = '" & sConsulta & "'"
        sHANA = "DELETE ""OUQR"" WHERE ""QName"" = '" & sConsulta & "'"
        oRecordSet = Nothing
        oRecordSet = _ParentAddon.SBO_Company.GetBusinessObject(BoRecordset)
        oRecordSet.DoQuery(CheckIfHana(sSQL, sHANA))
        '
        LiberarObjCOM(oRecordSet)
    End Function
    '
    Private Function AfegirCategoria_SDK() As Long
        '
        Dim lCategoria As Long
        Dim oCategoria As SAPbobsCOM.QueryCategories
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sKey As String
        Dim sSQL As String
        Dim sHANA As String
        '
        sKey = ""
        'Comprovo que no existeixi la categoria
        sSQL = "SELECT CategoryId FROM OQCN WHERE CatName = '" & c_Categoria & "'"
        sHANA = "SELECT ""CategoryId"" FROM ""OQCN"" WHERE ""CatName"" = '" & c_Categoria & "'"
        'oRecordSet = Nothing
        oRecordSet = _ParentAddon.SBO_Company.GetBusinessObject(BoRecordset)
        oRecordSet.DoQuery(CheckIfHana(sSQL, sHANA))

        If oRecordSet.EoF Then
            'Afegeixo la categoria SDK
            oCategoria = _ParentAddon.SBO_Company.GetBusinessObject(oQueryCategories)
            oCategoria.Name = c_Categoria
            If oCategoria.Add = 0 Then
                _ParentAddon.SBO_Company.GetNewObjectCode(sKey)
                lCategoria = NullToLong(sKey)
            End If
            '
            oCategoria = Nothing
        Else
            AfegirCategoria_SDK = oRecordSet.Fields.Item("CategoryID").Value
        End If
        '
        AfegirCategoria_SDK = lCategoria
        '
        LiberarObjCOM(oRecordSet)
    End Function
    '
    Private Sub EliminarBusquedas()
        '
        EliminarBusqueda("139", "U_SEIMailD", "", ACCIO.AccioConsulta)
        EliminarBusqueda("139", "U_SEIMailCC", "", ACCIO.AccioConsulta)
        EliminarBusqueda("139", "U_SEIMailCCO", "", ACCIO.AccioConsulta)
        EliminarBusqueda("139", "U_SEIIMail", "", ACCIO.AccioConsulta)
        EliminarBusqueda("139", "U_SEITipoC", "", ACCIO.AccioConsulta)

        EliminarBusqueda("140", "U_SEIMailD", "", ACCIO.AccioConsulta)
        EliminarBusqueda("140", "U_SEIMailCC", "", ACCIO.AccioConsulta)
        EliminarBusqueda("140", "U_SEIMailCCO", "", ACCIO.AccioConsulta)
        EliminarBusqueda("140", "U_SEIIMail", "", ACCIO.AccioConsulta)
        EliminarBusqueda("140", "U_SEITipoC", "", ACCIO.AccioConsulta)

        EliminarBusqueda("133", "U_SEIMailD", "", ACCIO.AccioConsulta)
        EliminarBusqueda("133", "U_SEIMailCC", "", ACCIO.AccioConsulta)
        EliminarBusqueda("133", "U_SEIMailCCO", "", ACCIO.AccioConsulta)
        EliminarBusqueda("133", "U_SEIIMail", "", ACCIO.AccioConsulta)
        EliminarBusqueda("133", "U_SEITipoC", "", ACCIO.AccioConsulta)

        EliminarBusqueda("142", "U_SEIMailD", "", ACCIO.AccioConsulta)
        EliminarBusqueda("142", "U_SEIMailCC", "", ACCIO.AccioConsulta)
        EliminarBusqueda("142", "U_SEIMailCCO", "", ACCIO.AccioConsulta)
        EliminarBusqueda("142", "U_SEIIMail", "", ACCIO.AccioConsulta)
        EliminarBusqueda("142", "U_SEITipoC", "", ACCIO.AccioConsulta)
        '
    End Sub

    Private Sub EliminarBusqueda(ByVal sForm As String, ByVal sItem As String, ByVal sCol As String, ByVal sAccio As ACCIO)
        '
        Dim sSQL As String
        Dim sHANA As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        '
        If sCol = "" Then sCol = "-1"
        '
        sSQL = ""
        sSQL = sSQL & " DELETE CSHS "
        sSQL = sSQL & " WHERE  FormID='" & sForm & "'"
        sSQL = sSQL & " AND    ItemID='" & sItem & "'"
        sSQL = sSQL & " AND    ColID ='" & sCol & "'"
        '
        sHANA = ""
        sHANA = sHANA & " DELETE ""CSHS"" "
        sHANA = sHANA & " WHERE  ""FormID""='" & sForm & "'"
        sHANA = sHANA & " AND    ""ItemID""='" & sItem & "'"
        sHANA = sHANA & " AND    ""ColID"" ='" & sCol & "'"

        oRecordSet = Nothing
        oRecordSet = _ParentAddon.SBO_Company.GetBusinessObject(BoRecordset)
        oRecordSet.DoQuery(CheckIfHana(sSQL, sHANA))
        '
        'Si és una búsqueda formatejada per valors definits, els valors estaran a la variables "sValors" separats per "|"
        If sAccio = ACCIO.AccioValors Then
            '
            sSQL = ""
            sSQL = sSQL & " DELETE CUVV "
            sSQL = sSQL & " WHERE  FormID='" & sForm & "'"
            sSQL = sSQL & " AND    ItemID='" & sItem & "'"
            sSQL = sSQL & " AND    ColID ='" & sCol & "'"
            '
            sHANA = ""
            sHANA = sHANA & " DELETE ""CUVV"" "
            sHANA = sHANA & " WHERE  ""FormID""='" & sForm & "'"
            sHANA = sHANA & " AND    ""ItemID""='" & sItem & "'"
            sHANA = sHANA & " AND    ""ColID"" ='" & sCol & "'"
            '
            oRecordSet = Nothing
            oRecordSet = _ParentAddon.SBO_Company.GetBusinessObject(BoRecordset)
            oRecordSet.DoQuery(CheckIfHana(sSQL, sHANA))
            '
        End If
        '
        oRecordSet = Nothing
        '
        LiberarObjCOM(oRecordSet)
    End Sub
    Private Sub AddFormQuery(ByVal sFormulario As String, _
                             ByVal sNomConsulta As String, _
                             ByVal sItem As String, _
                             ByVal sCol As String, _
                             Optional ByVal ActualizarAutomaticamente As SAPbobsCOM.BoYesNoEnum = tNO, _
                             Optional ByVal SiCampoModifica As SAPbobsCOM.BoYesNoEnum = tNO, _
                             Optional ByVal CampoQueSeModifica As String = "", _
                             Optional ByVal ActualizarRegularmente As SAPbobsCOM.BoYesNoEnum = tNO)

        Dim sSQL As String
        Dim sHANA As String
        Dim oRecordSet As SAPbobsCOM.Recordset

        If Not ExistsBusqueda(sFormulario, sItem, sCol) Then

            'Busco l'identificador de la consulta---------- Nom Consulta
            sSQL = "SELECT IntrnalKey FROM OUQR WHERE QName = '" & sNomConsulta & "'"
            sHANA = "SELECT ""IntrnalKey"" FROM ""OUQR"" WHERE ""QName"" = '" & sNomConsulta & "'"
            oRecordSet = _ParentAddon.SBO_Company.GetBusinessObject(BoRecordset)
            oRecordSet.DoQuery(CheckIfHana(sSQL, sHANA))

            If Not oRecordSet.EoF Then
                InserirBusqueda(sFormulario, sItem, sCol, oRecordSet.Fields.Item("IntrnalKey").Value, ActualizarAutomaticamente, SiCampoModifica, CampoQueSeModifica, ActualizarRegularmente)
            End If
        End If
    End Sub


    Private Function ContadorCategoriaSDK() As Long
        Dim sSQL As String
        Dim sHANA As String
        Dim oRecordSet As SAPbobsCOM.Recordset

        ContadorCategoriaSDK = -1

        'Busco l'identificador de la categoria SDK
        sSQL = "SELECT CategoryId FROM OQCN WHERE CatName = '" & c_Categoria & "'"
        sHANA = "SELECT ""CategoryId"" FROM ""OQCN"" WHERE ""CatName"" = '" & c_Categoria & "'"
        oRecordSet = _ParentAddon.SBO_Company.GetBusinessObject(BoRecordset)
        oRecordSet.DoQuery(CheckIfHana(sSQL, sHANA))

        If Not oRecordSet.EoF Then
            ContadorCategoriaSDK = oRecordSet.Fields.Item("CategoryId").Value
        End If
        LiberarObjCOM(oRecordSet)
    End Function
    '
    Private Sub AddQuery(ByVal lCategoria As Long, ByVal sNomConsulta As String, ByVal sSQL As String)

        ' lCategoria   -> Categoria SDK
        ' sNomConsulta -> Nombre de la Consulta
        ' sSQL         -> Sentencia SQL

        If Not ExisteixConsulta(sNomConsulta) Then
            InserirConsulta(lCategoria, sNomConsulta, sSQL)
        End If
    End Sub

#End Region
    '
End Class
