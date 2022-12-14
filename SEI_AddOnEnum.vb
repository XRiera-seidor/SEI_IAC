Imports System.Collections

Public Class SEI_AddOnEnum

#Region "Constantes Formularios SAP y Add-on"
    '
    Public Structure enAddOnFormType
        '
        ' Enumeraciones Formularios
        '
        Const f_FormulariosSBO_0 As String = "0"
        '-> Ventes
        Const f_OfertasVentas As String = "149"
        Const f_PedidosVentas As String = "139"
        Const f_EntregaVentas As String = "140"
        Const f_DevolucionVentas As String = "180"
        Const f_FacturaVentas As String = "133"
        Const f_FacturaAnticipoVentas As String = "65300"
        Const f_AbonoVentas As String = "179"

        '-> Compres        
        Const f_SolicitudCompras As String = "1470000200"
        Const f_SolicitudPedido As String = "540000988"
        Const f_PedidosCompras As String = "142"
        Const f_EntradaMercanciasCompras As String = "143"
        Const f_DevolucionCompras As String = "182"
        Const f_FacturaCompras As String = "141"
        Const f_AbonoCompras As String = "181"

        '-> Altres
        Const f_Asiento As String = "392"
        Const f_Articulos As String = "150"
        Const f_InterlocutoresComerciales As String = "134"
        Const f_Traslados As String = "940"
        Const f_ReciboProduccion As String = "65214"
        Const f_EmisioProduccion As String = "65213"
        Const f_OrdenFabricacion As String = "65211"
        Const f_EmpleadosVentas As String = "666"
        Const f_Alertas As String = "198"
        Const f_ChooseFromList_ITEMS As String = "10003"
        Const f_Portes As String = "3007"
        Const f_Pagos As String = "426"
        Const f_Errors_GESTION As String = "2000713003"
        ''        Const f_ModifyHist As String = "20000123"


        Const f_CFL_Articulos As String = "10003"


        Public EnumNew As String
    End Structure
    '
#End Region

#Region "Constantes Menus Sap y Add-on"
    '
    Public Structure enMenuUID

        ' Menus Add-on
        '
        Const MNU_Siguiente As String = "1288"
        Const MNU_Anterior As String = "1289"
        Const MNU_Primero As String = "1290"
        Const MNU_Ultimo As String = "1291"
        '
        Const MNU_Buscar As String = "1281"
        Const MNU_Crear As String = "1282"
        Const MNU_Eliminar As String = "1283"
        Const MNU_EliminarLinia As String = "1293"
        Const MNU_A�adirLinia As String = "1292"
        Const MNU_Duplicar As String = "1287"
        Const MNU_Cancelar As String = "1284"
        Const MNU_Cerrar As String = "1286"           ' Cerrar documento
        Const MNU_CerrarLinea As String = "1299"           ' Cerrar L�nea
        Const MNU_ShiftF2 As String = "7425"
        Const MNU_Restablecer As String = "1285"
        Const MNU_Proyectos As String = "8457"
        Const MNU_Catalogo_IC As String = "12545"

        Const MNU_Filtrar As String = "4870"
        '
        Const MNU_Cortar As String = "771"
        Const MNU_Copiar As String = "772"
        Const MNU_Pegar As String = "773"
        Const MNU_Borrar As String = "774"
        '
        Const MNU_SEI_ErrorsGESTION As String = "SEI_ErrorsGESTION"
        '' Const MNU_SEI_ModifyHist As String = "SEI_ModifyHist"

        Public EnumNew As String
    End Structure

#End Region

#Region "Constantes Tipos de Formulario Add-on"

    Public Enum enSBO_LoadFormTypes
        XmlFile = 0       ' Formularios .srf (xml) 
        LogicOnly = 1     ' Formulario de Sap 
        GuiByCode = 3
    End Enum

#End Region

#Region "Estructura VisorCrystal Formulas"
    Public Structure st_Formulas

        Public Nombre As String
        Public Valor As String

    End Structure

    Public Structure st_Parametro

        Public FieldName As String
        Public ParameterRange As CrystalDecisions.Shared.ParameterRangeValue
        Public Value As String

    End Structure

    Public Structure st_SubReportWhere

        Public NombreReport As String
        Public ValorWhere As String
        Public aParametros As ArrayList
        Public aFormulas As ArrayList

    End Structure

    Public Structure st_Formulas_SubReport

        Public Nombre As String
        Public Valor As String

    End Structure

    Public Structure st_SubReport

        Public Nombre As String
        Public Valor As String

    End Structure


    Public Enum eCrystal
        Incrustado = 0
        EnDirectorio = 1
        CampoBlob = 2
    End Enum

    ' En el proceso de facturaci�n Electronica
    Public Enum eEnviarFE
        EnDirectorio = 0
        Mail = 1
        Impresora = 2
    End Enum


#Region "Estructura Email"

    Public Structure st_Email

        Public EmailFrom As String
        Public EmailTo As String
        Public EmailSubject As String
        Public EmailBody As String
        Public aPDF As ArrayList
        Public aMails As ArrayList
        Public aAlbaranes As ArrayList
        Public Usuario As String
        Public Password As String
        Public ServidorSMTP As String
    End Structure

#End Region

#End Region

End Class
