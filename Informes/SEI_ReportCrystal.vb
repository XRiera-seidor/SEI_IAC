Option Explicit On
'
Imports SEI.SEI_IAC.SEI_AddOnEnum
Imports SEI.SEI_IAC.SEI_Addon
Imports System.Threading
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports System.IO
Imports System.Text
Imports System.Collections
Imports SAPbobsCOM.BoObjectTypes
Imports SEI.SEI_IAC.SEI_Encriptar_Password

Public Class SEI_ReportCrystal
    'Inherits SEI_Form
    Protected _BringToFront As Boolean  ' Poner el formulario delante 
    Protected _ParentAddon As SEI_Addon
    Protected _Where As String
    Protected _Informe As String
    Protected _Formulas As ArrayList
    Protected _SubReports_Where As ArrayList
    Protected _InformeCrystal As CrystalDecisions.CrystalReports.Engine.ReportDocument
    Protected _TipoConstructor As Integer
    Protected _Copias As Integer
    Protected _OrigenDatos As String ' "SQL" , "XML"
    Protected _NomInforme As String
    Protected _Impresora As String
    Protected _stSubReports_Where As st_SubReportWhere
    Protected _Printer As String
    Protected _PrintPDF As Boolean
    Protected _PrintPDF_File As String
    Protected _PrintPDF_FilePath As String
    Public Shared sMasterWord As String = "SEIDOR" ''Paraula de pas per encriptar la contrasenya

    ' Obtener Ruta Fichero PDF
    Public Property PrintPDF_FilePath() As String
        Get
            PrintPDF_FilePath = _PrintPDF_FilePath
        End Get
        Set(ByVal value As String)
            _PrintPDF_FilePath = value
        End Set
    End Property

    ' Nombre del Fichero
    Public Property PrintPDF_File() As String
        Get
            PrintPDF_File = _PrintPDF_File
        End Get
        Set(ByVal value As String)
            _PrintPDF_File = value
        End Set
    End Property

#Region "Constructor"

    Public Sub New(ByRef ParentAddon As SEI_Addon, _
                   ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
                   ByVal sWhere As String, _
                   Optional ByVal bBringToFront As Boolean = True, _
                   Optional ByVal NomInforme As String = "")

        _NomInforme = NomInforme
        _ParentAddon = ParentAddon
        _Where = sWhere
        _InformeCrystal = oInforme
        _TipoConstructor = 1          ' Flag para saber como se ha instanciado el objeto
        _BringToFront = bBringToFront ' Poner el formulario delante 
        '
    End Sub

    Public Sub New(ByRef ParentAddon As SEI_Addon, _
                   ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
                   ByVal sWhere As String, _
                   ByVal aFormulas As ArrayList, _
                   Optional ByVal bBringToFront As Boolean = True, _
                   Optional ByVal NomInforme As String = "")

        _NomInforme = NomInforme
        _ParentAddon = ParentAddon
        _Where = sWhere
        _InformeCrystal = oInforme
        _Formulas = aFormulas
        _TipoConstructor = 2  ' Flag para saber como se ha instanciado el objeto
        _BringToFront = bBringToFront ' Poner el formulario delante 
        '
    End Sub

    Public Sub New(ByRef ParentAddon As SEI_Addon, _
               ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
               ByVal sWhere As String, _
               ByVal aFormulas As ArrayList, _
               ByVal aSubReports_Where As ArrayList, _
               Optional ByVal bBringToFront As Boolean = True, _
                   Optional ByVal NomInforme As String = "")

        _NomInforme = NomInforme
        _ParentAddon = ParentAddon
        _Where = sWhere

        _InformeCrystal = oInforme
        _Formulas = aFormulas
        _SubReports_Where = aSubReports_Where
        _TipoConstructor = 3  ' Flag para saber como se ha instanciado el objeto
        _BringToFront = bBringToFront ' Poner el formulario delante 
        '
    End Sub

    Public Sub New(ByRef ParentAddon As SEI_Addon, _
                   ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
                   ByVal sWhere As String, _
                   ByVal aFormulas As ArrayList, _
                   ByVal iCopias As Integer, _
                   Optional ByVal NomInforme As String = "")

        _NomInforme = NomInforme
        _ParentAddon = ParentAddon
        _Where = sWhere
        _InformeCrystal = oInforme
        _Formulas = aFormulas
        _TipoConstructor = 5  ' Flag para saber como se ha instanciado el objeto
        _Copias = iCopias     ' Nº de Copias
        '
    End Sub

    ''Public Sub New(ByRef ParentAddon As SEI_Addon, _
    ''           ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
    ''           ByVal sOrigenDatos As String, _
    ''           Optional ByVal bBringToFront As Boolean = True, _
    ''        Optional ByVal NomInforme As String = "")

    ''    _NomInforme = NomInforme
    ''    _ParentAddon = ParentAddon
    ''    _InformeCrystal = oInforme
    ''    _OrigenDatos = sOrigenDatos
    ''    _TipoConstructor = 40  ' Flag para saber como se ha instanciado el objeto
    ''    _BringToFront = bBringToFront ' Poner el formulario delante 
    ''    '
    ''End Sub

    ''Public Sub New(ByRef ParentAddon As SEI_Addon, _
    ''          ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
    ''          ByVal sOrigenDatos As String, _
    ''          ByVal aFormulas As ArrayList, _
    ''          Optional ByVal bBringToFront As Boolean = True, _
    ''          Optional ByVal NomInforme As String = "")

    ''    _NomInforme = NomInforme
    ''    _ParentAddon = ParentAddon
    ''    _InformeCrystal = oInforme
    ''    _OrigenDatos = sOrigenDatos
    ''    _TipoConstructor = 41  ' Flag para saber como se ha instanciado el objeto
    ''    _Formulas = aFormulas
    ''    _BringToFront = bBringToFront ' Poner el formulario delante 
    ''    '
    ''End Sub

    Public Sub New(ByRef ParentAddon As SEI_Addon, _
                   ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
                   ByVal stSubReports_Where As st_SubReportWhere, _
                   ByVal aFormulas As ArrayList, _
                   ByVal sImpresora As String, _
                   ByVal iCopias As Integer, _
                   Optional ByVal NomInforme As String = "")

        _NomInforme = NomInforme
        _ParentAddon = ParentAddon
        _stSubReports_Where = stSubReports_Where
        _InformeCrystal = oInforme
        _Formulas = aFormulas
        _TipoConstructor = 6  ' Flag para saber como se ha instanciado el objeto
        _Copias = iCopias     ' Nº de Copias
        _Impresora = sImpresora
        _BringToFront = True
        '
    End Sub

    '
    Public Sub New(ByRef ParentAddon As SEI_Addon, _
                   ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
                   ByVal stSubReportWhere As st_SubReportWhere, _
                   ByVal bPrintPDF As Boolean, _
                   ByVal sPrintPDF_File As String)

        _stSubReports_Where = stSubReportWhere
        _ParentAddon = ParentAddon
        _InformeCrystal = oInforme
        _TipoConstructor = 7  ' Flag para saber como se ha instanciado el objeto
        _PrintPDF = bPrintPDF ' Exportar a PDF
        _PrintPDF_File = sPrintPDF_File     ' Nombre de fichero a exportar a PDF

    End Sub

#End Region

#Region "Funciones"
    '
    Function EncryptedPsswrd() As String
        ''Obtenemos la contraseña encriptada guardada en la tabla SEI_Configuracion
        Dim result As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = _ParentAddon.SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sSQL As String = ""
        Dim sHANA As String = ""

        sSQL = "SELECT U_SEIPsswrd FROM [@SEI_Configuracion]"
        sHANA = "SELECT ""U_SEIPsswrd"" FROM ""@SEI_CONFIGURACION"""
        oRecordSet.DoQuery(CheckIfHana(sSQL, sHANA))
        result = oRecordSet.Fields.Item("U_SEIPsswrd").Value
        EncryptedPsswrd = result
        '
        LiberarObjCOM(oRecordSet)
    End Function
    '
    Function GetUser() As String

        ''Obtenemos el nombre de usuario guardado en la tabla SEI_Configuracion
        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = _ParentAddon.SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sSQL As String = ""
        Dim sHANA As String = ""
        sSQL = "SELECT U_SEIUser FROM [@SEI_Configuracion] "
        sHANA = "SELECT ""U_SEIUser"" FROM ""@SEI_CONFIGURACION"" "
        oRecordSet.DoQuery(CheckIfHana(sSQL, sHANA))
        GetUser = oRecordSet.Fields.Item("U_SEIUser").Value
        LiberarObjCOM(oRecordSet)
    End Function
    '
    Private Function WindowsAuthentication() As Boolean
        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = _ParentAddon.SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sSQL As String = ""
        Dim sHANA As String = ""
        sSQL = "SELECT U_SEIWinAuth FROM [@SEI_Configuracion] "
        sHANA = "SELECT ""U_SEIWinAuth"" FROM ""@SEI_Configuracion"" "
        oRecordSet.DoQuery(CheckIfHana(sSQL, sHANA))
        If oRecordSet.Fields.Item("U_SEIWinAuth").Value = "Y" Then
            Return True
        End If
        Return False
    End Function
    '
    Private Sub LoadReport()
        '  
        Dim stFormula As st_Formulas
        Dim stSubReportWhere As st_SubReportWhere
        Dim stParametro As st_Parametro
        Dim sPath As String = Application.StartupPath
        Dim ReportPath As String = sPath & "\Informes\" & Me._Informe
        '
        Dim myLogin As New CrystalDecisions.Shared.TableLogOnInfo
        Dim myTable As CrystalDecisions.CrystalReports.Engine.Table
        Dim oReport As CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim crFormulas As FormulaFieldDefinitions
        '
        oReport = _InformeCrystal
        '
        'Dim oReport As new CrystalDecisions.CrystalReports.Engine.ReportDocument
        'oReport.Load(ReportPath)
        '
        ''Obtener el nombre de usuario: Lo obtememos a partir de la funcion GetUser()
        'myLogin.ConnectionInfo.UserID = GetUser()
        ''Obtener la contraseña: sMasterWord es la palabra de seguridad (SEIDOR), i EncryptedPsswrd es la contraseña encriptada
        'myLogin.ConnectionInfo.Password = UnEncryptStr(EncryptedPsswrd, sMasterWord)


        ''''TODO: Recuperar el mode en el que està l'aplicació per utilitzar una connexió o l'altra.
        ''''En el cas que sigui HANA i 64bits(AnyCPU) s'haura d'utilitzar la connexió específica.
        '''If b_Hana And IntPtr.Size = 8 Then
        '''    'Dim strConnection As String = "DRIVER= {B1CRHProxy};UID=" + IniGet(sPath & "\S_SEI_IAC.ini", "Parametros", "U") +
        '''    '    ";PWD=" + IniGet(sPath & "\S_SEI_IAC.ini", "Parametros", "P") + ";SERVERNODE=" + Me._ParentAddon.SBO_Company.Server +
        '''    '    ";DATABASE=" + Me._ParentAddon.SBO_Company.CompanyDB + ";"
        '''    Dim strConnection As String = "DRIVER= {B1CRHProxy};UID=" + GetUser() +
        '''       ";PWD=" + UnEncryptStr(EncryptedPsswrd, sMasterWord) + ";SERVERNODE=" + Me._ParentAddon.SBO_Company.Server +
        '''       ";DATABASE=" + Me._ParentAddon.SBO_Company.CompanyDB + ";"

        '''    Dim logonProps2 As NameValuePairs2 = oReport.DataSourceConnections(0).LogonProperties
        '''    logonProps2.Set("Provider", "B1CRHProxy")
        '''    logonProps2.Set("Server Type", "B1CRHProxy")
        '''    logonProps2.Set("Connection String", strConnection)
        '''    logonProps2.Set("Locale Identifier", "1033")
        '''    oReport.DataSourceConnections(0).SetLogonProperties(logonProps2)
        '''    'oReport.DataSourceConnections(0).SetConnection(Me._ParentAddon.SBO_Company.Server, Me._ParentAddon.SBO_Company.CompanyDB, IniGet(sPath & "\S_FINESTRA.ini", "Parametros", "U"), IniGet(sPath & "\S_FINESTRA.ini", "Parametros", "P"))
        '''    oReport.DataSourceConnections(0).SetConnection(Me._ParentAddon.SBO_Company.Server, Me._ParentAddon.SBO_Company.CompanyDB, GetUser, UnEncryptStr(EncryptedPsswrd, sMasterWord))
        '''    'ElseIf IniGet(sPath & "\S_SEI_IAC.ini", "Parametros", "IS") = "Y" Then
        '''ElseIf WindowsAuthentication() Then
        '''    myLogin.ConnectionInfo.IntegratedSecurity = True
        '''    myLogin.ConnectionInfo.DatabaseName = Me._ParentAddon.SBO_Company.CompanyDB
        '''    myLogin.ConnectionInfo.ServerName = Me._ParentAddon.SBO_Company.Server
        '''    'myLogin.ConnectionInfo.UserID = IniGet(sPath & "\S_SEI_IAC.ini", "Parametros", "U")
        '''    'myLogin.ConnectionInfo.Password = IniGet(sPath & "\S_SEI_IAC.ini", "Parametros", "P")
        '''Else
        '''    'myLogin.ConnectionInfo.UserID = IniGet(sPath & "\S_SEI_IAC.ini", "Parametros", "U")
        '''    'myLogin.ConnectionInfo.Password = IniGet(sPath & "\S_SEI_IAC.ini", "Parametros", "P")
        '''    myLogin.ConnectionInfo.UserID = GetUser()
        '''    myLogin.ConnectionInfo.Password = UnEncryptStr(EncryptedPsswrd, sMasterWord)
        '''    myLogin.ConnectionInfo.DatabaseName = Me._ParentAddon.SBO_Company.CompanyDB
        '''    myLogin.ConnectionInfo.ServerName = Me._ParentAddon.SBO_Company.Server
        '''End If
        '''
        myLogin.ConnectionInfo.UserID = IniGet(sPath & "\S_SEI_IAC.ini", "Parametros", "U")
        myLogin.ConnectionInfo.Password = IniGet(sPath & "\S_SEI_IAC.ini", "Parametros", "P")
        myLogin.ConnectionInfo.DatabaseName = Me._ParentAddon.SBO_Company.CompanyDB
        myLogin.ConnectionInfo.ServerName = Me._ParentAddon.SBO_Company.Server

        '
        '-----------------------------------------------------------------------------------
        ' Conexion Tablas
        '-----------------------------------------------------------------------------------
        For Each myTable In oReport.Database.Tables
                myTable.ApplyLogOnInfo(myLogin)
                'myTable.Location = "@SEI"
                'objReport.Database.Tables("MyTable").SetDataSource(objDataSet.Tables("MyTable"))    
            Next
            '-----------------------------------------------------------------------------------
            ' Formulas
            '-----------------------------------------------------------------------------------
            If Not IsNothing(_Formulas) Then
                crFormulas = oReport.DataDefinition.FormulaFields
                '
                For Each stFormula In _Formulas
                    crFormulas.Item(stFormula.Nombre).Text = "'" & stFormula.Valor.ToString & "'"
                Next
            End If
            '
            '-----------------------------------------------------------------------------------
            ' SubReports
            '-----------------------------------------------------------------------------------
            'Crystal Report's report document object
            'Dim objReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            'objReport.VerifyDatabase()

            'Sub report object of crystal report.
            'Dim mySubReportObject As CrystalDecisions.CrystalReports.Engine.SubreportObject

            'Sub report document of crystal report.
            Dim mySubRepDoc As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            '
            For Each mySubRepDoc In oReport.Subreports
                For Each myTable In mySubRepDoc.Database.Tables
                    myTable.ApplyLogOnInfo(myLogin)
                Next
                '
                ' Where Subinformes
                '
                If Not IsNothing(_SubReports_Where) Then
                    For Each stSubReportWhere In _SubReports_Where
                        If stSubReportWhere.NombreReport = mySubRepDoc.Name Then
                            '-----------------------------------------------------------------------------------
                            ' Formulas
                            '-----------------------------------------------------------------------------------
                            If Not IsNothing(stSubReportWhere.aFormulas) Then
                                crFormulas = mySubRepDoc.DataDefinition.FormulaFields
                                '
                                For Each stFormula In stSubReportWhere.aFormulas
                                    crFormulas.Item(stFormula.Nombre).Text = stFormula.Valor.ToString
                                Next
                            End If

                            'mySubRepDoc.DataDefinition.RecordSelectionFormula = stSubReportWhere.ValorWhere
                        End If
                    Next
                End If
            Next
        '-----------------------------------------------------------------------------------
        '

        If Not IsNothing(_stSubReports_Where.aParametros) Then
            For Each stParametro In _stSubReports_Where.aParametros
                '-> Abans de passar el paràmetre al informe mirem si aquest informe el té
                If Not oReport.ParameterFields.Find(stParametro.FieldName, "") Is Nothing Then
                    oReport.SetParameterValue(stParametro.FieldName, stParametro.Value.Split("|"))
                End If
            Next
        Else
            ' Si no tiene Parametros hay que rellenar la formula de seleccion
            oReport.DataDefinition.RecordSelectionFormula = Me._Where
        End If

        If Me._PrintPDF = True Then
            ''Dim sDirectorio As String
            ''sDirectorio = Application.StartupPath
            ''''     Me.PrintPDF_FilePath = sDirectorio & "\" & Me._PrintPDF_File & ".pdf"
            Me.PrintPDF_FilePath = Me._PrintPDF_File
            '
            oReport.ExportToDisk(ExportFormatType.PortableDocFormat, Me.PrintPDF_FilePath)
        Else
            If NullToText(_Impresora).Trim <> "" Then
                    oReport.PrintOptions.PrinterName = _Impresora
                End If
                '
                If Me._Copias <> 0 Then
                    oReport.PrintToPrinter(Me._Copias, True, 0, 0)
                Else
                    oReport.PrintToPrinter(1, True, 0, 0)
                End If
            End If
            '
            oReport.Close()
            '
    End Sub

    Public Sub Imprimir()
        '
        Dim sError As String = ""

        'Me._ParentAddon.SBO_Application.StatusBar.SetText("Cargando documento, un momento por favor...", _
        '                    SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        '
        ' Variable para controlar el estado de la impresión
        Dim bPrintOk As Boolean = True
        '
        Try
            '
            LoadReport()
            '
        Catch ex As Exception
            '
            ' Si salta cualquier excepción en el proceso de impresión, ponemos la variable a falso
            bPrintOk = False
            sError = ex.Message
            PrintPDF_FilePath = ""
            Me._ParentAddon.SBO_Application.MessageBox(ex.ToString)
            '
        End Try
        '
        ' Mostramos al usuario el resultado de la impresión de la oferta
        '
        If bPrintOk Then
            Me._ParentAddon.SBO_Application.StatusBar.SetText("Documento impreso correctamente", _
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Else
            Me._ParentAddon.SBO_Application.StatusBar.SetText("Ha ocurrido un error al imprimir el documento. Causa:" & sError, _
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
        '
    End Sub

    Public Sub VistaPrevia()
        '
        Me._ParentAddon.SBO_Application.StatusBar.SetText("Previsualización en curso, un momento por favor...", _
                            SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        '
        ' Variable para controlar el estado de la impresión
        Dim bPrintOk As Boolean = True
        Dim oFormVisor As SEI_VisorCrystal
        '
        Try
            '
            Select Case Me._TipoConstructor
                '
                Case Is = 1
                    oFormVisor = New SEI_VisorCrystal(Me._ParentAddon, Me._InformeCrystal, Me._Where, _NomInforme)
                    If Me._BringToFront Then
                        Dim oWindowsSbo As SEI_WindowsSbo
                        oWindowsSbo = New SEI_WindowsSbo(Me._ParentAddon)
                        oFormVisor.ShowDialog(oWindowsSbo)
                    Else
                        oFormVisor.ShowDialog()
                        oFormVisor.Dispose()
                        oFormVisor = Nothing
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    End If
                    '
                Case Is = 2
                    oFormVisor = New SEI_VisorCrystal(Me._ParentAddon, Me._InformeCrystal, Me._Where, Me._Formulas, _NomInforme)
                    If Me._BringToFront Then
                        Dim oWindowsSbo As SEI_WindowsSbo
                        oWindowsSbo = New SEI_WindowsSbo(Me._ParentAddon)
                        oFormVisor.ShowDialog(oWindowsSbo)
                        oFormVisor = Nothing
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    Else
                        oFormVisor.ShowDialog()
                        oFormVisor.Dispose()
                        oFormVisor = Nothing
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    End If
                    '
                Case Is = 3
                    oFormVisor = New SEI_VisorCrystal(Me._ParentAddon, Me._InformeCrystal, Me._Where, Me._Formulas, Me._SubReports_Where, _NomInforme)
                    If Me._BringToFront Then
                        Dim oWindowsSbo As SEI_WindowsSbo
                        oWindowsSbo = New SEI_WindowsSbo(Me._ParentAddon)
                        oFormVisor.ShowDialog(oWindowsSbo)
                        oFormVisor.Dispose()
                        oFormVisor = Nothing
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    Else
                        oFormVisor.ShowDialog()
                        oFormVisor.Dispose()
                        oFormVisor = Nothing
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    End If

                Case Is = 6
                    oFormVisor = New SEI_VisorCrystal(Me._ParentAddon, Me._InformeCrystal, Me._Where, Me._Formulas, Me._SubReports_Where, Me._stSubReports_Where, _NomInforme)
                    If Me._BringToFront Then
                        Dim oWindowsSbo As SEI_WindowsSbo
                        oWindowsSbo = New SEI_WindowsSbo(Me._ParentAddon)
                        oFormVisor.ShowDialog(oWindowsSbo)
                        oFormVisor.Dispose()
                        oFormVisor = Nothing
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    Else
                        oFormVisor.ShowDialog()
                        oFormVisor.Dispose()
                        oFormVisor = Nothing
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    End If
                    '
                    'Case Is = 40
                    '    oFormVisor = New SEI_VisorCrystal(Me._ParentAddon, Me._Informe, Me._OrigenDatos, Me._DsDatos)
                    '    If Me._BringToFront Then
                    '        Dim oWindowsSbo As SEI_WindowsSbo
                    '        oWindowsSbo = New SEI_WindowsSbo(Me._ParentAddon)
                    '        oFormVisor.ShowDialog(oWindowsSbo)
                    '        oFormVisor.Dispose()
                    '        oFormVisor = Nothing
                    '        GC.Collect()
                    '        GC.WaitForPendingFinalizers()
                    '    Else
                    '        oFormVisor.ShowDialog()
                    '        oFormVisor.Dispose()
                    '        oFormVisor = Nothing
                    '        GC.Collect()
                    '        GC.WaitForPendingFinalizers()
                    '    End If

                    'Case Is = 41
                    '    oFormVisor = New SEI_VisorCrystal(Me._ParentAddon, Me._InformeCrystal, Me._OrigenDatos, Me._DsDatos, Me._Formulas)
                    '    If Me._BringToFront Then
                    '        Dim oWindowsSbo As SEI_WindowsSbo
                    '        oWindowsSbo = New SEI_WindowsSbo(Me._ParentAddon)
                    '        oFormVisor.ShowDialog(oWindowsSbo)
                    '        oFormVisor.Dispose()
                    '        oFormVisor = Nothing
                    '        GC.Collect()
                    '        GC.WaitForPendingFinalizers()
                    '    Else
                    '        oFormVisor.ShowDialog()
                    '        oFormVisor.Dispose()
                    '        oFormVisor = Nothing
                    '        GC.Collect()
                    '        GC.WaitForPendingFinalizers()
                    '    End If
                    '
            End Select
            '
        Catch ex As Exception
            '
            ' Si salta cualquier excepción en el proceso de impresión, ponemos la variable a falso
            bPrintOk = False
            Me._ParentAddon.SBO_Application.MessageBox(ex.Message)
            '
        End Try
        '
        ' Mostramos al usuario el resultado de la impresión de la oferta
        '
        If bPrintOk Then
            Me._ParentAddon.SBO_Application.StatusBar.SetText("Documento visualizado correctamente", _
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Else
            Me._ParentAddon.SBO_Application.StatusBar.SetText("Ha ocurrido un error al visualizar el documento", _
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
        '
    End Sub

#End Region

    Function m_SBO_Company() As SAPbobsCOM.Company
        Throw New NotImplementedException
    End Function

End Class
