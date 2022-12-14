'
Imports SEI.SEI_IAC.SEI_AddOnEnum
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Text
Imports System.Collections

Public Class SEI_VisorCrystal
    Protected _ParentAddon As SEI_Addon
    Protected _Where As String
    Protected _Formulas_SubReport As ArrayList
    Protected _Informe As String
    Protected _InformeCrystal As CrystalDecisions.CrystalReports.Engine.ReportDocument
    Protected _Formulas As ArrayList
    Protected _SubReport As ArrayList
    Protected _NomInforme As String
    Protected _stSubReports_Where As st_SubReportWhere

#Region "Constructor"
    Public Sub New(ByRef ParentAddon As SEI_Addon, ByVal oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal sWhere As String, _
                   Optional ByVal NomInforme As String = "")

        _NomInforme = NomInforme

        _ParentAddon = ParentAddon
        _Where = sWhere
        _InformeCrystal = oInforme

        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub

    Public Sub New(ByRef ParentAddon As SEI_Addon, _
                   ByVal oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
                   ByVal sWhere As String, _
                   ByVal aFormulas As ArrayList, _
                   Optional ByVal NomInforme As String = "")

        _NomInforme = NomInforme
        _ParentAddon = ParentAddon
        _Where = sWhere
        _InformeCrystal = oInforme
        _Formulas = aFormulas


        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub

    Public Sub New(ByRef ParentAddon As SEI_Addon, _
                  ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
                  ByVal sWhere As String, _
                  ByVal aFormulas As ArrayList, _
                  ByVal aFormulas_SubReport As ArrayList, _
                   Optional ByVal NomInforme As String = "")

        _NomInforme = NomInforme
        _ParentAddon = ParentAddon
        _Where = sWhere
        _InformeCrystal = oInforme
        _Formulas = aFormulas
        _Formulas_SubReport = aFormulas_SubReport

        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub

    Public Sub New(ByRef ParentAddon As SEI_Addon, _
                  ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
                  ByVal sWhere As String, _
                  ByVal aFormulas As ArrayList, _
                  ByVal aFormulas_SubReport As ArrayList, _
                  ByVal stSubReports_Where As st_SubReportWhere, _
                   Optional ByVal NomInforme As String = "")

        _NomInforme = NomInforme
        _ParentAddon = ParentAddon
        _Where = sWhere
        _InformeCrystal = oInforme
        _Formulas = aFormulas
        _Formulas_SubReport = aFormulas_SubReport
        _stSubReports_Where = stSubReports_Where

        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub


#End Region

    Public Sub ConfigureCrystalReports()
        '  
        Dim stFormula As st_Formulas
        Dim stFormula_S As st_Formulas_SubReport
        Dim sPath As String = System.Windows.Forms.Application.StartupPath
        Dim ReportPath As String = sPath & "\Informes\" & Me._Informe
        '
        Dim myLogin As New CrystalDecisions.Shared.TableLogOnInfo
        Dim myTable As CrystalDecisions.CrystalReports.Engine.Table
        Dim oReport As CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim crFormulas As FormulaFieldDefinitions
        Dim crFormulas_S As FormulaFieldDefinitions

        Dim stParametro As st_Parametro
        '
        ''oReport.Load(ReportPath)
        oReport = _InformeCrystal
        '''Escriure_Ficher_TXT("E:\SEIDOR\LOG_LoginCrystal.log", "")
        '''Escriure_Ficher_TXT("E:\SEIDOR\LOG_LoginCrystal.log", Now.ToString("dd/MM/yyyy hh:mm:ss") & " -> Informe: " & _InformeCrystal.ToString & ": ")

        myLogin.ConnectionInfo.ServerName = Me._ParentAddon.SBO_Company.Server
        '''Escriure_Ficher_TXT("E:\SEIDOR\LOG_LoginCrystal.log", "     Servidor: " & myLogin.ConnectionInfo.ServerName.ToString)
        myLogin.ConnectionInfo.UserID = IniGet(sPath & "\S_SEI_IAC.ini", "Parametros", "U")    ' Usuario
        '''Escriure_Ficher_TXT("E:\SEIDOR\LOG_LoginCrystal.log", "     UserID: " & myLogin.ConnectionInfo.UserID.ToString)
        myLogin.ConnectionInfo.Password = IniGet(sPath & "\S_SEI_IAC.ini", "Parametros", "P")  ' Password
        '''Escriure_Ficher_TXT("E:\SEIDOR\LOG_LoginCrystal.log", "     Password: " & myLogin.ConnectionInfo.Password.ToString)
        myLogin.ConnectionInfo.DatabaseName = Me._ParentAddon.SBO_Company.CompanyDB
        '''Escriure_Ficher_TXT("E:\SEIDOR\LOG_LoginCrystal.log", "     DatabaseName: " & myLogin.ConnectionInfo.DatabaseName.ToString)
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

        'Sub report object of crystal report.
        'Dim mySubReportObject As CrystalDecisions.CrystalReports.Engine.SubreportObject

        'Sub report document of crystal report.
        Dim mySubRepDoc As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        For Each mySubRepDoc In oReport.Subreports
            For Each myTable In mySubRepDoc.Database.Tables
                myTable.ApplyLogOnInfo(myLogin)
            Next
            crFormulas_S = mySubRepDoc.DataDefinition.FormulaFields
            '
            For Each stFormula_S In _Formulas_SubReport
                crFormulas_S.Item(stFormula_S.Nombre).Text = "" & stFormula_S.Valor.ToString & ""
            Next

        Next
        '-----------------------------------------------------------------------------------
        '
        If Not IsNothing(_stSubReports_Where.aParametros) Then
            For Each stParametro In _stSubReports_Where.aParametros
                oReport.SetParameterValue(stParametro.FieldName, stParametro.Value.Split("|"))
            Next
        Else
            ' Si no tiene Parametros hay que rellenar la formula de seleccion
            oReport.DataDefinition.RecordSelectionFormula = Me._Where
        End If

        ''''-> ORDENACIÓ
        '''oReport.DataDefinition.SortFields(0).Field = oReport.Database.Tables("OWOR").Fields("DocEntry")
        '''oReport.DataDefinition.SortFields(0).SortDirection = SortDirection.AscendingOrder

        '''oReport.DataDefinition.SortFields(1).Field = oReport.Database.Tables("WOR1").Fields("ItemCode")
        '''oReport.DataDefinition.SortFields(1).SortDirection = SortDirection.DescendingOrder
        '
        CrystalReportViewer.ShowExportButton = True
        CrystalReportViewer.ShowGroupTreeButton = False
        CrystalReportViewer.ReportSource = oReport
        '
    End Sub
    '
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub CrystalReportViewer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CrystalReportViewer.Load
        ConfigureCrystalReports()
        Me.WindowState = FormWindowState.Maximized

    End Sub

End Class