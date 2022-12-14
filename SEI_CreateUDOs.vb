'
Option Explicit On
'
Imports SAPbobsCOM
Imports SAPbobsCOM.BoFldSubTypes
Imports SAPbobsCOM.BoObjectTypes
Imports SAPbobsCOM.BoFieldTypes

Public Class SEI_CreateUDOs

#Region "Variables"

    Protected m_ParentAddon As SEI_Addon
    
#End Region

#Region "Constructor"

    Public Sub New(ByRef ParentAddon As SEI_Addon)
        m_ParentAddon = ParentAddon
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

#Region "Funciones"

    Public Sub AddUserDefinedObjects()
        GC.Collect()
        Dim oUDO As SAPbobsCOM.UserObjectsMD
        Dim version As String

        Try

            ''oUDO = m_ParentAddon.SBO_Company.GetBusinessObject(oUserObjectsMD)
            ''version = m_ParentAddon.SBO_Company.Version
            '''-> Configuración Envio Mails
            ''If Not oUDO.GetByKey("SEI_ConfMail".ToUpper) Then
            ''    oUDO = m_ParentAddon.SBO_Company.GetBusinessObject(oUserObjectsMD)
            ''    oUDO.Code = "SEI_ConfMail".ToUpper
            ''    oUDO.Name = "Configuración Envio Mails"
            ''    oUDO.TableName = "SEI_ConfMail".ToUpper
            ''    oUDO.ObjectType = BoUDOObjType.boud_MasterData
            ''    'Comprovem si la versió de SAP és inferior a la 882, ja que no tenen les següents propietats en l'objecte UDO
            ''    If Left(CInt(version), 3) > 881 Then
            ''        oUDO.MenuCaption = "Configuración Envio Mails"
            ''        oUDO.MenuItem = BoYesNoEnum.tYES
            ''        oUDO.MenuUID = "SEI_CONFMAIL"
            ''        oUDO.FatherMenuID = "8192"
            ''        oUDO.Position = 11
            ''        oUDO.EnableEnhancedForm = BoYesNoEnum.tNO
            ''    End If
            ''    '
            ''    oUDO.ManageSeries = BoYesNoEnum.tNO
            ''    oUDO.CanCancel = BoYesNoEnum.tNO
            ''    oUDO.CanClose = BoYesNoEnum.tNO
            ''    oUDO.CanDelete = BoYesNoEnum.tYES
            ''    oUDO.CanFind = BoYesNoEnum.tNO
            ''    oUDO.CanLog = BoYesNoEnum.tNO
            ''    oUDO.CanYearTransfer = BoYesNoEnum.tNO

            ''    oUDO.CanCreateDefaultForm = BoYesNoEnum.tYES
            ''    oUDO.FormColumns.FormColumnAlias = "Code"
            ''    oUDO.FormColumns.FormColumnDescription = "Id"
            ''    oUDO.FormColumns.Add()
            ''    oUDO.FormColumns.FormColumnAlias = "U_SEITipoD"
            ''    oUDO.FormColumns.FormColumnDescription = "Tipo Documento"
            ''    oUDO.FormColumns.Add()
            ''    oUDO.FormColumns.FormColumnAlias = "U_SEICrear"
            ''    oUDO.FormColumns.FormColumnDescription = "Mandar Mail al crear"
            ''    oUDO.FormColumns.Add()
            ''    oUDO.FormColumns.FormColumnAlias = "U_SEIBoton"
            ''    oUDO.FormColumns.FormColumnDescription = "Botón de reenviar Mail"
            ''    oUDO.FormColumns.Add()
            ''    oUDO.FormColumns.FormColumnAlias = "U_SEIMasivo"
            ''    oUDO.FormColumns.FormColumnDescription = "Enviar Mail Masivo"
            ''    oUDO.FormColumns.Add()
            ''    oUDO.FormColumns.FormColumnAlias = "U_SEIActiv"
            ''    oUDO.FormColumns.FormColumnDescription = "Tipo de Actividad"
            ''    oUDO.FormColumns.Add()
            ''    oUDO.FormColumns.FormColumnAlias = "U_SEIAsunto"
            ''    oUDO.FormColumns.FormColumnDescription = "Asunto de la Actividad"
            ''    oUDO.FormColumns.Add()
            ''    oUDO.FormColumns.FormColumnAlias = "U_SEIComents"
            ''    oUDO.FormColumns.FormColumnDescription = "Comentarios de la Actividad"
            ''    oUDO.FormColumns.Add()
            ''    oUDO.FormColumns.FormColumnAlias = "U_SEIElimi"
            ''    oUDO.FormColumns.FormColumnDescription = "Eliminar PDF después de enviar"
            ''    oUDO.FormColumns.Add()
            ''    oUDO.FormColumns.FormColumnAlias = "U_SEIRutEsp"
            ''    oUDO.FormColumns.FormColumnDescription = "Ruta PDF Por Documento"
            ''    oUDO.FormColumns.Add()

            ''    oUDO.ChildTables.TableName = "SEI_ConfMail_D"

            ''    If oUDO.Add <> 0 Then
            ''        m_ParentAddon.SBO_Application.MessageBox(m_ParentAddon.SBO_Company.GetLastErrorDescription)
            ''    End If
            ''End If

        Catch ex As Exception
            m_ParentAddon.SBO_Application.MessageBox(ex.Message)
        Finally
            LiberarObjCOM(oUDO)
        End Try

    End Sub

#End Region

#Region "Funciones Auxiliares"

#End Region

End Class
