Option Explicit On
'
Imports SEI.SEI_IAC.SEI_AddOnEnum

Public MustInherit Class SEI_Form

    Inherits System.Object

#Region "Variables"
    Private Shared m_intFormCount As Integer
    Protected m_ParentAddon As SEI_Addon
    Protected m_FormUID As String
    Protected m_SBO_Form As SAPbouiCOM.Form
    Protected m_SBO_LoadFormType As enSBO_LoadFormTypes
    Protected m_Formtype As String
    Protected UniqueIdentifier As String

#End Region

#Region "Constructor"
    '
    Public Sub New(ByRef ParentAddon As SEI_Addon, _
               ByVal SBO_LoadFormType As SEI_AddOnEnum.enSBO_LoadFormTypes, _
               ByVal pst_FormType As String, _
               Optional ByVal FormUID As String = "", Optional ByVal UDOName As String = "")
        Try
            m_FormUID = FormUID
            m_ParentAddon = ParentAddon
            m_intFormCount += 1

            m_SBO_LoadFormType = SBO_LoadFormType
            m_Formtype = pst_FormType

            Select Case m_SBO_LoadFormType
                Case enSBO_LoadFormTypes.LogicOnly
                    'set SAP form  
                    ' event caught by addon is FORM_LOAD, update existing form logic only
                    m_SBO_Form = m_ParentAddon.SBO_Application.Forms.Item(m_FormUID)

                Case enSBO_LoadFormTypes.GuiByCode
                    'm_SBO_Form = m_ParentAddon.CreateSBO_FormByCode(FormType, UDOName)
                    'm_FormUID = m_SBO_Form.UniqueID
                Case enSBO_LoadFormTypes.XmlFile
                    'set SAP form  
                    ' event caught by addon is FORM_LOAD, update existing form 
                    ' or create custom form via xml file
                    'If m_FormUID = "" Then
                    m_SBO_Form = m_ParentAddon.CreateSBO_Form(FormUID, Me.FormType, UDOName)
                    m_SBO_Form.Visible = True
                    m_FormUID = m_SBO_Form.UniqueID
                    'Else
                    '    'm_SBO_Form = m_ParentAddon.UpdateSBO_Form(Me.ToString(), m_FormUID)
                    'End If
            End Select
            m_ParentAddon.col_SBOFormsOpened.Add(Me, m_FormUID)

        Catch excE As Exception
            m_ParentAddon.SBO_Application.MessageBox(excE.Message.ToString)
        End Try
    End Sub

#End Region

#Region "Propiedades"
    Public ReadOnly Property FormType() As String
        Get
            Return m_Formtype
        End Get
    End Property
    Public ReadOnly Property SBO_LoadFormType() As enSBO_LoadFormTypes
        Get
            Return m_SBO_LoadFormType
        End Get
    End Property
    Public Shared ReadOnly Property Count() As Integer
        Get
            Return m_intFormCount
        End Get
    End Property

    Property Form() As SAPbouiCOM.Form
        Get
            Form = m_SBO_Form
            'UID = Form.UniqueID
        End Get
        Set(ByVal value As SAPbouiCOM.Form)
            m_SBO_Form = value
        End Set
    End Property

    ReadOnly Property UniqueID() As String
        Get
            Return m_FormUID
        End Get
    End Property

    Property UID() As String
        Get
            UID = UniqueIdentifier
        End Get
        Set(ByVal value As String)
            UniqueIdentifier = value
        End Set
    End Property

    Protected ReadOnly Property SBO_Application() As SAPbouiCOM.Application
        Get
            Return m_ParentAddon.SBO_Application
        End Get
    End Property

    Protected ReadOnly Property SBO_Company() As SAPbobsCOM.Company
        Get
            Return m_ParentAddon.SBO_Company
        End Get
    End Property

#End Region

#Region "MustOverride"
    Public MustOverride Sub HANDLE_PRINT_EVENT(ByRef eventInfo As SAPbouiCOM.PrintEventInfo, ByRef BubbleEvent As Boolean)
    Public MustOverride Sub HANDLE_REPORT_DATA_EVENT(ByRef eventInfo As SAPbouiCOM.ReportDataInfo, ByRef BubbleEvent As Boolean)
    Public MustOverride Function HANDLE_DATA_EVENT(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) As Integer
    Public MustOverride Sub HANDLE_FORM_EVENTS(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Public MustOverride Sub HANDLE_MENU_EVENTS(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
    Public MustOverride Sub HANDLE_RIGHT_CLICK_EVENTS(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
    Public MustOverride Function ModalFormEventAllowed(ByRef pVal As SAPbouiCOM.MenuEvent) As Boolean
#End Region

#Region "Overrides"
    Public Overrides Function ToString() As String
        Dim aName() As String
        aName = Split(MyBase.ToString(), ".")
        Return aName(aName.Length - 1)
    End Function

    Protected Overrides Sub Finalize()
        'clean up class
        GC.WaitForPendingFinalizers()
        GC.Collect()
        MyBase.Finalize()
    End Sub

#End Region

End Class

