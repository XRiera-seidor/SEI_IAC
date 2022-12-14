Option Strict Off
Option Explicit On
'
Imports SEI.SEI_IAC.SEI_AddOnEnum
'
Public Class SEI_AddingMenuItems
    '
#Region "Variables"
    Private oParentAddon As SEI_Addon
    Private sErrNumber As String
    Private iErrMsg As Integer
    '
    Dim oMenus As SAPbouiCOM.Menus
    Dim oMenuItem As SAPbouiCOM.MenuItem
    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams

#End Region

#Region "Contructor"
    Public Sub New(ByRef o_ParentAddon As SEI_Addon)
        oParentAddon = o_ParentAddon
    End Sub
#End Region
    '
#Region "Funciones"

    Public Sub AddMenus()
        If RecuperarValores(oParentAddon.SBO_Company, "U_SEIPEIG", "OUSR", "UserId".Split, oParentAddon.SBO_Company.UserSignature.ToString.Split("|")) = "S" Then
            AddMenu_ErrorGESTION()
        End If
        ''   AddMenu_ModifyHist()

    End Sub

    Private Sub AddMenu_ErrorGESTION()
        '
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        '
        Dim sMenu As String
        Dim sSubMenu As String
        Dim sModulo As String

        '-----------------------------------------------------------------------------
        sModulo = "Gestión" ' Descripcion en el mensaje de Error
        sMenu = "3328"
        sSubMenu = SEI_AddOnEnum.enMenuUID.MNU_SEI_ErrorsGESTION
        '-----------------------------------------------------------------------------

        Try
            oCreationPackage = oParentAddon.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

            oMenuItem = oParentAddon.SBO_Application.Menus.Item(sMenu)            '// modules menu
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = sSubMenu
            oCreationPackage.String = "Errores GESTION"
            oCreationPackage.Position = oMenuItem.SubMenus.Count + 1

            If Not oMenuItem.SubMenus.Exists(sSubMenu) Then
                oMenuItem.SubMenus.AddEx(oCreationPackage)
            End If

        Catch e As System.Exception
            oParentAddon.SBO_Application.StatusBar.SetText("No se ha podido añadir el Sub-Menú " & sMenu & " en el Módulo de " & sModulo, SAPbouiCOM.BoMessageTime.bmt_Short)
            oParentAddon.SBO_Application.MessageBox(e.ToString, 1)
        End Try

    End Sub

    Private Sub AddMenu_ModifyHist()
        '
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        '
        Dim sMenu As String
        Dim sSubMenu As String
        Dim sModulo As String
        '''
        '''-----------------------------------------------------------------------------
        ''sModulo = "Recursos Humanos" ' Descripcion en el mensaje de Error
        ''sMenu = "43544"
        ''sSubMenu = SEI_AddOnEnum.enMenuUID.MNU_SEI_ModifyHist
        '''-----------------------------------------------------------------------------
        '''
        ''Try
        ''    oCreationPackage = oParentAddon.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

        ''    oMenuItem = oParentAddon.SBO_Application.Menus.Item(sMenu)            '// modules menu
        ''    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        ''    oCreationPackage.UniqueID = sSubMenu
        ''    oCreationPackage.String = "Historial Partes"
        ''    oCreationPackage.Position = 14 'oMenuItem.SubMenus.Count + 1
        ''    '
        ''    If Not oMenuItem.SubMenus.Exists(sSubMenu) Then
        ''        oMenuItem.SubMenus.AddEx(oCreationPackage)
        ''    End If
        ''    '
        ''Catch e As System.Exception
        ''    oParentAddon.SBO_Application.StatusBar.SetText("No se ha podido añadir el Sub-Menú " & sMenu & " en el Módulo de " & sModulo, SAPbouiCOM.BoMessageTime.bmt_Short)
        ''    oParentAddon.SBO_Application.MessageBox(e.ToString, 1)
        ''End Try
        '''
    End Sub

#End Region

#Region "Funciones Auxiliares"
    Public Sub DeleteMenuItem(ByVal sMenu As String, _
                              ByVal sSubMenu As String)
        '
        Dim oMenuItem As SAPbouiCOM.MenuItem
        '
        Try
            '
            'oMenuItem = Me.oParentAddon.SBO_Application.Menus.Item(sMenu)
            If Me.oParentAddon.SBO_Application.Menus.Exists(sSubMenu) Then
                oMenuItem = Me.oParentAddon.SBO_Application.Menus.Item(sMenu)
            Else
                Exit Sub
            End If
            '
            If oMenuItem Is Nothing Then
                Exit Sub
            Else
                oMenuItem.SubMenus.RemoveEx(sSubMenu)
            End If
            '
        Catch ex As Exception
            Me.oParentAddon.SBO_Application.MessageBox("No se ha podido Eliminar el Sub-Menu " & sSubMenu)
        End Try
        '
    End Sub

#End Region

End Class
