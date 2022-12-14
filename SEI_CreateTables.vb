'

Option Explicit On
'
Imports SEI.SEI_IAC.SEI_Encriptar_Password
Imports SAPbobsCOM
Imports SAPbobsCOM.BoFldSubTypes
Imports SAPbobsCOM.BoObjectTypes
Imports SAPbobsCOM.BoFieldTypes
'
Public Class SEI_CreateTables

#Region "Variables"
    Protected m_ParentAddon As SEI_Addon
    Private _Form As SAPbouiCOM.Form
    '
    Dim lRetCode As Long
    Dim lErrCode As Long
    Dim sErrMsg As String
#End Region

#Region "Constructor"
    Public Sub New(ByRef ParentAddon As SEI_Addon)
        m_ParentAddon = ParentAddon
    End Sub
    Public Sub New(ByRef o_ParentAddon As SEI_Addon, ByRef o_Form As SAPbouiCOM.Form)
        m_ParentAddon = o_ParentAddon
        _Form = o_Form
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region
    '   
#Region "Funciones"

    Protected ReadOnly Property SBO_Company() As SAPbobsCOM.Company
        Get
            Return m_ParentAddon.SBO_Company
        End Get
    End Property

    Public Sub AddUserDefinedData()
        '
        '[PROGRAMAR] (Tablas y Campos de Usuario)

        AddUserTables()
        AddUserFields()
        '[PROGRAMAR] (Contadores Tablas Temporales)
        'CrearContador("SEI_HR")  'Contador Hojas de Ruta

    End Sub
    '
    
    '
    Private Sub AddUserTables()

        ''AddUserTable("SEI_CONFMAIL", "Configuración Envio Mails", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        ''AddUserTable("SEI_CONFMAIL_D", "Config. Envio Mails Detalles", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        ''AddUserTable("SEI_CONFSMTP", "Configuración SMTP", SAPbobsCOM.BoUTBTableType.bott_NoObject)

    End Sub
    '
    Private Sub AddUserFields()
        '
        Dim strTabla As String

        '-------------------------------------------
        '-------- Usuarios --------
        '-------------------------------------------
        strTabla = "OUSR"
        addCampAlpha(strTabla, "SEIPEIG", "Permiso:Errores Import.GESTION", 1, "S|Sí;N|No",, "N")


        '------------------------------------------
        '----- Detalles Empresa-------
        '------------------------------------------
        strTabla = "OADM"
        addCampAlpha(strTabla, "SEI_BD_GESTION", "BBDD GESTION", 50)

        '-------------------------------------------
        '-------- Articles --------
        '-------------------------------------------
        ''strTabla = "OITM"
        ''addCampAlpha(strTabla, "SEIFam", "Familia", 30,,,, "SEIFAMILIA")
        ''addCampAlpha(strTabla, "SEITip", "Tipologia", 30,,,, "SEITIPOL")
        ''addCampAlpha(strTabla, "SEICateg", "Categoria", 30,,,, "SEICATEGORIAS")

        '-------------------------------------------
        '-------- Documents --------
        '-------------------------------------------
        ''strTabla = "OPRQ" '-> Solicitud de compra
        ''addCampAlpha(strTabla, "SEICliente", "Cliente para Oferta de venta", 15)
        ''addCampNumericFloat(strTabla, "SEICoef", "Coeficiente Incremento Oferta", BoFldSubTypes.st_Percentage)

        '-------------------------------------------
        '-------- Categorias --------
        '-------------------------------------------
        ''strTabla = "@SEICATEGORIAS"
        ''addCampAlpha(strTabla, "SEIActiva", "żActiva?", 1, "S|Sí;N|No",, "S")


    End Sub

#End Region
    '
#Region "Funciones Auxiliares"
    '
    Private Sub Preparar_Tabla()
        'Esta funcion inicializa una linia con el nombre de usuario para la tabla @SEI_Configuracion
        Dim cRecordSet As SAPbobsCOM.Recordset
        cRecordSet = SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sSQL As String = ""
        Dim sHANA As String = ""
        sSQL &= "SELECT COUNT(*) FROM [@SEI_CONFIGURACION]"
        sHANA = "SELECT COUNT(*) FROM ""@SEI_CONFIGURACION"""
        cRecordSet.DoQuery(CheckIfHana(sSQL, sHANA))
        If (cRecordSet.Fields.Item(0).Value = 0) Then  ''Si no existe ningun dato en la tabla lo insertamos
            '
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim ls As String = ""
            Dim lss As String = ""
            ls &= "INSERT INTO [@SEI_CONFIGURACION]"
            ls &= " values ('SQL', 'sql', 'sa','', 'N')" '"sa" es el usuario de acceso a la base de datos
            '
            lss &= "INSERT INTO ""@SEI_CONFIGURACION"""
            lss &= " values ('SQL', 'sql', 'sa','', 'N')"
            oRecordSet.DoQuery(CheckIfHana(ls, lss))
            '
        End If
        '
    End Sub
    '
    Private Function AddUserTable(ByVal Nom As String, ByVal Descripcio As String, ByVal Tipus As SAPbobsCOM.BoUTBTableType) As Long
        '
        If Not UserTablesExist(Nom) Then
            Dim oUTables As SAPbobsCOM.UserTablesMD
            oUTables = m_ParentAddon.SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)

            oUTables.TableName = Nom
            oUTables.TableDescription = Descripcio
            oUTables.TableType = Tipus
            lRetCode = oUTables.Add

            If lRetCode <> 0 Then
                m_ParentAddon.SBO_Company.GetLastError(lErrCode, sErrMsg)
                m_ParentAddon.SBO_Application.MessageBox(sErrMsg)
            End If
            '
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUTables)
            '
        End If
        '
    End Function
    Private Function addCampNumeric(ByVal vTaula As String, _
                                    ByVal vCamp As String, _
                                    ByVal vDesc As String, _
                                    ByVal vEditSize As Integer, _
                                    Optional ByVal vValorDefecte As String = "", _
                                    Optional ByVal sLinkTable As String = "") As Long
        '
        'Per afegir un camp numeric.
        'Per exemple:
        '    addCampNumeric "OCRD", "SEI_NED", "Numero cli. EDI"
        '
        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        Dim bTablaUsuario As Boolean
        '
        If InStr(vTaula, "@") <> 0 Then
            vTaula = Replace(vTaula, "@", "")
            bTablaUsuario = True
        End If
        '
        If Not bTablaUsuario Then
            '
            If UserFieldsExist(vTaula, vCamp) = False Then
                oUserFieldsMD = Me.m_ParentAddon.SBO_Company.GetBusinessObject(oUserFields)
                oUserFieldsMD.TableName = vTaula
                oUserFieldsMD.Name = vCamp
                oUserFieldsMD.Description = Left(vDesc, 30)
                oUserFieldsMD.Type = db_Numeric
                oUserFieldsMD.EditSize = vEditSize
                If sLinkTable <> "" Then
                    oUserFieldsMD.LinkedTable = sLinkTable
                End If
                If vValorDefecte <> "" Then
                    oUserFieldsMD.DefaultValue = vValorDefecte
                End If

                lRetCode = oUserFieldsMD.Add
                addCampNumeric = lRetCode
                If lRetCode <> 0 Then
                    Me.m_ParentAddon.SBO_Company.GetLastError(lErrCode, sErrMsg)
                    MsgBox(sErrMsg)
                End If
                '
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
                '
            End If
        Else
            '
            If UserFieldsExistT(vTaula, vCamp) = False Then
                oUserFieldsMD = Me.m_ParentAddon.SBO_Company.GetBusinessObject(oUserFields)
                oUserFieldsMD.TableName = vTaula
                oUserFieldsMD.Name = vCamp
                oUserFieldsMD.Description = Left(vDesc, 30)
                oUserFieldsMD.Type = db_Numeric
                oUserFieldsMD.EditSize = vEditSize
                If vValorDefecte <> "" Then
                    oUserFieldsMD.DefaultValue = vValorDefecte
                End If

                lRetCode = oUserFieldsMD.Add
                addCampNumeric = lRetCode
                If lRetCode <> 0 Then
                    Me.m_ParentAddon.SBO_Company.GetLastError(lErrCode, sErrMsg)
                    MsgBox(sErrMsg)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
                '
            End If
        End If
        '
    End Function

    Private Function addCampNumericFloat(ByVal vTaula As String, _
                    ByVal vCamp As String, _
                    ByVal vDesc As String, _
                    ByVal vSubtipus As BoFldSubTypes, _
                    Optional ByVal vValorDefecte As String = "") As Long
        '
        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        Dim bTablaUsuario As Boolean
        '
        If InStr(vTaula, "@") <> 0 Then
            vTaula = Replace(vTaula, "@", "")
            bTablaUsuario = True
        End If
        '
        If Not bTablaUsuario Then
            '
            If UserFieldsExist(vTaula, vCamp) = False Then
                oUserFieldsMD = Me.m_ParentAddon.SBO_Company.GetBusinessObject(oUserFields)
                oUserFieldsMD.TableName = vTaula
                oUserFieldsMD.Name = vCamp
                oUserFieldsMD.Description = Left(vDesc, 30)
                oUserFieldsMD.Type = db_Float
                oUserFieldsMD.SubType = vSubtipus
                If vValorDefecte <> "" Then
                    oUserFieldsMD.DefaultValue = vValorDefecte
                End If

                lRetCode = oUserFieldsMD.Add
                addCampNumericFloat = lRetCode
                If lRetCode <> 0 Then
                    Me.m_ParentAddon.SBO_Company.GetLastError(lErrCode, sErrMsg)
                    MsgBox(sErrMsg)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
                '
            End If
        Else
            '
            If UserFieldsExistT(vTaula, vCamp) = False Then
                oUserFieldsMD = Me.m_ParentAddon.SBO_Company.GetBusinessObject(oUserFields)
                oUserFieldsMD.TableName = vTaula
                oUserFieldsMD.Name = vCamp
                oUserFieldsMD.Description = Left(vDesc, 30)
                oUserFieldsMD.Type = db_Float
                oUserFieldsMD.SubType = vSubtipus
                If vValorDefecte <> "" Then
                    oUserFieldsMD.DefaultValue = vValorDefecte
                End If

                lRetCode = oUserFieldsMD.Add
                addCampNumericFloat = lRetCode
                If lRetCode <> 0 Then
                    Me.m_ParentAddon.SBO_Company.GetLastError(lErrCode, sErrMsg)
                    MsgBox(Me.m_ParentAddon.SBO_Company.GetLastErrorDescription)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
                '
            End If
        End If
        '
    End Function
    '
    Private Function addCampMemo(ByVal vTaula As String, _
                    ByVal vCamp As String, _
                    ByVal vDesc As String, _
                    ByVal vSubtipus As BoFldSubTypes, _
                    Optional ByVal vValorDefecte As String = "") As Long
        '
        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        Dim bTablaUsuario As Boolean
        '
        If InStr(vTaula, "@") <> 0 Then
            vTaula = Replace(vTaula, "@", "")
            bTablaUsuario = True
        End If
        '
        If Not bTablaUsuario Then
            '
            If UserFieldsExist(vTaula, vCamp) = False Then
                oUserFieldsMD = Me.m_ParentAddon.SBO_Company.GetBusinessObject(oUserFields)
                oUserFieldsMD.TableName = vTaula
                oUserFieldsMD.Name = vCamp
                oUserFieldsMD.Description = Left(vDesc, 30)
                oUserFieldsMD.Type = db_Memo
                oUserFieldsMD.SubType = vSubtipus           ' st_Image , st_Link
                If vValorDefecte <> "" Then
                    oUserFieldsMD.DefaultValue = vValorDefecte
                End If

                lRetCode = oUserFieldsMD.Add
                addCampMemo = lRetCode
                If lRetCode <> 0 Then
                    Me.m_ParentAddon.SBO_Company.GetLastError(lErrCode, sErrMsg)
                    MsgBox(sErrMsg)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
                '
            End If
        Else
            '
            If UserFieldsExistT(vTaula, vCamp) = False Then
                oUserFieldsMD = Me.m_ParentAddon.SBO_Company.GetBusinessObject(oUserFields)
                oUserFieldsMD.TableName = vTaula
                oUserFieldsMD.Name = vCamp
                oUserFieldsMD.Description = Left(vDesc, 30)
                oUserFieldsMD.Type = db_Memo
                oUserFieldsMD.SubType = vSubtipus
                If vValorDefecte <> "" Then
                    oUserFieldsMD.DefaultValue = vValorDefecte
                End If

                lRetCode = oUserFieldsMD.Add
                addCampMemo = lRetCode
                If lRetCode <> 0 Then
                    Me.m_ParentAddon.SBO_Company.GetLastError(lErrCode, sErrMsg)
                    MsgBox(sErrMsg)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
                '
            End If
        End If
        '
    End Function

    Private Function addCampData(ByVal vTaula As String, _
                                ByVal vCamp As String, _
                                ByVal vDesc As String, _
                                Optional ByVal vSubtipus As BoFldSubTypes = st_None, _
                                Optional ByVal vValorDefecte As String = "") As Long
        '
        'Per afegir un camp de data
        'Per exemple:
        '    addCampData "SEIEDI", "SEI_Data", "Data exportacio"
        '
        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        Dim bTablaUsuario As Boolean
        '
        If InStr(vTaula, "@") <> 0 Then
            vTaula = Replace(vTaula, "@", "")
            bTablaUsuario = True
        End If
        '
        If Not bTablaUsuario Then
            '
            If UserFieldsExist(vTaula, vCamp) = False Then
                oUserFieldsMD = Me.m_ParentAddon.SBO_Company.GetBusinessObject(oUserFields)
                oUserFieldsMD.TableName = vTaula
                oUserFieldsMD.Name = vCamp
                oUserFieldsMD.Description = Left(vDesc, 30)
                oUserFieldsMD.Type = db_Date
                oUserFieldsMD.SubType = vSubtipus
                If vValorDefecte <> "" Then
                    oUserFieldsMD.DefaultValue = vValorDefecte
                End If

                lRetCode = oUserFieldsMD.Add
                addCampData = lRetCode
                If lRetCode <> 0 Then
                    Me.m_ParentAddon.SBO_Company.GetLastError(lErrCode, sErrMsg)
                    MsgBox(sErrMsg)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            End If
        Else
            '
            If UserFieldsExistT(vTaula, vCamp) = False Then
                oUserFieldsMD = Me.m_ParentAddon.SBO_Company.GetBusinessObject(oUserFields)
                oUserFieldsMD.TableName = vTaula
                oUserFieldsMD.Name = vCamp
                oUserFieldsMD.Description = Left(vDesc, 30)
                oUserFieldsMD.Type = db_Date
                oUserFieldsMD.SubType = vSubtipus
                If vValorDefecte <> "" Then
                    oUserFieldsMD.DefaultValue = vValorDefecte
                End If

                lRetCode = oUserFieldsMD.Add
                addCampData = lRetCode
                If lRetCode <> 0 Then
                    Me.m_ParentAddon.SBO_Company.GetLastError(lErrCode, sErrMsg)
                    MsgBox(sErrMsg)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            End If
        End If
        '
    End Function

    Private Function addCampHora(ByVal vTaula As String, ByVal vCamp As String, _
                                ByVal vDesc As String, _
                                Optional ByVal vSubtipus As BoFldSubTypes = st_Time, _
                                Optional ByVal vValorDefecte As String = "") As Long
        '
        'Per afegir un camp d'hora
        'Per exemple:
        '    addCampHora "SEIEDI", "SEI_Hora", "Hora exportacio"
        '
        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        Dim bTablaUsuario As Boolean
        '
        If InStr(vTaula, "@") <> 0 Then
            vTaula = Replace(vTaula, "@", "")
            bTablaUsuario = True
        End If
        '
        If Not bTablaUsuario Then
            '
            If UserFieldsExist(vTaula, vCamp) = False Then
                oUserFieldsMD = Me.m_ParentAddon.SBO_Company.GetBusinessObject(oUserFields)
                oUserFieldsMD.TableName = vTaula
                oUserFieldsMD.Name = vCamp
                oUserFieldsMD.Description = Left(vDesc, 30)
                oUserFieldsMD.Type = db_Date
                oUserFieldsMD.SubType = vSubtipus
                If vValorDefecte <> "" Then
                    oUserFieldsMD.DefaultValue = vValorDefecte
                End If

                lRetCode = oUserFieldsMD.Add
                addCampHora = lRetCode
                If lRetCode <> 0 Then
                    Me.m_ParentAddon.SBO_Company.GetLastError(lErrCode, sErrMsg)
                    MsgBox(sErrMsg)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            End If
        Else
            '
            If UserFieldsExistT(vTaula, vCamp) = False Then
                oUserFieldsMD = Me.m_ParentAddon.SBO_Company.GetBusinessObject(oUserFields)
                oUserFieldsMD.TableName = vTaula
                oUserFieldsMD.Name = vCamp
                oUserFieldsMD.Description = Left(vDesc, 30)
                oUserFieldsMD.Type = db_Date
                oUserFieldsMD.SubType = vSubtipus
                If vValorDefecte <> "" Then
                    oUserFieldsMD.DefaultValue = vValorDefecte
                End If

                lRetCode = oUserFieldsMD.Add
                addCampHora = lRetCode
                If lRetCode <> 0 Then
                    Me.m_ParentAddon.SBO_Company.GetLastError(lErrCode, sErrMsg)
                    MsgBox(sErrMsg)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            End If
        End If
        '
    End Function

    Private Function addCampAlpha(ByVal vTaula As String, _
                        ByVal vCamp As String, _
                        ByVal vDesc As String, _
                        ByVal vLong As Long, _
                        Optional ByVal vLlistaValors As String = "", _
                        Optional ByVal vSubtipus As BoFldSubTypes = st_None, _
                        Optional ByVal vValorDefecte As String = "", _
                        Optional ByVal vTablaEnlazada As String = "") As Long
        '
        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        Dim a, b, i

        If UserFieldsExist(vTaula, vCamp) = False Then
            oUserFieldsMD = Me.m_ParentAddon.SBO_Company.GetBusinessObject(oUserFields)
            oUserFieldsMD.TableName = vTaula
            oUserFieldsMD.Name = vCamp
            oUserFieldsMD.Description = Left(vDesc, 30)
            oUserFieldsMD.Type = db_Alpha
            oUserFieldsMD.SubType = vSubtipus
            oUserFieldsMD.EditSize = vLong
            oUserFieldsMD.Size = vLong
            '
            If Len(vLlistaValors) > 0 Then
                a = Explode(vLlistaValors, ";")
                For i = 0 To UBound(a)
                    b = Explode(a(i), "|")
                    oUserFieldsMD.ValidValues.Value = b(0)
                    oUserFieldsMD.ValidValues.Description = b(1)
                    oUserFieldsMD.ValidValues.Add()
                Next i
            End If

            If vValorDefecte <> "" Then
                oUserFieldsMD.DefaultValue = vValorDefecte
            End If

            If vTablaEnlazada <> "" Then
                oUserFieldsMD.LinkedTable = vTablaEnlazada
            End If

            lRetCode = oUserFieldsMD.Add
            addCampAlpha = lRetCode
            If lRetCode <> 0 Then
                Me.m_ParentAddon.SBO_Company.GetLastError(lErrCode, sErrMsg)
                MsgBox("No se ha creado el campo de usuario '" & vCamp & "' a la tabla '" & vTaula & "'. Error: " & sErrMsg)
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)

            '
        End If

    End Function

    Function Explode(ByVal vCadena As String, ByVal vSeparador As String) As String()
        '
        'Retorna un array de valors de cadenes separades
        'exemple: explode("1|Un;2|Dos",";") retorna un array de 2 elements "1|Un" i "2|Dos"
        '
        Dim i As Integer
        Dim a() As String
        Dim fl As Boolean
        ReDim a(0)
        '
        fl = False
        a(0) = ""
        Do
            i = InStr(vCadena, vSeparador)
            If i > 0 Then
                If Not fl Then
                    fl = True
                Else
                    ReDim Preserve a(0 To UBound(a) + 1)
                End If
                a(UBound(a)) = Mid(vCadena, 1, i - 1)
                If Len(vCadena) = Len(vSeparador) Then
                    'L'últim element és buit. l'afegim a l'arrai
                    ReDim Preserve a(0 To UBound(a) + 1)
                    a(UBound(a)) = ""
                End If
                vCadena = Mid(vCadena, i + (Len(vSeparador)))
            ElseIf Len(vCadena) > 0 Then
                If Not fl Then
                    fl = True
                Else
                    ReDim Preserve a(0 To UBound(a) + 1)
                End If
                a(UBound(a)) = vCadena
            End If
        Loop Until i = 0
        Explode = a
    End Function

    Function Implode(ByVal vArray As Array, ByVal vSeparador As String) As String
        '
        'Funcio inversa a explode
        'Retorna una cadena del valor dels arrais separades pel separador
        'exemple: implode(array("1|Un","2|Dos"),";") retorna la cadena "1|Un;2|Dos"
        '
        Dim i As Integer
        Dim a As String
        a = ""
        For i = LBound(vArray) To UBound(vArray)
            a = a & vArray(i) & vSeparador
        Next i
        If Len(a) > 0 Then
            a = Mid(a, 1, Len(a) - Len(vSeparador))
        End If
        Implode = a
    End Function
    '
    Private Function UserTablesExist(ByVal sTableName As String) As Boolean
        Dim oTmpRecordset As SAPbobsCOM.Recordset
        Dim ls As String = ""
        Dim sHANA As String = ""

        oTmpRecordset = Me.m_ParentAddon.SBO_Company.GetBusinessObject(BoRecordset)
        ls = "Select Count(*) From OUTB Where TableName = '" & sTableName & "'"
        sHANA = "Select Count(*) From ""OUTB"" Where ""TableName"" = '" & sTableName & "'"
        oTmpRecordset.DoQuery(CheckIfHana(ls, sHANA))

        If oTmpRecordset.Fields.Item(0).Value > 0 Then
            UserTablesExist = True
        Else
            UserTablesExist = False
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oTmpRecordset)


    End Function

    Private Function UserFieldsExist(ByVal sTableName As String, ByVal sFieldName As String) As Boolean
        Dim oTmpRecordset As SAPbobsCOM.Recordset
        Dim ls As String = ""
        Dim sHANA As String = ""
        oTmpRecordset = Me.m_ParentAddon.SBO_Company.GetBusinessObject(BoRecordset)
        '
        ls = "Select Count(*) From CUFD Where TableId = '" & sTableName & "' And AliasID = '" & sFieldName & "'"
        sHANA = "Select Count(*) From ""CUFD"" Where ""TableID"" = '" & sTableName & "' And ""AliasID"" = '" & sFieldName & "'"
        '
        oTmpRecordset.DoQuery(CheckIfHana(ls, sHANA))

        If oTmpRecordset.Fields.Item(0).Value > 0 Then
            UserFieldsExist = True
        Else
            UserFieldsExist = False
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oTmpRecordset)

    End Function
    '
    Private Function CrearIndexTaula(ByVal NomTaula As String, ByVal NomIndex As String, ByRef Camps As String()) As Long
        Dim oUKeys As SAPbobsCOM.UserKeysMD
        Dim i As Long

        oUKeys = m_ParentAddon.SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)
        oUKeys.TableName = NomTaula
        oUKeys.KeyName = NomIndex

        oUKeys.Elements.ColumnAlias = Camps(0)
        For i = 1 To Camps.Length - 1
            oUKeys.Elements.Add()
            oUKeys.Elements.ColumnAlias = Camps(i)
        Next i

        oUKeys.Unique = SAPbobsCOM.BoYesNoEnum.tYES
        CrearIndexTaula = oUKeys.Add

        If CrearIndexTaula <> 0 Then
            m_ParentAddon.SBO_Company.GetLastError(lErrCode, sErrMsg)
            m_ParentAddon.SBO_Application.SetStatusBarMessage(sErrMsg)
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUKeys)
    End Function
    Private Function UserFieldID(ByVal sTableName As String, ByVal sFieldName As String) As Long
        '
        Dim oTmpRecordset As SAPbobsCOM.Recordset
        Dim ls As String = ""
        Dim sHANA As String = ""
        '
        oTmpRecordset = Me.m_ParentAddon.SBO_Company.GetBusinessObject(BoRecordset)
        '
        ls = "Select FieldID From CUFD Where TableId = '" & sTableName & "' And AliasID = '" & sFieldName & "'"
        sHANA = "Select ""FieldID"" From ""CUFD"" Where ""TableId"" = '" & sTableName & "' And ""AliasID"" = '" & sFieldName & "'"
        oTmpRecordset.DoQuery(CheckIfHana(ls, sHANA))
        '
        UserFieldID = -1
        '
        If Not oTmpRecordset.EoF Then
            UserFieldID = oTmpRecordset.Fields.Item("FieldID").Value
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oTmpRecordset)
        '
    End Function
    '
    Private Function UserKeyID(ByVal sTableName As String, ByVal sKeyName As String) As Long
        '
        Dim oTmpRecordset As SAPbobsCOM.Recordset
        Dim ls As String = ""
        Dim sHANA As String = ""
        '
        oTmpRecordset = Me.m_ParentAddon.SBO_Company.GetBusinessObject(BoRecordset)
        '
        sTableName = "@" & sTableName
        '
        ls = "Select KeyId From OUKD Where TableName = '" & sTableName & "' And KeyName = '" & sKeyName & "'"
        sHANA = "Select ""KeyId"" From ""OUKD"" Where ""TableName"" = '" & sTableName & "' And ""KeyName"" = '" & sKeyName & "'"
        oTmpRecordset.DoQuery(CheckIfHana(ls, sHANA))
        '
        UserKeyID = -1
        '
        If Not oTmpRecordset.EoF Then
            UserKeyID = oTmpRecordset.Fields.Item("KeyId").Value
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oTmpRecordset)
        '
    End Function

    Private Function UserFieldsExistT(ByVal sTableName As String, ByVal sFieldName As String) As Boolean
        Dim oTmpRecordset As SAPbobsCOM.Recordset
        Dim ls As String = ""
        Dim sHANA As String = ""

        oTmpRecordset = Me.m_ParentAddon.SBO_Company.GetBusinessObject(BoRecordset)
        If Mid(sTableName, 1, 1) <> "@" Then
            sTableName = "@" & sTableName
        End If
        ls = "Select Count(*) From CUFD Where TableId = '" & sTableName & "' And AliasID = '" & sFieldName & "'"
        sHANA = "Select Count(*) From ""CUFD"" Where ""TableID"" = '" & sTableName & "' And ""AliasID"" = '" & sFieldName & "'"
        oTmpRecordset.DoQuery(CheckIfHana(ls, sHANA))

        If oTmpRecordset.Fields.Item(0).Value > 0 Then
            UserFieldsExistT = True
        Else
            UserFieldsExistT = False
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oTmpRecordset)


    End Function

#End Region
    '
#Region "Contadores"
    Public Sub CrearContador(ByVal sNombreContador As String)
        '
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim sCode As String
        '
        Try
            If Not ExisteContador(sNombreContador) Then
                '
                oUserTable = Me.m_ParentAddon.SBO_Company.UserTables.Item("SEI_CONTADORES")
                sCode = ""
                '
                ''''sCode = ObtenerCode_IMA_CONTADORS(Me.m_ParentAddon).ToString
                '
                oUserTable.Code = sCode
                oUserTable.Name = sNombreContador
                If oUserTable.Add <> 0 Then
                    Throw New Exception(RecuperarErrorSap(Me.m_ParentAddon))
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable)

            End If

        Catch ex As Exception
            Me.m_ParentAddon.SBO_Application.MessageBox(ex.Message.ToString)
        End Try
    End Sub
    '
    Private Function ExisteContador(ByVal sContador As String, Optional ByRef sCode As String = "")

        Dim sSQL As String
        Dim sHANA As String
        Dim oRcs As SAPbobsCOM.Recordset
        '
        ExisteContador = False
        '
        sSQL = ""
        sSQL = sSQL & "SELECT Code ,Name FROM [@SEI_CONTADORES] WHERE Name='" & sContador & "'"
        '
        sHANA = ""
        sHANA = sHANA & "SELECT ""Code"" ,""Name"" FROM ""@SEI_CONTADORES"" WHERE ""Name""='" & sContador & "'"

        oRcs = Me.m_ParentAddon.SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRcs.DoQuery(CheckIfHana(sSQL, sHANA))
        '
        If Not oRcs.EoF Then
            sCode = oRcs.Fields.Item("Code").Value
            ExisteContador = True
        End If
        '
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRcs)
        '
    End Function

#End Region
    '
End Class
