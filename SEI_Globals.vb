'
Option Explicit On
'
Imports System.Text
Imports System.IO

Module SEI_Globals
#Region "Variables"
    Public Const b_Hana As Boolean = False
    Public sBBDD_GESTION As String

#End Region
#Region "Funciones Fichero.INI"
    ' Leer una clave de un fichero INI
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Integer, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    '
    Public Function IniGet(ByVal sFileName As String, ByVal sSection As String, ByVal sKeyName As String, Optional ByVal sDefault As String = "") As String
        '--------------------------------------------------------------------------
        ' Devuelve el valor de una clave de un fichero INI
        ' Los par�metros son:
        '   sFileName   El fichero INI
        '   sSection    La secci�n de la que se quiere leer
        '   sKeyName    Clave
        '   sDefault    Valor opcional que devolver� si no se encuentra la clave
        '--------------------------------------------------------------------------
        ' sSection ->   "Parametros"
        ' sKeyName ->   "U" , "I" , "P"
        '
        ' [Parametros]
        ' U = sa
        ' I = IG
        ' P =seidor.65

        Dim ret As Integer
        Dim sRetVal As String
        '
        sRetVal = New String(Chr(0), 255)
        '
        ret = GetPrivateProfileString(sSection, sKeyName, sDefault, sRetVal, Len(sRetVal), sFileName)
        If ret = 0 Then
            Return sDefault
        Else
            Return Left(sRetVal, ret)
        End If
    End Function

#End Region

#Region "Funciones Tipos de Datos"
    Function NullToDate(ByVal Valor As String) As Date
        If IsNothing(Valor) Then
            NullToDate = "01/01/1900"
        ElseIf IsDBNull(Valor) Or Trim(Valor.ToString) = "" Then
            NullToDate = "01/01/1900"
        Else
            NullToDate = Convert.ToDateTime(Valor.ToString)  ' Pasar a data
        End If
    End Function

    Function NullToDoble(ByRef Valor As String) As Double
        If IsNothing(Valor) Then
            NullToDoble = 0
        ElseIf IsDBNull(Valor) Or Trim(Valor.GetType.ToString) = "" Or Trim(Valor.ToString) = "" Then
            NullToDoble = 0
        Else
            NullToDoble = Convert.ToDouble(Valor.ToString)  ' Pasar a double
        End If
    End Function
    '
    Function NullToLong(ByVal Valor As Object) As Long
        If IsNothing(Valor) Then
            Return 0
        ElseIf IsDBNull(Valor) Or Trim(Valor.GetType.ToString) = "" Then
            Return 0
        Else
            Return CType(Valor, Long)   ' Pasar a Long
        End If
    End Function
    '
    Function NullToInt(ByVal Valor As Object) As Integer
        If IsNothing(Valor) Then
            Return 0
        ElseIf IsDBNull(Valor) Or Trim(Valor.GetType.ToString) = "" Then
            Return 0
        Else
            Return Convert.ToInt32(Valor.ToString)   ' Pasar a integer
        End If
    End Function
    '
    Function NullToText(ByVal Valor As Object) As String
        If IsNothing(Valor) Then
            Return ""
        ElseIf IsDBNull(Valor) Or Valor.GetType.ToString = "" Then
            Return ""
        Else
            Return Valor
        End If
    End Function
    '
    Function CeroToBlancos(ByVal Valor As Object) As String
        If IsDBNull(Valor) Or Trim(Valor.GetType.ToString) = "0" Then
            Return ""
        Else
            Return Valor.ToString
        End If
    End Function

    Function Formato_Decimales_IG(ByVal Valor As Object) As String
        Valor = Valor.ToString.Replace(".", "")
        Valor = Valor.ToString.Replace(",", ".")
        Return Valor.ToString
    End Function
    '
    Function Formato_Decimales_ES(ByVal Valor As Object) As String
        Valor = Valor.ToString.Replace(",", "")
        Valor = Valor.ToString.Replace(".", ",")
        Return Valor.ToString
    End Function

#End Region

#Region "Funciones Funcionalidad Sap"
    Public Function IsSuperUser(ByRef oAddon As SEI_Addon, ByRef lUser As Integer) As SAPbobsCOM.BoYesNoEnum
        '
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim sSQL As String
        Dim sHANA As String
        '
        oRecordset = oAddon.SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '
        sSQL = ""
        sSQL = sSQL & " SELECT INTERNAL_K, SUPERUSER"
        sSQL = sSQL & " FROM  OUSR"
        sSQL = sSQL & " WHERE INTERNAL_K=" & lUser
        '
        sHANA = ""
        sHANA = sHANA & " SELECT ""INTERNAL_K"", ""SUPERUSER"""
        sHANA = sHANA & " FROM  ""OUSR"""
        sHANA = sHANA & " WHERE ""INTERNAL_K""=" & lUser

        oRecordset.DoQuery(CheckIfHana(sSQL, sHANA))
        '
        If Not oRecordset.EoF Then
            '
            If UCase(oRecordset.Fields.Item("SUPERUSER").Value) = "Y" Then
                Return SAPbobsCOM.BoYesNoEnum.tYES
            Else
                Return SAPbobsCOM.BoYesNoEnum.tNO
            End If
            '
        Else
            '
            Return SAPbobsCOM.BoYesNoEnum.tNO
            '
        End If
        '
    End Function

    Public Function NuevoId_UDO(ByRef oAddon As SEI_Addon, ByVal sUDO As String) As Long
        '
        Dim sSQL As String
        Dim sHANA As String
        Dim oRecordset As SAPbobsCOM.Recordset

        Try

            sSQL = "SELECT AutoKey FROM ONNM WHERE ObjectCode = '" & sUDO & "'"
            sHANA = "SELECT ""AutoKey"" FROM ""ONNM"" WHERE ""ObjectCode"" = '" & sUDO & "'"
            oRecordset = oAddon.SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordset.DoQuery(CheckIfHana(sSQL, sHANA))
            If oRecordset.EoF Then
                MsgBox("No se ha encontrado el UDO: '" & sUDO & "'")
                oAddon.SBO_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                End
            Else
                If NullToLong(oRecordset.Fields.Item("AutoKey").Value) = 0 Then
                    NuevoId_UDO = 1
                Else
                    NuevoId_UDO = oRecordset.Fields.Item("AutoKey").Value
                End If
                sSQL = "UPDATE ONNM SET AutoKey = AutoKey + 1 WHERE ObjectCode = '" & sUDO & "'"
                sHANA = "UPDATE ""ONNM"" SET ""AutoKey"" = ""AutoKey"" + 1 WHERE ""ObjectCode"" = '" & sUDO & "'"
                oRecordset = oAddon.SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordset.DoQuery(CheckIfHana(sSQL, sHANA))
            End If

        Catch ExcE As Exception
            If oAddon.SBO_Company.InTransaction Then
                oAddon.SBO_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            oAddon.SBO_Application.MessageBox(ExcE.Message.ToString, 1)
        End Try

    End Function
    '
    Public Function Refresh(ByRef oAddon As SEI_Addon, ByVal stipo As String) As Long

        Dim nI As Integer
        '
        For nI = 0 To oAddon.SBO_Application.Forms.Count - 1
            If oAddon.SBO_Application.Forms.Item(nI).TypeEx = stipo Then
                oAddon.SBO_Application.Forms.Item(nI).Refresh()
                oAddon.SBO_Application.Forms.Item(nI).Update()
                Exit For
            End If
        Next
        '
    End Function
    '
    Public Function DBgetvalue(ByVal oForm As SAPbouiCOM.Form, _
                                ByVal sCamp As String, _
                                ByVal sTaula As String) As String
        '
        Dim oDBDataSource As SAPbouiCOM.DBDataSource
        '
        oDBDataSource = oForm.DataSources.DBDataSources.Item(sTaula)
        '
        Return Trim(oDBDataSource.GetValue(sCamp, 0))
        '
    End Function
    '
    Public Function RecuperarErrorSap(ByRef oAddon As SEI_Addon) As String
        Dim sError As String
        '
        oAddon.ErrCode = 0
        oAddon.ErrMsg = ""
        oAddon.SBO_Company.GetLastError(oAddon.ErrCode, oAddon.ErrMsg)
        sError = "Error: " & oAddon.ErrCode & " " & oAddon.ErrMsg
        '
        Return sError
        '
    End Function

#End Region

#Region "Funciones Varias"

    Public Sub Escriure_Ficher_TXT(ByVal sRutaFitxer As String, ByVal sTexteLinia As String)
        Dim swFitxer As StreamWriter

        swFitxer = New StreamWriter(sRutaFitxer, True, Encoding.GetEncoding(1252)) ''Encoding.ASCII
        ''swFitxer = File.AppendText(sRutaFitxer)
        swFitxer.WriteLine(sTexteLinia)

        swFitxer.Flush()
        swFitxer.Close()

    End Sub

    Public Function RecuperarValores(ByRef SBO_Company As SAPbobsCOM.Company, _
                                     ByVal ls_Camp As String, _
                                     ByVal ls_taula As String, _
                                     ByVal Array_Claus() As String, _
                                     ByVal Array_Valors() As String, _
                                     Optional ByVal Filtro As String = "") As String
        '
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim ls_Sql As String
        Dim lb_buit As Boolean
        Dim li_NumCamps As Integer
        Dim li_i As Integer
        Dim Claus() As String
        Dim Valors() As String
        Dim sCamps As String
        '
        sCamps = ls_Camp
        '
        li_NumCamps = UBound(Array_Claus)
        ReDim Claus(li_NumCamps)
        For li_i = 0 To li_NumCamps
            Claus(li_i) = Array_Claus(li_i)
        Next
        '
        li_NumCamps = UBound(Array_Valors)
        ReDim Valors(li_NumCamps)
        For li_i = 0 To li_NumCamps
            Valors(li_i) = Array_Valors(li_i)
        Next
        '
        ls_Sql = "SELECT " & sCamps & " FROM " + ls_taula
        ls_Sql = ls_Sql + " WHERE "
        For li_i = 0 To li_NumCamps
            ls_Sql = ls_Sql + Claus(li_i) + " = '" + NullToText(Valors(li_i)) + "'"
            If li_i <> li_NumCamps Then ls_Sql = ls_Sql + " AND "
        Next
        '
        oRecordSet = SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '
        ls_Sql = ls_Sql & Filtro
        '
        oRecordSet.DoQuery(ls_Sql)
        '
        '---- Comprovar Si la consulta t� registres -----------------------
        '
        lb_buit = oRecordSet.BoF And oRecordSet.EoF
        '
        'DoEvents
        '
        If Not (lb_buit) Then
            RecuperarValores = NullToText(oRecordSet.Fields.Item(0).Value)
        Else
            RecuperarValores = ""
        End If
        '
    End Function
    '
    Private Function EliminarAlias(ByVal sValor As String)
        '
        Dim sCampo As String
        Dim lPunto As Long
        '
        EliminarAlias = sValor
        '
        lPunto = InStr(sValor, ".")
        If lPunto <> 0 Then
            sCampo = Mid(sValor, lPunto + 1, Len(sValor) - lPunto)
            EliminarAlias = sCampo
        Else
            EliminarAlias = sValor
        End If
        '
    End Function

    Public Sub LiberarObjCOM(ByRef oObjCOM As Object)

        'Liberar y destruir Objecto com
        If Not IsNothing(oObjCOM) Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oObjCOM)
            oObjCOM = Nothing
            GC.Collect()
        End If

    End Sub

    Public Function Camp_XML(ByVal sXML As String, ByVal sCamp As String) As String
        '
        Dim iIni As Integer
        Dim iFin As Integer
        '<?xml version="1.0" encoding="UTF-16" ?><DocumentParams><DocEntry>7</DocEntry></DocumentParams>
        '
        'iIni = InStr(1, sXML, "DocEntry>")
        'iFin = InStr(iIni + 1, sXML, "</DocEntry")
        iIni = InStr(1, sXML, sCamp & ">")
        iFin = InStr(iIni + 1, sXML, "</" & sCamp)
        If iIni <> 0 And iFin <> 0 Then
            Camp_XML = NullToText(Mid(sXML, iIni + Len(sCamp) + 1, iFin - iIni - Len(sCamp) - 1))
        Else
            Camp_XML = 0
        End If
        '
    End Function

    Public Function Arrodonir(ByVal value As Decimal, ByVal numdecimals As Integer)
        Dim xPotencia10 = Math.Pow(10, numdecimals)
        value = value * xPotencia10
        Dim x As Decimal
        ' agafa el valor decimal 
        Dim y As Decimal = value - Math.Floor(value)

        If y >= 0.5 Then
            x = Math.Ceiling(value)
        Else
            x = Math.Floor(value)
        End If
        x = x / xPotencia10
        Return x
    End Function


    Public Function GenerarAlerta(ByRef SBO_Company As SAPbobsCOM.Company, _
                           ByRef sError As String, _
                           ByVal sText As String, ByVal sAssumte As String, _
                           Optional ByVal sUsuariAlerta_3 As String = "", _
                           Optional ByVal sUsuariAlerta_4 As String = "", _
                           Optional ByVal sLinkCol As String = "", _
                           Optional ByVal sLinkTxt As String = "", _
                           Optional ByVal sLinkObj As String = "", _
                           Optional ByVal sLinkVal As String = "") As Boolean
        Dim oMessage As SAPbobsCOM.Messages
        Dim sUsuariAlerta_1 As String
        Dim sUsuariAlerta_2 As String
        Dim sSQL As String
        Dim sHANA As String
        Dim oRS As SAPbobsCOM.Recordset = SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try

            If sUsuariAlerta_3 = "" Then
                sSQL = "SELECT U_SEIUs1AU, U_SEIUs2AU FROM OADM"
                sHANA = "SELECT ""U_SEIUs1AU"", ""U_SEIUs2AU"" FROM OADM"
                oRS.DoQuery(CheckIfHana(sSQL, sHANA))
                sUsuariAlerta_1 = oRS.Fields.Item("U_SEIUs1AU").Value
                sUsuariAlerta_2 = oRS.Fields.Item("U_SEIUs2AU").Value
                'sUsuariAlerta_1 = RecuperarValores(SBO_Company, "U_SEIUs1AU", "OADM", "'1'".Split, ("1").Split)
                'sUsuariAlerta_2 = RecuperarValores(SBO_Company, "U_SEIUs2AU", "OADM", "'1'".Split, ("1").Split)
            Else
                sUsuariAlerta_1 = sUsuariAlerta_3
                sUsuariAlerta_2 = sUsuariAlerta_4
            End If

            If sUsuariAlerta_2 = sUsuariAlerta_1 Then
                sUsuariAlerta_2 = ""
            End If

            oMessage = SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages)

            With oMessage
                If sLinkCol <> "" Then
                    ''.AddDataColumn("FullCost", "FULL DE COST", SAPbobsCOM.BoObjectTypes.oOrders, 7)
                    .AddDataColumn(sLinkCol, sLinkTxt, sLinkObj, sLinkVal)
                End If
                .MessageText = sText
                .Subject = sAssumte
                .Priority = SAPbobsCOM.BoMsgPriorities.pr_High
                With .Recipients
                    .SetCurrentLine(0)
                    .UserCode = sUsuariAlerta_1
                    .SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
                    .SendEmail = SAPbobsCOM.BoYesNoEnum.tNO
                    .SendFax = SAPbobsCOM.BoYesNoEnum.tNO
                    .SendSMS = SAPbobsCOM.BoYesNoEnum.tNO
                    .UserType = SAPbobsCOM.BoMsgRcpTypes.rt_InternalUser
                End With

                If sUsuariAlerta_2 <> "" Then
                    With .Recipients
                        .Add()
                        .SetCurrentLine(1)
                        .UserCode = sUsuariAlerta_2
                        .SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
                        .SendEmail = SAPbobsCOM.BoYesNoEnum.tNO
                        .SendFax = SAPbobsCOM.BoYesNoEnum.tNO
                        .SendSMS = SAPbobsCOM.BoYesNoEnum.tNO
                        .UserType = SAPbobsCOM.BoMsgRcpTypes.rt_InternalUser
                    End With
                End If

                If .Add <> 0 Then
                    GenerarAlerta = False
                    sError = SBO_Company.GetLastErrorDescription
                Else
                    GenerarAlerta = True
                    sError = SBO_Company.GetNewObjectKey
                End If
            End With

        Catch ex As Exception
            GenerarAlerta = False
            If SBO_Company.InTransaction Then
                SBO_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            sError = ex.Message
        Finally
            LiberarObjCOM(oMessage)
        End Try

    End Function

    Public Function ValidarNIE(ByVal sNIE As String) As Boolean
        Dim sNif As String
        sNif = ""

        Select Case Left(sNIE, 1)
            Case "X"
                sNif = "0" & Mid(sNIE, 2)

            Case "Y"
                sNif = "1" & Mid(sNIE, 2)

            Case "Z"
                sNif = "2" & Mid(sNIE, 2)

            Case Else
                ValidarNIE = False
                Exit Function
        End Select

        ValidarNIE = ValidarNIF(sNif)

    End Function

    Public Function ValidarNIF(ByVal sNIF As String) As Boolean

        Const c_LletresNIF As String = "TRWAGMYFPDXBNJZSQVHLCKE"

        Dim sNumNIF As String
        Dim sLletra As String
        Dim dResta As Double
        '
        'DNI/23, s'agafa la resta i se li suma 1.
        'Es busca el resultat a la taula c_LletresNIF
        '
        ValidarNIF = True
        If sNIF <> "" Then
            If IsNumeric(Mid(sNIF, 1, 1)) Then
                If (Len(sNIF) = 9) Then
                    sNumNIF = Left(sNIF, 8)
                    sLletra = Right(sNIF, 1)
                    If Not IsNumeric(sNumNIF) Then
                        ValidarNIF = False
                    Else
                        dResta = Val(sNumNIF)
                        dResta = dResta Mod 23
                        dResta = dResta + 1
                        If Mid(c_LletresNIF, dResta, 1) <> UCase(sLletra) Then
                            ValidarNIF = False
                        End If
                    End If
                Else
                    ValidarNIF = False
                End If
            Else
                ValidarNIF = False
            End If
        End If

    End Function

    Public Function ValidarCIF(ByVal sCIF As String) As Boolean
        Dim strLetra As String, strNumero As String, strDigit As String
        Dim strDigitAux As String
        Dim auxNum As Integer
        Dim i As Integer
        Dim Suma As Integer
        Dim letras As String

        If sCIF = "" Then
            'No valido
            ValidarCIF = True
            Exit Function
        End If

        letras = "ABCDEFGHJKLMNPQRSUVW"
        sCIF = UCase(sCIF)

        If Len(sCIF) < 9 Or Not IsNumeric(Mid(sCIF, 2, 7)) Then
            ValidarCIF = False
            Exit Function
        End If

        strLetra = Mid(sCIF, 1, 1)     'letra del CIF
        strNumero = Mid(sCIF, 2, 7)    'Codigo de Control
        strDigit = Mid(sCIF, 9)        'CIF menos primera y ultima posiciones

        If InStr(letras, strLetra) = 0 Then 'comprobamos la letra del CIF (1� posicion)
            ValidarCIF = False
            Exit Function
        End If

        For i = 1 To 7

            If i Mod 2 = 0 Then
                Suma = Suma + CInt(Mid(strNumero, i, 1))
            Else
                auxNum = CInt(Mid(strNumero, i, 1)) * 2
                Suma = Suma + (auxNum \ 10) + (auxNum Mod 10)
            End If

        Next

        Suma = (10 - (Suma Mod 10)) Mod 10

        Select Case strLetra
            Case "K", "P", "Q", "S", "R", "W", "L", "M", "N", "C"
                If Suma = 0 Then Suma = 10
                Suma = Suma + 64
                strDigitAux = Chr(Suma)

            Case Else
                strDigitAux = CStr(Suma)

        End Select

        If strDigit = strDigitAux Then
            ValidarCIF = True
        Else
            ValidarCIF = False
        End If

    End Function

    Function Validar_DigitControlBanc(ByVal Bank As Integer, ByVal SubBank As Integer, ByVal Account As Double) As String
        Dim sBank As String
        Dim sSubBank As String
        Dim sAccount As String
        Dim Temporal As Integer

        sBank = Format(Bank, "0000")
        sSubBank = Format(SubBank, "0000")
        sAccount = Format(Account, "0000000000")

        Temporal = 0
        Temporal = Temporal + Mid(sBank, 1, 1) * 4
        Temporal = Temporal + Mid(sBank, 2, 1) * 8
        Temporal = Temporal + Mid(sBank, 3, 1) * 5
        Temporal = Temporal + Mid(sBank, 4, 1) * 10
        Temporal = Temporal + Mid(sSubBank, 1, 1) * 9
        Temporal = Temporal + Mid(sSubBank, 2, 1) * 7
        Temporal = Temporal + Mid(sSubBank, 3, 1) * 3
        Temporal = Temporal + Mid(sSubBank, 4, 1) * 6
        Temporal = 11 - (Temporal Mod 11)
        If Temporal = 11 Then
            Validar_DigitControlBanc = "0"
        ElseIf Temporal = 10 Then
            Validar_DigitControlBanc = "1"
        Else
            Validar_DigitControlBanc = Format(Temporal, "0")
        End If

        Temporal = 0
        Temporal = Temporal + Mid(sAccount, 1, 1) * 1
        Temporal = Temporal + Mid(sAccount, 2, 1) * 2
        Temporal = Temporal + Mid(sAccount, 3, 1) * 4
        Temporal = Temporal + Mid(sAccount, 4, 1) * 8
        Temporal = Temporal + Mid(sAccount, 5, 1) * 5
        Temporal = Temporal + Mid(sAccount, 6, 1) * 10
        Temporal = Temporal + Mid(sAccount, 7, 1) * 9
        Temporal = Temporal + Mid(sAccount, 8, 1) * 7
        Temporal = Temporal + Mid(sAccount, 9, 1) * 3
        Temporal = Temporal + Mid(sAccount, 10, 1) * 6
        Temporal = 11 - (Temporal Mod 11)
        If Temporal = 11 Then
            Validar_DigitControlBanc = Validar_DigitControlBanc + "0"
        ElseIf Temporal = 10 Then
            Validar_DigitControlBanc = Validar_DigitControlBanc + "1"
        Else
            Validar_DigitControlBanc = Validar_DigitControlBanc + Format(Temporal, "0")
        End If

    End Function
    '
    Public Function CheckIfHana(ByVal sSQLQuery As String, ByVal sHanaQuery As String) As String
        '
        If b_Hana = True Then
            Return sHanaQuery
        End If
        Return sSQLQuery
        '
    End Function
#End Region
    '
End Module
