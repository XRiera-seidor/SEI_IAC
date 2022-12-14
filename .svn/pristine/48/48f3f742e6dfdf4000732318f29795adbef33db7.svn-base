Imports System.Diagnostics

Public Class SEI_WindowsSbo
    Implements System.Windows.Forms.IWin32Window

    Private _hwnd As IntPtr

    Public Sub New(ByVal handle As IntPtr)
        Me._hwnd = handle
    End Sub

    'Public Sub New(ByRef oSEI_AddOn As SEI_Addon)
    '    Dim MyProcs() As Process
    '    Dim i As Integer
    '    MyProcs = Process.GetProcessesByName("SAP Business One")
    '    i = 0   '    i = oSEI_AddOn.SBO_Application.AppId  '--> Faltaria solucionar si hi ha més d'un SBO funcionant
    '    Me._hwnd = MyProcs(i).MainWindowHandle
    'End Sub

    Public Sub New(ByRef oSEI_AddOn As SEI_Addon)
        Dim MyProcs() As Process
        Dim i As Integer

        Dim currentProcess As Process
        currentProcess = Process.GetCurrentProcess()

        MyProcs = Process.GetProcessesByName("SAP Business One")
        i = 0   '    i = oSEI_AddOn.SBO_Application.AppId  '--> Faltaria solucionar si hi ha més d'un SBO funcionant (Solucionat)
        If System.Windows.Forms.SystemInformation.TerminalServerSession Then
            '-> Si s'està executant amb Terminal server, busquem quina sessió estem per determinar quin és el SAP de la Sessió
            For i = 0 To MyProcs.Length
                If MyProcs(i).SessionId = currentProcess.SessionId Then
                    Exit For
                End If
            Next
        End If
        Me._hwnd = MyProcs(i).MainWindowHandle

    End Sub

    Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
        Get
            Return _hwnd
        End Get
    End Property
End Class
