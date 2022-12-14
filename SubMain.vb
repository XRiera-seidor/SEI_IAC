Option Explicit On
Option Strict On
'
Module SubMain
    Private m_SBOAddon As SEI_Addon


    '<MTAThread()> Public Sub Main()
    '

    '  STAThreadAttribute
    Public Sub Main() 'STAThreadAttribute
        Try
            Dim bRunApplication As Boolean
            m_SBOAddon = New SEI_SBOAddon("SEI_IAC", bRunApplication)

            ' Starting the Application
            If bRunApplication = True Then
                '
                System.Windows.Forms.Application.Run()

            Else
                System.Windows.Forms.Application.Exit()
            End If


        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try

    End Sub

    Public Function GetM_SBOAddon() As SEI_Addon

        GetM_SBOAddon = m_SBOAddon

    End Function
End Module
