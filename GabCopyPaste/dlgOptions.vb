Imports System.Windows.Forms
Imports Microsoft.Win32

Public Class dlgOptions

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        Dim hStartKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run", True)
        If hStartKey IsNot Nothing Then

            Try

                If checkBoxStartup.Checked Then

                    Dim strPath As String = Application.ExecutablePath
                    strPath = """" + strPath + """"

                    hStartKey.SetValue("GabCopyPaste", strPath)

                Else

                    hStartKey.DeleteValue("GabCopyPaste")

                End If
            Catch

                'Catching exceptions is for communists
            End Try

            hStartKey.Close()
        End If

        My.Settings.LoadOnStartup = checkBoxStartup.Checked
        My.Settings.Save()

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        My.Settings.Reload()

        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
