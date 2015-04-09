Public Class Program
    <STAThread()> _
    Shared Sub Main()

        ' Declare a variable named frm1 of type Form1.
        Dim frmMain As frmMain
        frmMain = New frmMain()
        frmMain.Visible = False
        Application.Run()

    End Sub
End Class
