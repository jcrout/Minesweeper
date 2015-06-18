Public Class frmMain
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim G As New Minesweeper(Me, 9, 9, 10, 0, 16, 16, Nothing, False)
    End Sub
End Class