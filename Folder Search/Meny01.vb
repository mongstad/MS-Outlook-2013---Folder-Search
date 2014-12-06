Imports Microsoft.Office.Tools.Ribbon




Public Class Meny01
    


    Private Sub txtSearch_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles txtSearch.TextChanged
        My.Settings.strSearchString = txtSearch.Text


    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSearch.Click
        Form1.Close()
        Form1.Show()
    End Sub

    
End Class
