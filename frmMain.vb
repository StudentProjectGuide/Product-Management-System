Public Class frmMain

    Private Sub frmMain_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        End
    End Sub

    Private Sub btnRM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRM.Click
        frmRawMaterialCategory.Show()
        Me.Hide()
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        End
    End Sub

    Private Sub btnRegistration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRegistration.Click
        frmRegistration.Show()
        Me.Hide()
    End Sub

    Private Sub btnLoginInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoginInfo.Click
        frmLoginDetails.Show()
        Me.Hide()
    End Sub

    Private Sub btnRMReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRMReport.Click
        frmRawMaterialsRecord.Show()
        Me.Hide()
    End Sub

    Private Sub btnProducts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProducts.Click
        frmProduct.Show()
        Me.Hide()
    End Sub

    Private Sub btnProductsReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProductsReport.Click
        Me.Hide()
        frmProductRecord1.Show()
    End Sub
End Class
