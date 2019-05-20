Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO

Public Class frmProductRecord1

    Private Sub frmGuestRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetData()
    End Sub

    Private Sub txtCustomers_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProductName.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Product_Name as [Product Name],Shaft_Name as [Shaft Name],Length as [Shaft Length],Stamping_Name as [Stamping Name], Stamping_Od as [Stamping OD], Stamping_Id as [Stamping ID], Stamping_Type as [Stamping Type],Commutator_Name as [Commutator Name], C_Od as [Commutator OD], C_Id as [Commutator ID],Copper_Length as [Copper Length],Fan_Name as [Fan Name], F_Od as [Fan OD], F_Id as [Fan ID],F_Width as [Width],F_Step as [Step],Copper_Wire as [Copper Wire] from Product,Shaft,Stamping,Commutator,Fan,CopperWire,Product_Shaft,Product_Stamping,Product_Commutator,Product_Fan,Product_CopperWire where Shaft.ID=Product_Shaft.ShaftID and Product.ID=Product_Shaft.ProductID and Stamping.ID=Product_Stamping.StampingID and Product.ID=Product_Stamping.ProductID and Commutator.ID=Product_Commutator.CommutatorID and Product.ID=Product_Commutator.ProductID and Fan.ID=Product_Fan.FanID and Product.ID=Product_Fan.ProductID and CopperWire.ID=Product_CopperWire.CopperWireID and Product.ID=Product_CopperWire.ProductID and Product_Name like '" & txtProductName.Text & "%' order by Product_name", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Product")
            dataGridView1.DataSource = myDataSet.Tables("Product").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub GetData()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Product_Name as [Product Name],Shaft_Name as [Shaft Name],Length as [Shaft Length],Stamping_Name as [Stamping Name], Stamping_Od as [Stamping OD], Stamping_Id as [Stamping ID], Stamping_Type as [Stamping Type],Commutator_Name as [Commutator Name], C_Od as [Commutator OD], C_Id as [Commutator ID],Copper_Length as [Copper Length],Fan_Name as [Fan Name], F_Od as [Fan OD], F_Id as [Fan ID],F_Width as [Width],F_Step as [Step],Copper_Wire as [Copper Wire] from Product,Shaft,Stamping,Commutator,Fan,CopperWire,Product_Shaft,Product_Stamping,Product_Commutator,Product_Fan,Product_CopperWire where Shaft.ID=Product_Shaft.ShaftID and Product.ID=Product_Shaft.ProductID and Stamping.ID=Product_Stamping.StampingID and Product.ID=Product_Stamping.ProductID and Commutator.ID=Product_Commutator.CommutatorID and Product.ID=Product_Commutator.ProductID and Fan.ID=Product_Fan.FanID and Product.ID=Product_Fan.ProductID and CopperWire.ID=Product_CopperWire.CopperWireID and Product.ID=Product_CopperWire.ProductID order by Product_name", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Product")
            dataGridView1.DataSource = myDataSet.Tables("Product").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dataGridView1_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles dataGridView1.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If dataGridView1.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            dataGridView1.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        txtProductName.Text = ""
        txtCommutatorName.Text = ""
        txtCopperWire.Text = ""
        txtFanName.Text = ""
        txtShaftName.Text = ""
        txtStampingName.Text = ""
        GetData()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If dataGridView1.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = dataGridView1.RowCount - 1
            colsTotal = dataGridView1.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = dataGridView1.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = dataGridView1.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub frmProductRecord1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub

    Private Sub txtShaftName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtShaftName.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Product_Name as [Product Name],Shaft_Name as [Shaft Name],Length as [Shaft Length],Stamping_Name as [Stamping Name], Stamping_Od as [Stamping OD], Stamping_Id as [Stamping ID], Stamping_Type as [Stamping Type],Commutator_Name as [Commutator Name], C_Od as [Commutator OD], C_Id as [Commutator ID],Copper_Length as [Copper Length],Fan_Name as [Fan Name], F_Od as [Fan OD], F_Id as [Fan ID],F_Width as [Width],F_Step as [Step],Copper_Wire as [Copper Wire] from Product,Shaft,Stamping,Commutator,Fan,CopperWire,Product_Shaft,Product_Stamping,Product_Commutator,Product_Fan,Product_CopperWire where Shaft.ID=Product_Shaft.ShaftID and Product.ID=Product_Shaft.ProductID and Stamping.ID=Product_Stamping.StampingID and Product.ID=Product_Stamping.ProductID and Commutator.ID=Product_Commutator.CommutatorID and Product.ID=Product_Commutator.ProductID and Fan.ID=Product_Fan.FanID and Product.ID=Product_Fan.ProductID and CopperWire.ID=Product_CopperWire.CopperWireID and Product.ID=Product_CopperWire.ProductID and Shaft_Name like '" & txtShaftName.Text & "%' order by Shaft_name", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Product")
            dataGridView1.DataSource = myDataSet.Tables("Product").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtStampingName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStampingName.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Product_Name as [Product Name],Shaft_Name as [Shaft Name],Length as [Shaft Length],Stamping_Name as [Stamping Name], Stamping_Od as [Stamping OD], Stamping_Id as [Stamping ID], Stamping_Type as [Stamping Type],Commutator_Name as [Commutator Name], C_Od as [Commutator OD], C_Id as [Commutator ID],Copper_Length as [Copper Length],Fan_Name as [Fan Name], F_Od as [Fan OD], F_Id as [Fan ID],F_Width as [Width],F_Step as [Step],Copper_Wire as [Copper Wire] from Product,Shaft,Stamping,Commutator,Fan,CopperWire,Product_Shaft,Product_Stamping,Product_Commutator,Product_Fan,Product_CopperWire where Shaft.ID=Product_Shaft.ShaftID and Product.ID=Product_Shaft.ProductID and Stamping.ID=Product_Stamping.StampingID and Product.ID=Product_Stamping.ProductID and Commutator.ID=Product_Commutator.CommutatorID and Product.ID=Product_Commutator.ProductID and Fan.ID=Product_Fan.FanID and Product.ID=Product_Fan.ProductID and CopperWire.ID=Product_CopperWire.CopperWireID and Product.ID=Product_CopperWire.ProductID and Stamping_Name like '" & txtStampingName.Text & "%' order by Stamping_name", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Product")
            dataGridView1.DataSource = myDataSet.Tables("Product").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtCommutatorName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCommutatorName.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Product_Name as [Product Name],Shaft_Name as [Shaft Name],Length as [Shaft Length],Stamping_Name as [Stamping Name], Stamping_Od as [Stamping OD], Stamping_Id as [Stamping ID], Stamping_Type as [Stamping Type],Commutator_Name as [Commutator Name], C_Od as [Commutator OD], C_Id as [Commutator ID],Copper_Length as [Copper Length],Fan_Name as [Fan Name], F_Od as [Fan OD], F_Id as [Fan ID],F_Width as [Width],F_Step as [Step],Copper_Wire as [Copper Wire] from Product,Shaft,Stamping,Commutator,Fan,CopperWire,Product_Shaft,Product_Stamping,Product_Commutator,Product_Fan,Product_CopperWire where Shaft.ID=Product_Shaft.ShaftID and Product.ID=Product_Shaft.ProductID and Stamping.ID=Product_Stamping.StampingID and Product.ID=Product_Stamping.ProductID and Commutator.ID=Product_Commutator.CommutatorID and Product.ID=Product_Commutator.ProductID and Fan.ID=Product_Fan.FanID and Product.ID=Product_Fan.ProductID and CopperWire.ID=Product_CopperWire.CopperWireID and Product.ID=Product_CopperWire.ProductID and Commutator_Name like '" & txtCommutatorName.Text & "%' order by Commutator_name", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Product")
            dataGridView1.DataSource = myDataSet.Tables("Product").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtFanName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFanName.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Product_Name as [Product Name],Shaft_Name as [Shaft Name],Length as [Shaft Length],Stamping_Name as [Stamping Name], Stamping_Od as [Stamping OD], Stamping_Id as [Stamping ID], Stamping_Type as [Stamping Type],Commutator_Name as [Commutator Name], C_Od as [Commutator OD], C_Id as [Commutator ID],Copper_Length as [Copper Length],Fan_Name as [Fan Name], F_Od as [Fan OD], F_Id as [Fan ID],F_Width as [Width],F_Step as [Step],Copper_Wire as [Copper Wire] from Product,Shaft,Stamping,Commutator,Fan,CopperWire,Product_Shaft,Product_Stamping,Product_Commutator,Product_Fan,Product_CopperWire where Shaft.ID=Product_Shaft.ShaftID and Product.ID=Product_Shaft.ProductID and Stamping.ID=Product_Stamping.StampingID and Product.ID=Product_Stamping.ProductID and Commutator.ID=Product_Commutator.CommutatorID and Product.ID=Product_Commutator.ProductID and Fan.ID=Product_Fan.FanID and Product.ID=Product_Fan.ProductID and CopperWire.ID=Product_CopperWire.CopperWireID and Product.ID=Product_CopperWire.ProductID and Fan_Name like '" & txtFanName.Text & "%' order by Fan_name", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Product")
            dataGridView1.DataSource = myDataSet.Tables("Product").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtCopperWire_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCopperWire.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT Product_Name as [Product Name],Shaft_Name as [Shaft Name],Length as [Shaft Length],Stamping_Name as [Stamping Name], Stamping_Od as [Stamping OD], Stamping_Id as [Stamping ID], Stamping_Type as [Stamping Type],Commutator_Name as [Commutator Name], C_Od as [Commutator OD], C_Id as [Commutator ID],Copper_Length as [Copper Length],Fan_Name as [Fan Name], F_Od as [Fan OD], F_Id as [Fan ID],F_Width as [Width],F_Step as [Step],Copper_Wire as [Copper Wire] from Product,Shaft,Stamping,Commutator,Fan,CopperWire,Product_Shaft,Product_Stamping,Product_Commutator,Product_Fan,Product_CopperWire where Shaft.ID=Product_Shaft.ShaftID and Product.ID=Product_Shaft.ProductID and Stamping.ID=Product_Stamping.StampingID and Product.ID=Product_Stamping.ProductID and Commutator.ID=Product_Commutator.CommutatorID and Product.ID=Product_Commutator.ProductID and Fan.ID=Product_Fan.FanID and Product.ID=Product_Fan.ProductID and CopperWire.ID=Product_CopperWire.CopperWireID and Product.ID=Product_CopperWire.ProductID and Copper_Wire like '" & txtCopperWire.Text & "%' order by Copper_Wire", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Product")
            dataGridView1.DataSource = myDataSet.Tables("Product").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class