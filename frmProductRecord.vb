Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO

Public Class frmProductRecord

    Private Sub frmGuestRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetData()
    End Sub

    Private Sub txtCustomers_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProductName.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ID as [Product ID], Product_Name as [Product Name],Image1,Image2 from Product where Product_Name like '" & txtProductName.Text & "%' order by Product_name", con)
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
            cmd = New OleDbCommand("SELECT ID as [Product ID], Product_Name as [Product Name],Image1,Image2 from Product order by Product_name", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Product")
            dataGridView1.DataSource = myDataSet.Tables("Product").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = dataGridView1.SelectedRows(0)
            Me.Hide()
            frmProduct.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmProduct.txtProductID.Text = dr.Cells(0).Value.ToString()
            frmProduct.txtProductName.Text = dr.Cells(1).Value.ToString()
            Dim data As Byte() = DirectCast(dr.Cells(2).Value, Byte())
            Dim ms As New MemoryStream(data)
            frmProduct.PictureBox1.Image = Image.FromStream(ms)
            Dim data1 As Byte() = DirectCast(dr.Cells(3).Value, Byte())
            Dim ms1 As New MemoryStream(data1)
            frmProduct.PictureBox2.Image = Image.FromStream(ms1)
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT ShaftID,Shaft_Name,Length from Shaft,Product,Product_Shaft where Shaft.ID=Product_Shaft.ShaftID and Product.ID=Product_Shaft.ProductID and ProductID=" & dr.Cells(0).Value & "", con)
            rdr = cmd.ExecuteReader()
            While rdr.Read()

                Dim lst As New ListViewItem()
                lst.SubItems.Add(rdr(0).ToString().Trim())
                lst.SubItems.Add(rdr(1).ToString().Trim())
                lst.SubItems.Add(rdr(2).ToString().Trim())
                frmProduct.ListView1.Items.Add(lst)
            End While
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT StampingID,Stamping_Name, Stamping_Od, Stamping_Id, Stamping_Type from Stamping,Product,Product_Stamping where Stamping.ID=Product_Stamping.StampingID and Product.ID=Product_Stamping.ProductID and ProductID=" & dr.Cells(0).Value & "", con)
            rdr = cmd.ExecuteReader()
            While rdr.Read()

                Dim lst As New ListViewItem()
                lst.SubItems.Add(rdr(0).ToString().Trim())
                lst.SubItems.Add(rdr(1).ToString().Trim())
                lst.SubItems.Add(rdr(2).ToString().Trim())
                lst.SubItems.Add(rdr(3).ToString().Trim())
                lst.SubItems.Add(rdr(4).ToString().Trim())
                frmProduct.ListView2.Items.Add(lst)
            End While
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT CommutatorID,Commutator_Name, C_Od, C_Id,Copper_Length from Commutator,Product,Product_Commutator where Commutator.ID=Product_Commutator.CommutatorID and Product.ID=Product_Commutator.ProductID and ProductID=" & dr.Cells(0).Value & "", con)
            rdr = cmd.ExecuteReader()
            While rdr.Read()

                Dim lst As New ListViewItem()
                lst.SubItems.Add(rdr(0).ToString().Trim())
                lst.SubItems.Add(rdr(1).ToString().Trim())
                lst.SubItems.Add(rdr(2).ToString().Trim())
                lst.SubItems.Add(rdr(3).ToString().Trim())
                lst.SubItems.Add(rdr(4).ToString().Trim())
                frmProduct.ListView3.Items.Add(lst)
            End While
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT FanID,Fan_Name, F_Od, F_Id,F_Width,F_Step from Fan,Product,Product_Fan where Fan.ID=Product_Fan.FanID and Product.ID=Product_Fan.ProductID and ProductID=" & dr.Cells(0).Value & "", con)
            rdr = cmd.ExecuteReader()
            While rdr.Read()

                Dim lst As New ListViewItem()
                lst.SubItems.Add(rdr(0).ToString().Trim())
                lst.SubItems.Add(rdr(1).ToString().Trim())
                lst.SubItems.Add(rdr(2).ToString().Trim())
                lst.SubItems.Add(rdr(3).ToString().Trim())
                lst.SubItems.Add(rdr(4).ToString().Trim())
                lst.SubItems.Add(rdr(5).ToString().Trim())
                frmProduct.ListView4.Items.Add(lst)
            End While
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT CopperWireID,Copper_Wire from CopperWire,Product,Product_CopperWire where CopperWire.ID=Product_CopperWire.CopperWireID and Product.ID=Product_CopperWire.ProductID and ProductID=" & dr.Cells(0).Value & "", con)
            rdr = cmd.ExecuteReader()
            While rdr.Read()

                Dim lst As New ListViewItem()
                lst.SubItems.Add(rdr(0).ToString().Trim())
                lst.SubItems.Add(rdr(1).ToString().Trim())
                frmProduct.ListView5.Items.Add(lst)
            End While
            frmProduct.btnUpdate.Enabled = True
            frmProduct.btnDelete.Enabled = True
            frmProduct.btnSave.Enabled = False
            frmProduct.btnPrint.Enabled = True
            frmProduct.txtProductName.Focus()
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
        GetData()
    End Sub

End Class