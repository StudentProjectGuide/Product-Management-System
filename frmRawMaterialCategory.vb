Imports System.Data.OleDb
Imports System.IO

Public Class frmRawMaterialCategory

    Private Sub frmRawMaterialCategory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetData()
        GetData1()
        GetData2()
        GetData3()
        GetData4()
    End Sub
    Sub FullReset()
        'Reset()
        'Reset1()
        'Reset2()
        'Reset3()
        Reset4()
    End Sub
    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        Try
            With OpenFileDialog1
                .Filter = ("Images |*.png; *.bmp; *.jpg;*.jpeg; *.gif;")
                .FilterIndex = 4
            End With
            'Clear the file name
            OpenFileDialog1.FileName = ""
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                PictureBox1.Image = Image.FromFile(OpenFileDialog1.FileName)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub btnBrowse1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse1.Click
        Try
            With OpenFileDialog1
                .Filter = ("Images |*.png; *.bmp; *.jpg;*.jpeg; *.gif;")
                .FilterIndex = 4
            End With
            'Clear the file name
            OpenFileDialog1.FileName = ""
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                PictureBox2.Image = Image.FromFile(OpenFileDialog1.FileName)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse2.Click
        Try
            With OpenFileDialog1
                .Filter = ("Images |*.png; *.bmp; *.jpg;*.jpeg; *.gif;")
                .FilterIndex = 4
            End With
            'Clear the file name
            OpenFileDialog1.FileName = ""
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                PictureBox3.Image = Image.FromFile(OpenFileDialog1.FileName)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse3.Click
        Try
            With OpenFileDialog1
                .Filter = ("Images |*.png; *.bmp; *.jpg;*.jpeg; *.gif;")
                .FilterIndex = 4
            End With
            'Clear the file name
            OpenFileDialog1.FileName = ""
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                PictureBox4.Image = Image.FromFile(OpenFileDialog1.FileName)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtShaftName.Text)) = 0 Then
            MessageBox.Show("Please enter shaft name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtShaftName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtShaftLength.Text)) = 0 Then
            MessageBox.Show("Please enter shaft length", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtShaftLength.Focus()
            Exit Sub
        End If
        If PictureBox1.Image Is Nothing Then
            MessageBox.Show("Please browse and select image", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            btnBrowse.Focus()
            Exit Sub
        End If
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select Shaft_name from Shaft where Shaft_name=@d1"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtShaftName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Shaft Name Already Exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtShaftName.Text = ""
                txtShaftName.Focus()
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cd As String = "insert Into Shaft(Shaft_Name,Length,S_Image) VALUES (@d1,@d2,@d3)"
            cmd = New OleDbCommand(cd)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtShaftName.Text)
            cmd.Parameters.AddWithValue("@d2", txtShaftLength.Text)
            Dim ms As New MemoryStream()
            Dim bmpImage As New Bitmap(PictureBox1.Image)
            bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim p As New OleDbParameter("@d3", OleDbType.VarBinary)
            p.Value = data
            cmd.Parameters.Add(p)
            cmd.ExecuteNonQuery()
            con.Close()
            btnSave.Enabled = False
            GetData()
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Reset()
        txtShaftName.Text = ""
        txtShaftLength.Text = ""
        PictureBox1.Image = Nothing
        btnSave.Enabled = True
        btnUpdate.Enabled = False
        btnDelete.Enabled = False
        txtSearchByShaftName.Text = ""
        GetData()
        txtShaftName.Focus()
    End Sub
    Sub Reset1()
        txtStampName.Text = ""
        txtStampOD.Text = ""
        txtStampID.Text = ""
        cmbStampType.SelectedIndex = -1
        PictureBox2.Image = Nothing
        BtnSave1.Enabled = True
        btnUpdate1.Enabled = False
        BtnDelete1.Enabled = False
        txtSearchByStampName.Text = ""
        txtStampName.Focus()
        GetData1()
    End Sub
    Sub Reset2()
        txtCommutatorName.Text = ""
        txtCommutatorOD.Text = ""
        txtCommutatorID.Text = ""
        txtCommutatorCopperLength.Text = ""
        PictureBox3.Image = Nothing
        btnSave2.Enabled = True
        btnUpdate2.Enabled = False
        btnDelete2.Enabled = False
        txtSearchByCommutatorName.Text = ""
        txtCommutatorName.Focus()
        GetData2()
    End Sub
    Sub Reset3()
        txtFanName.Text = ""
        txtFanOD.Text = ""
        txtFanID.Text = ""
        txtFanWidth.Text = ""
        cmbFanStep.SelectedIndex = -1
        PictureBox4.Image = Nothing
        btnSave3.Enabled = True
        btnUpdate3.Enabled = False
        btnDelete3.Enabled = False
        txtSearchByFanName.Text = ""
        txtFanName.Focus()
        GetData3()
    End Sub
    Sub Reset4()
        txtCopperWireName.Text = ""
        btnSave4.Enabled = True
        btnUpdate4.Enabled = False
        btnDelete4.Enabled = False
        txtCopperWireName.Focus()
        txtSearchByCopperWireName.Text = ""
        GetData4()
    End Sub
    Sub GetData()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT * from Shaft Order by Shaft_Name", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView1.Rows.Clear()
            While (rdr.Read() = True)
                DataGridView1.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub GetData1()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT * from Stamping order by Stamping_Name", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView2.Rows.Clear()
            While (rdr.Read() = True)
                DataGridView2.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub GetData2()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT * from Commutator order by Commutator_Name", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView3.Rows.Clear()
            While (rdr.Read() = True)
                DataGridView3.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub GetData3()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT * from Fan order by Fan_Name", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView4.Rows.Clear()
            While (rdr.Read() = True)
                DataGridView4.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub GetData4()
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT * from CopperWire order by Copper_Wire", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView5.Rows.Clear()
            While (rdr.Read() = True)
                DataGridView5.Rows.Add(rdr(0), rdr(1))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Reset()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub DeleteRecord()
        Try
            Dim RowsAffected As Integer = 0
            con = New OleDbConnection(cs)
            con.Open()
            Dim cq1 As String = "delete from Shaft where ID=" & txtID1.Text & ""
            cmd = New OleDbCommand(cq1)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
                GetData()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub DeleteRecord1()
        Try
            Dim RowsAffected As Integer = 0
            con = New OleDbConnection(cs)
            con.Open()
            Dim cq1 As String = "delete from Stamping where ID=" & txtID2.Text & ""
            cmd = New OleDbCommand(cq1)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset1()
                GetData1()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset1()
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub DeleteRecord2()
        Try
            Dim RowsAffected As Integer = 0
            con = New OleDbConnection(cs)
            con.Open()
            Dim cq1 As String = "delete from Commutator where ID=" & txtID3.Text & ""
            cmd = New OleDbCommand(cq1)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset2()
                GetData2()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset2()
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub DeleteRecord3()
        Try
            Dim RowsAffected As Integer = 0
            con = New OleDbConnection(cs)
            con.Open()
            Dim cq1 As String = "delete from Fan where ID=" & txtID4.Text & ""
            cmd = New OleDbCommand(cq1)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset3()
                GetData3()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset3()
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub DeleteRecord4()
        Try
            Dim RowsAffected As Integer = 0
            con = New OleDbConnection(cs)
            con.Open()
            Dim cq1 As String = "delete from CopperWire where ID=" & txtID5.Text & ""
            cmd = New OleDbCommand(cq1)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset4()
                GetData4()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset4()
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            txtID1.Text = dr.Cells(0).Value.ToString()
            txtShaftName.Text = dr.Cells(1).Value.ToString()
            txtShaftLength.Text = dr.Cells(2).Value.ToString()
            Dim data As Byte() = DirectCast(dr.Cells(3).Value, Byte())
            Dim ms As New MemoryStream(data)
            PictureBox1.Image = Image.FromStream(ms)
            btnUpdate.Enabled = True
            btnDelete.Enabled = True
            btnSave.Enabled = False
            txtShaftName.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView1.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView1.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub DataGridView2_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView2.SelectedRows(0)
            txtID2.Text = dr.Cells(0).Value.ToString()
            txtStampName.Text = dr.Cells(1).Value.ToString()
            txtStampOD.Text = dr.Cells(2).Value.ToString()
            txtStampID.Text = dr.Cells(3).Value.ToString()
            cmbStampType.Text = dr.Cells(4).Value.ToString()
            Dim data As Byte() = DirectCast(dr.Cells(5).Value, Byte())
            Dim ms As New MemoryStream(data)
            PictureBox2.Image = Image.FromStream(ms)
            btnUpdate1.Enabled = True
            BtnDelete1.Enabled = True
            BtnSave1.Enabled = False
            txtStampName.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView2_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView2.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView2.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView2.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub DataGridView3_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView3.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView3.SelectedRows(0)
            txtID3.Text = dr.Cells(0).Value.ToString()
            txtCommutatorName.Text = dr.Cells(1).Value.ToString()
            txtCommutatorOD.Text = dr.Cells(2).Value.ToString()
            txtCommutatorID.Text = dr.Cells(3).Value.ToString()
            txtCommutatorCopperLength.Text = dr.Cells(4).Value.ToString()
            Dim data As Byte() = DirectCast(dr.Cells(5).Value, Byte())
            Dim ms As New MemoryStream(data)
            PictureBox3.Image = Image.FromStream(ms)
            btnUpdate2.Enabled = True
            btnDelete2.Enabled = True
            btnSave2.Enabled = False
            txtCommutatorName.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView3_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView3.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView3.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView3.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub DataGridView4_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView4.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView4.SelectedRows(0)
            txtID4.Text = dr.Cells(0).Value.ToString()
            txtFanName.Text = dr.Cells(1).Value.ToString()
            txtFanOD.Text = dr.Cells(2).Value.ToString()
            txtFanID.Text = dr.Cells(3).Value.ToString()
            txtFanWidth.Text = dr.Cells(4).Value.ToString()
            cmbFanStep.Text = dr.Cells(5).Value.ToString()
            Dim data As Byte() = DirectCast(dr.Cells(6).Value, Byte())
            Dim ms As New MemoryStream(data)
            PictureBox4.Image = Image.FromStream(ms)
            btnUpdate3.Enabled = True
            btnDelete3.Enabled = True
            btnSave3.Enabled = False
            txtFanName.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView4_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView4.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView4.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView4.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub DataGridView5_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView5.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView5.SelectedRows(0)
            txtID5.Text = dr.Cells(0).Value.ToString()
            txtCopperWireName.Text = dr.Cells(1).Value.ToString()
            btnUpdate4.Enabled = True
            btnDelete4.Enabled = True
            btnSave4.Enabled = False
            txtCopperWireName.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView5_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView5.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView5.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView5.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub btnNew1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew1.Click
        Reset1()
    End Sub

    Private Sub BtnDelete1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete1.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord1()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(txtShaftName.Text)) = 0 Then
            MessageBox.Show("Please enter shaft name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtShaftName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtShaftLength.Text)) = 0 Then
            MessageBox.Show("Please enter shaft length", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtShaftLength.Focus()
            Exit Sub
        End If
        If PictureBox1.Image Is Nothing Then
            MessageBox.Show("Please browse and select image", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            btnBrowse.Focus()
            Exit Sub
        End If
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim cd As String = "Update Shaft set Shaft_Name=@d1,Length=@d2,S_Image=@d3 where ID=" & txtID1.Text & ""
            cmd = New OleDbCommand(cd)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtShaftName.Text)
            cmd.Parameters.AddWithValue("@d2", txtShaftLength.Text)
            Dim ms As New MemoryStream()
            Dim bmpImage As New Bitmap(PictureBox1.Image)
            bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim p As New OleDbParameter("@d3", OleDbType.VarBinary)
            p.Value = data
            cmd.Parameters.Add(p)
            cmd.ExecuteNonQuery()
            con.Close()
            btnUpdate.Enabled = False
            GetData()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtSearchByShaftName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchByShaftName.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT * from Shaft where Shaft_Name like '" & txtSearchByShaftName.Text & "%'  Order by Shaft_Name", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView1.Rows.Clear()
            While (rdr.Read() = True)
                DataGridView1.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TabControl1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        FullReset()
    End Sub

    Private Sub BtnSave1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave1.Click
        If Len(Trim(txtStampName.Text)) = 0 Then
            MessageBox.Show("Please enter stamp name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtStampName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtStampOD.Text)) = 0 Then
            MessageBox.Show("Please enter OD", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtStampOD.Focus()
            Exit Sub
        End If
        If Len(Trim(txtStampID.Text)) = 0 Then
            MessageBox.Show("Please enter ID", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtStampID.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbStampType.Text)) = 0 Then
            MessageBox.Show("Please select stamp type", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbStampType.Focus()
            Exit Sub
        End If
        If PictureBox2.Image Is Nothing Then
            MessageBox.Show("Please browse and select image", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            btnBrowse1.Focus()
            Exit Sub
        End If
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select Stamping_name from Stamping where Stamping_name=@d1"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtStampName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Stamp Name Already Exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtStampName.Text = ""
                txtStampName.Focus()
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            If Val(txtStampID.Text) = Val(txtStampOD.Text) Then
                MessageBox.Show("Stamp OD and ID are same", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cd As String = "insert Into Stamping(Stamping_Name, Stamping_Od, Stamping_Id, Stamping_Type, Stamping_Image) VALUES (@d1,@d2,@d3,@d4,@d5)"
            cmd = New OleDbCommand(cd)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtStampName.Text)
            cmd.Parameters.AddWithValue("@d2", txtStampOD.Text)
            cmd.Parameters.AddWithValue("@d3", txtStampID.Text)
            cmd.Parameters.AddWithValue("@d4", cmbStampType.Text)
            Dim ms As New MemoryStream()
            Dim bmpImage As New Bitmap(PictureBox2.Image)
            bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim p As New OleDbParameter("@d5", OleDbType.VarBinary)
            p.Value = data
            cmd.Parameters.Add(p)
            cmd.ExecuteNonQuery()
            con.Close()
            BtnSave1.Enabled = False
            GetData1()
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtShaftLength_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtShaftLength.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtShaftLength.Text
            Dim selectionStart = Me.txtShaftLength.SelectionStart
            Dim selectionLength = Me.txtShaftLength.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an Integereger that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub txtStampOD_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtStampOD.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtStampOD.Text
            Dim selectionStart = Me.txtStampOD.SelectionStart
            Dim selectionLength = Me.txtStampOD.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an Integereger that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub txtStampID_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtStampID.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtStampID.Text
            Dim selectionStart = Me.txtStampID.SelectionStart
            Dim selectionLength = Me.txtStampID.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an Integereger that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub txtCommutatorOD_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCommutatorOD.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtCommutatorOD.Text
            Dim selectionStart = Me.txtCommutatorOD.SelectionStart
            Dim selectionLength = Me.txtCommutatorOD.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an Integereger that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub txtCommutatorID_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCommutatorID.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtCommutatorID.Text
            Dim selectionStart = Me.txtCommutatorID.SelectionStart
            Dim selectionLength = Me.txtCommutatorID.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an Integereger that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub txtCommutatorCopperLength_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCommutatorCopperLength.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtCommutatorCopperLength.Text
            Dim selectionStart = Me.txtCommutatorCopperLength.SelectionStart
            Dim selectionLength = Me.txtCommutatorCopperLength.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an Integereger that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub txtFanOD_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFanOD.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtFanOD.Text
            Dim selectionStart = Me.txtFanOD.SelectionStart
            Dim selectionLength = Me.txtFanOD.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an Integereger that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub txtFanID_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFanID.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtFanID.Text
            Dim selectionStart = Me.txtFanID.SelectionStart
            Dim selectionLength = Me.txtFanID.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an Integereger that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub txtFanWidth_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFanWidth.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtFanWidth.Text
            Dim selectionStart = Me.txtFanWidth.SelectionStart
            Dim selectionLength = Me.txtFanWidth.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an Integereger that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub btnUpdate1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate1.Click
        If Len(Trim(txtStampName.Text)) = 0 Then
            MessageBox.Show("Please enter stamp name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtStampName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtStampOD.Text)) = 0 Then
            MessageBox.Show("Please enter OD", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtStampOD.Focus()
            Exit Sub
        End If
        If Len(Trim(txtStampID.Text)) = 0 Then
            MessageBox.Show("Please enter ID", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtStampID.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbStampType.Text)) = 0 Then
            MessageBox.Show("Please select stamp type", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbStampType.Focus()
            Exit Sub
        End If
        If PictureBox2.Image Is Nothing Then
            MessageBox.Show("Please browse and select image", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            btnBrowse1.Focus()
            Exit Sub
        End If
        Try
            If Val(txtStampID.Text) = Val(txtStampOD.Text) Then
                MessageBox.Show("Stamp OD and ID are same", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cd As String = "Update Stamping set Stamping_Name=@d1, Stamping_Od=@d2, Stamping_Id=@d3, Stamping_Type=@d4, Stamping_Image=@d5 where ID=" & txtID2.Text & ""
            cmd = New OleDbCommand(cd)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtStampName.Text)
            cmd.Parameters.AddWithValue("@d2", txtStampOD.Text)
            cmd.Parameters.AddWithValue("@d3", txtStampID.Text)
            cmd.Parameters.AddWithValue("@d4", cmbStampType.Text)
            Dim ms As New MemoryStream()
            Dim bmpImage As New Bitmap(PictureBox2.Image)
            bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim p As New OleDbParameter("@d5", OleDbType.VarBinary)
            p.Value = data
            cmd.Parameters.Add(p)
            cmd.ExecuteNonQuery()
            con.Close()
            btnUpdate1.Enabled = False
            GetData1()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtSearchByStampName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchByStampName.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT * from Stamping where Stamping_Name like '" & txtSearchByStampName.Text & "%'  Order by Stamping_Name", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView2.Rows.Clear()
            While (rdr.Read() = True)
                DataGridView2.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnNew2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew2.Click
        Reset2()
    End Sub

    Private Sub btnSave2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave2.Click
        If Len(Trim(txtCommutatorName.Text)) = 0 Then
            MessageBox.Show("Please enter Commutator name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCommutatorName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCommutatorOD.Text)) = 0 Then
            MessageBox.Show("Please enter OD", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCommutatorOD.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCommutatorID.Text)) = 0 Then
            MessageBox.Show("Please enter ID", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCommutatorID.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCommutatorCopperLength.Text)) = 0 Then
            MessageBox.Show("Please enter copper length", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCommutatorCopperLength.Focus()
            Exit Sub
        End If
        If PictureBox3.Image Is Nothing Then
            MessageBox.Show("Please browse and select image", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            btnBrowse2.Focus()
            Exit Sub
        End If
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select Commutator_name from Commutator where Commutator_name=@d1"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtCommutatorName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Commutator Name Already Exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtCommutatorName.Text = ""
                txtCommutatorName.Focus()
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            If Val(txtCommutatorID.Text) = Val(txtCommutatorOD.Text) Then
                MessageBox.Show("Commutator OD and ID are same", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cd As String = "insert Into Commutator(Commutator_Name, C_Od, C_Id,Copper_Length,C_Image) VALUES (@d1,@d2,@d3,@d4,@d5)"
            cmd = New OleDbCommand(cd)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtCommutatorName.Text)
            cmd.Parameters.AddWithValue("@d2", txtCommutatorOD.Text)
            cmd.Parameters.AddWithValue("@d3", txtCommutatorID.Text)
            cmd.Parameters.AddWithValue("@d4", txtCommutatorCopperLength.Text)
            Dim ms As New MemoryStream()
            Dim bmpImage As New Bitmap(PictureBox3.Image)
            bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim p As New OleDbParameter("@d5", OleDbType.VarBinary)
            p.Value = data
            cmd.Parameters.Add(p)
            cmd.ExecuteNonQuery()
            con.Close()
            btnSave2.Enabled = False
            GetData2()
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete2.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord2()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate2.Click
        If Len(Trim(txtCommutatorName.Text)) = 0 Then
            MessageBox.Show("Please enter Commutator name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCommutatorName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCommutatorOD.Text)) = 0 Then
            MessageBox.Show("Please enter OD", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCommutatorOD.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCommutatorID.Text)) = 0 Then
            MessageBox.Show("Please enter ID", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCommutatorID.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCommutatorCopperLength.Text)) = 0 Then
            MessageBox.Show("Please enter copper length", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCommutatorCopperLength.Focus()
            Exit Sub
        End If
        If PictureBox3.Image Is Nothing Then
            MessageBox.Show("Please browse and select image", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            btnBrowse2.Focus()
            Exit Sub
        End If
        Try
            If Val(txtCommutatorID.Text) = Val(txtCommutatorOD.Text) Then
                MessageBox.Show("Commutator OD and ID are same", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cd As String = "Update Commutator set Commutator_Name=@d1, C_Od=@d2, C_Id=@d3,Copper_Length=@d4,C_Image=@d5 where ID=" & txtID3.Text & ""
            cmd = New OleDbCommand(cd)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtCommutatorName.Text)
            cmd.Parameters.AddWithValue("@d2", txtCommutatorOD.Text)
            cmd.Parameters.AddWithValue("@d3", txtCommutatorID.Text)
            cmd.Parameters.AddWithValue("@d4", txtCommutatorCopperLength.Text)
            Dim ms As New MemoryStream()
            Dim bmpImage As New Bitmap(PictureBox3.Image)
            bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim p As New OleDbParameter("@d5", OleDbType.VarBinary)
            p.Value = data
            cmd.Parameters.Add(p)
            cmd.ExecuteNonQuery()
            con.Close()
            btnUpdate2.Enabled = False
            GetData2()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtSearchByCommutatorName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchByCommutatorName.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT * from Commutator where Commutator_Name like '" & txtSearchByCommutatorName.Text & "%' order by Commutator_Name", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView3.Rows.Clear()
            While (rdr.Read() = True)
                DataGridView3.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnNew3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew3.Click
        Reset3()
    End Sub

    Private Sub btnSave3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave3.Click
        If Len(Trim(txtFanName.Text)) = 0 Then
            MessageBox.Show("Please enter Fan name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFanName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtFanOD.Text)) = 0 Then
            MessageBox.Show("Please enter OD", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFanOD.Focus()
            Exit Sub
        End If
        If Len(Trim(txtFanID.Text)) = 0 Then
            MessageBox.Show("Please enter ID", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFanID.Focus()
            Exit Sub
        End If
        If Len(Trim(txtFanWidth.Text)) = 0 Then
            MessageBox.Show("Please enter width", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFanWidth.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbFanStep.Text)) = 0 Then
            MessageBox.Show("Please select step", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbFanStep.Focus()
            Exit Sub
        End If
        If PictureBox4.Image Is Nothing Then
            MessageBox.Show("Please browse and select image", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            btnBrowse3.Focus()
            Exit Sub
        End If
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select Fan_name from Fan where Fan_name=@d1"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtFanName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Fan Name Already Exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtFanName.Text = ""
                txtFanName.Focus()
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            If Val(txtFanID.Text) = Val(txtFanOD.Text) Then
                MessageBox.Show("Fan OD and ID are same", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cd As String = "insert Into Fan(Fan_Name, F_Od, F_Id,F_Width,F_Step,F_Image) VALUES (@d1,@d2,@d3,@d4,@d5,@d6)"
            cmd = New OleDbCommand(cd)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtFanName.Text)
            cmd.Parameters.AddWithValue("@d2", txtFanOD.Text)
            cmd.Parameters.AddWithValue("@d3", txtFanID.Text)
            cmd.Parameters.AddWithValue("@d4", txtFanWidth.Text)
            cmd.Parameters.AddWithValue("@d5", cmbFanStep.Text)
            Dim ms As New MemoryStream()
            Dim bmpImage As New Bitmap(PictureBox4.Image)
            bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim p As New OleDbParameter("@d6", OleDbType.VarBinary)
            p.Value = data
            cmd.Parameters.Add(p)
            cmd.ExecuteNonQuery()
            con.Close()
            btnSave3.Enabled = False
            GetData3()
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete3.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord3()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete4.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord4()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate3.Click
        If Len(Trim(txtFanName.Text)) = 0 Then
            MessageBox.Show("Please enter Fan name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFanName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtFanOD.Text)) = 0 Then
            MessageBox.Show("Please enter OD", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFanOD.Focus()
            Exit Sub
        End If
        If Len(Trim(txtFanID.Text)) = 0 Then
            MessageBox.Show("Please enter ID", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFanID.Focus()
            Exit Sub
        End If
        If Len(Trim(txtFanWidth.Text)) = 0 Then
            MessageBox.Show("Please enter width", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFanWidth.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbFanStep.Text)) = 0 Then
            MessageBox.Show("Please select step", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbFanStep.Focus()
            Exit Sub
        End If
        If PictureBox4.Image Is Nothing Then
            MessageBox.Show("Please browse and select image", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            btnBrowse3.Focus()
            Exit Sub
        End If
        Try
            
            If Val(txtFanID.Text) = Val(txtFanOD.Text) Then
                MessageBox.Show("Fan OD and ID are same", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cd As String = "Update Fan set Fan_Name=@d1, F_Od=@d2, F_Id=@d3,F_Width=@d4,F_Step=@d5,F_Image=@d6 where ID=" & txtID4.Text & ""
            cmd = New OleDbCommand(cd)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtFanName.Text)
            cmd.Parameters.AddWithValue("@d2", txtFanOD.Text)
            cmd.Parameters.AddWithValue("@d3", txtFanID.Text)
            cmd.Parameters.AddWithValue("@d4", txtFanWidth.Text)
            cmd.Parameters.AddWithValue("@d5", cmbFanStep.Text)
            Dim ms As New MemoryStream()
            Dim bmpImage As New Bitmap(PictureBox4.Image)
            bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim p As New OleDbParameter("@d6", OleDbType.VarBinary)
            p.Value = data
            cmd.Parameters.Add(p)
            cmd.ExecuteNonQuery()
            con.Close()
            btnUpdate4.Enabled = False
            GetData3()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtSearchByFanName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchByFanName.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT * from Fan where Fan_Name like '" & txtSearchByFanName.Text & "%' order by Fan_Name", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView4.Rows.Clear()
            While (rdr.Read() = True)
                DataGridView4.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnNew4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew4.Click
        Reset4()
    End Sub

    Private Sub btnSave4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave4.Click
        If Len(Trim(txtCopperWireName.Text)) = 0 Then
            MessageBox.Show("Please enter copper wire name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCopperWireName.Focus()
            Exit Sub
        End If
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select Copper_Wire from CopperWire where Copper_Wire=@d1"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtCopperWireName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Copper Wire Name Already Exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtCopperWireName.Text = ""
                txtCopperWireName.Focus()
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cd As String = "insert Into CopperWire(Copper_Wire) VALUES (@d1)"
            cmd = New OleDbCommand(cd)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtCopperWireName.Text)
            cmd.ExecuteNonQuery()
            con.Close()
            btnSave4.Enabled = False
            GetData4()
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate4.Click
        If Len(Trim(txtCopperWireName.Text)) = 0 Then
            MessageBox.Show("Please enter copper wire name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCopperWireName.Focus()
            Exit Sub
        End If
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select Copper_Wire from CopperWire where Copper_Wire=@d1"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtCopperWireName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Copper Wire Name Already Exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtCopperWireName.Text = ""
                txtCopperWireName.Focus()
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cd As String = "Update CopperWire set Copper_Wire=@d1 where ID=" & txtID5.Text & ""
            cmd = New OleDbCommand(cd)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtCopperWireName.Text)
            cmd.ExecuteNonQuery()
            con.Close()
            btnUpdate4.Enabled = False
            GetData4()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtSearchByCopperWireName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchByCopperWireName.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT * from CopperWire where Copper_Wire like '" & txtSearchByCopperWireName.Text & "%' order by Copper_Wire", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView5.Rows.Clear()
            While (rdr.Read() = True)
                DataGridView5.Rows.Add(rdr(0), rdr(1))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmRawMaterialCategory_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub
End Class
