Imports System.Data.OleDb
Imports System.IO

Public Class frmProduct

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
    Sub fillShaft()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct Shaft_Name from Shaft", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbShaftName.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbShaftName.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub fillStamping()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct Stamping_Name from Stamping", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbStampingName.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbStampingName.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub fillCommutator()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct Commutator_Name from Commutator", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbCommutatorName.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbCommutatorName.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub fillFan()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct Fan_Name from Fan", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbFanName.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbFanName.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub fillCopperWire()
        Try
            Dim CN As New OleDbConnection(cs)
            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct Copper_Wire from CopperWire", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbCopperWire.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbCopperWire.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmProduct_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillCommutator()
        fillCopperWire()
        fillFan()
        fillShaft()
        fillStamping()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveShaft.Click
        Try
            If (ListView1.SelectedItems.Count > 0) Then
                Dim itmCnt, i, t As Integer
                ListView1.FocusedItem.Remove()
                itmCnt = ListView1.Items.Count
                t = 1
                For i = 1 To itmCnt + 1

                    'Dim lst1 As New ListViewItem(i)
                    'ListView1.Items(i).SubItems(0).Text = t
                    t = t + 1
                Next
                btnRemoveShaft.Enabled = False
                cmbShaftName.Focus()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRmoveStamping_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveStamping.Click
        Try
            If (ListView2.SelectedItems.Count > 0) Then
                Dim itmCnt, i, t As Integer
                ListView2.FocusedItem.Remove()
                itmCnt = ListView2.Items.Count
                t = 1
                For i = 1 To itmCnt + 1

                    'Dim lst1 As New ListViewItem(i)
                    'ListView1.Items(i).SubItems(0).Text = t
                    t = t + 1
                Next
                btnRemoveStamping.Enabled = False
                cmbStampingName.Focus()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRemoveCommutator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveCommutator.Click
        Try
            If (ListView3.SelectedItems.Count > 0) Then
                Dim itmCnt, i, t As Integer
                ListView3.FocusedItem.Remove()
                itmCnt = ListView3.Items.Count
                t = 1
                For i = 1 To itmCnt + 1

                    'Dim lst1 As New ListViewItem(i)
                    'ListView1.Items(i).SubItems(0).Text = t
                    t = t + 1
                Next
                btnRemoveCommutator.Enabled = False
                cmbCommutatorName.Focus()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRemoveCopperWire_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveCopperWire.Click
        Try
            If (ListView5.SelectedItems.Count > 0) Then
                Dim itmCnt, i, t As Integer
                ListView5.FocusedItem.Remove()
                itmCnt = ListView5.Items.Count
                t = 1
                For i = 1 To itmCnt + 1

                    'Dim lst1 As New ListViewItem(i)
                    'ListView1.Items(i).SubItems(0).Text = t
                    t = t + 1
                Next
                btnRemoveCopperWire.Enabled = False
                cmbCopperWire.Focus()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveFan.Click
        Try
            If (ListView4.SelectedItems.Count > 0) Then
                Dim itmCnt, i, t As Integer
                ListView4.FocusedItem.Remove()
                itmCnt = ListView4.Items.Count
                t = 1
                For i = 1 To itmCnt + 1

                    'Dim lst1 As New ListViewItem(i)
                    'ListView1.Items(i).SubItems(0).Text = t
                    t = t + 1
                Next
                btnRemoveFan.Enabled = False
                cmbFanName.Focus()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        btnRemoveShaft.Enabled = True
    End Sub

    Private Sub ListView2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView2.SelectedIndexChanged
        btnRemoveStamping.Enabled = True
    End Sub

    Private Sub ListView3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView3.SelectedIndexChanged
        btnRemoveCommutator.Enabled = True
    End Sub

    Private Sub ListView5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView5.SelectedIndexChanged

        btnRemoveCopperWire.Enabled = True
    End Sub

    Private Sub ListView4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView4.SelectedIndexChanged
        btnRemoveFan.Enabled = True
    End Sub

    Private Sub cmbShaftName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbShaftName.SelectedIndexChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT ID,Length FROM Shaft WHERE Shaft_Name=@d1"
            cmd.Parameters.AddWithValue("@d1", cmbShaftName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtShaftID.Text = rdr.GetValue(0)
                txtShaftLength.Text = rdr.GetValue(1)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub cmbStampingName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbStampingName.SelectedIndexChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT ID,Stamping_Od, Stamping_Id, Stamping_Type FROM Stamping WHERE Stamping_Name=@d1"
            cmd.Parameters.AddWithValue("@d1", cmbStampingName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtStampPKID.Text = rdr.GetValue(0)
                txtStampOD.Text = rdr.GetValue(1)
                txtStampID.Text = rdr.GetValue(2)
                txtStampType.Text = rdr.GetString(3)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub cmbCommutatorName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCommutatorName.SelectedIndexChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT ID,C_Od, C_Id,Copper_Length FROM Commutator WHERE Commutator_Name=@d1"
            cmd.Parameters.AddWithValue("@d1", cmbCommutatorName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtCommutatorPKID.Text = rdr.GetValue(0)
                txtCommutatorOD.Text = rdr.GetValue(1)
                txtCommutatorID.Text = rdr.GetValue(2)
                txtCommutatorCopperLength.Text = rdr.GetValue(3)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub cmbCopperWire_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCopperWire.SelectedIndexChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT ID FROM CopperWire WHERE Copper_Wire=@d1"
            cmd.Parameters.AddWithValue("@d1", cmbCopperWire.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtCopperWireID.Text = rdr.GetValue(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub cmbFanName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFanName.SelectedIndexChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT ID,F_Od, F_Id,F_Width,F_Step FROM Fan WHERE Fan_Name=@d1"
            cmd.Parameters.AddWithValue("@d1", cmbFanName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtFanPkID.Text = rdr.GetValue(0)
                txtFanOD.Text = rdr.GetValue(1)
                txtFanID.Text = rdr.GetValue(2)
                txtFanWidth.Text = rdr.GetValue(3)
                txtStep.Text = rdr.GetString(4)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnAddShaft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddShaft.Click
        Try
            If Len(Trim(cmbShaftName.Text)) = 0 Then
                MessageBox.Show("Please select shaft name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbShaftName.Focus()
                Exit Sub
            End If

            Dim temp As Integer
            temp = ListView1.Items.Count()
            If temp = 0 Then
                Dim i As Integer
                Dim lst As New ListViewItem(i)
                lst.SubItems.Add(txtShaftID.Text)
                lst.SubItems.Add(cmbShaftName.Text)
                lst.SubItems.Add(txtShaftLength.Text)
                ListView1.Items.Add(lst)
                i = i + 1
                cmbShaftName.Text = ""
                txtShaftID.Text = ""
                txtShaftLength.Text = ""
                cmbShaftName.Focus()
                Exit Sub
            End If
            For j = 0 To temp - 1
                If (ListView1.Items(j).SubItems(2).Text = cmbShaftName.Text) Then
                    MessageBox.Show("Shaft Name Already added", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    cmbShaftName.Text = ""
                    txtShaftID.Text = ""
                    txtShaftLength.Text = ""
                    cmbShaftName.Focus()
                    Exit Sub
                End If
            Next j
            Dim k As Integer
            Dim lst1 As New ListViewItem(k)
            lst1.SubItems.Add(txtShaftID.Text)
            lst1.SubItems.Add(cmbShaftName.Text)
            lst1.SubItems.Add(txtShaftLength.Text)
            ListView1.Items.Add(lst1)
            k = k + 1
            cmbShaftName.Text = ""
            txtShaftID.Text = ""
            txtShaftLength.Text = ""
            cmbShaftName.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnAddStamping_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddStamping.Click
        Try
            If Len(Trim(cmbStampingName.Text)) = 0 Then
                MessageBox.Show("Please select stamping name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbStampingName.Focus()
                Exit Sub
            End If

            Dim temp As Integer
            temp = ListView2.Items.Count()
            If temp = 0 Then
                Dim i As Integer
                Dim lst As New ListViewItem(i)
                lst.SubItems.Add(txtStampPKID.Text)
                lst.SubItems.Add(cmbStampingName.Text)
                lst.SubItems.Add(txtStampOD.Text)
                lst.SubItems.Add(txtStampID.Text)
                lst.SubItems.Add(txtStampType.Text)
                ListView2.Items.Add(lst)
                i = i + 1
                txtStampPKID.Text = ""
                cmbStampingName.Text = ""
                txtStampOD.Text = ""
                txtStampID.Text = ""
                txtStampType.Text = ""
                cmbStampingName.Focus()
                Exit Sub
            End If
            For j = 0 To temp - 1
                If (ListView2.Items(j).SubItems(2).Text = cmbStampingName.Text) Then
                    MessageBox.Show("Stamping Name Already added", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtStampPKID.Text = ""
                    cmbStampingName.Text = ""
                    txtStampOD.Text = ""
                    txtStampID.Text = ""
                    txtStampType.Text = ""
                    cmbStampingName.Focus()
                    Exit Sub
                End If
            Next j
            Dim k As Integer
            Dim lst1 As New ListViewItem(k)
            lst1.SubItems.Add(txtStampPKID.Text)
            lst1.SubItems.Add(cmbStampingName.Text)
            lst1.SubItems.Add(txtStampOD.Text)
            lst1.SubItems.Add(txtStampID.Text)
            lst1.SubItems.Add(txtStampType.Text)
            ListView2.Items.Add(lst1)
            k = k + 1
            txtStampPKID.Text = ""
            cmbStampingName.Text = ""
            txtStampOD.Text = ""
            txtStampID.Text = ""
            txtStampType.Text = ""
            cmbStampingName.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If Len(Trim(cmbCommutatorName.Text)) = 0 Then
                MessageBox.Show("Please select commutator name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbCommutatorName.Focus()
                Exit Sub
            End If

            Dim temp As Integer
            temp = ListView3.Items.Count()
            If temp = 0 Then
                Dim i As Integer
                Dim lst As New ListViewItem(i)
                lst.SubItems.Add(txtCommutatorPKID.Text)
                lst.SubItems.Add(cmbCommutatorName.Text)
                lst.SubItems.Add(txtCommutatorOD.Text)
                lst.SubItems.Add(txtCommutatorID.Text)
                lst.SubItems.Add(txtCommutatorCopperLength.Text)
                ListView3.Items.Add(lst)
                i = i + 1
                txtCommutatorPKID.Text = ""
                cmbCommutatorName.Text = ""
                txtCommutatorOD.Text = ""
                txtCommutatorID.Text = ""
                txtCommutatorCopperLength.Text = ""
                cmbCommutatorName.Focus()
                Exit Sub
            End If
            For j = 0 To temp - 1
                If (ListView3.Items(j).SubItems(2).Text = cmbCommutatorName.Text) Then
                    MessageBox.Show("Commutator Name Already added", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtCommutatorPKID.Text = ""
                    cmbCommutatorName.Text = ""
                    txtCommutatorOD.Text = ""
                    txtCommutatorID.Text = ""
                    txtCommutatorCopperLength.Text = ""
                    cmbCommutatorName.Focus()
                    Exit Sub
                End If
            Next j
            Dim k As Integer
            Dim lst1 As New ListViewItem(k)
            lst1.SubItems.Add(txtCommutatorPKID.Text)
            lst1.SubItems.Add(cmbCommutatorName.Text)
            lst1.SubItems.Add(txtCommutatorOD.Text)
            lst1.SubItems.Add(txtCommutatorID.Text)
            lst1.SubItems.Add(txtCommutatorCopperLength.Text)
            ListView3.Items.Add(lst1)
            k = k + 1
            txtCommutatorPKID.Text = ""
            cmbCommutatorName.Text = ""
            txtCommutatorOD.Text = ""
            txtCommutatorID.Text = ""
            txtCommutatorCopperLength.Text = ""
            cmbCommutatorName.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddFan.Click
        Try
            If Len(Trim(cmbFanName.Text)) = 0 Then
                MessageBox.Show("Please select Fan name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbFanName.Focus()
                Exit Sub
            End If

            Dim temp As Integer
            temp = ListView4.Items.Count()
            If temp = 0 Then
                Dim i As Integer
                Dim lst As New ListViewItem(i)
                lst.SubItems.Add(txtFanPkID.Text)
                lst.SubItems.Add(cmbFanName.Text)
                lst.SubItems.Add(txtFanOD.Text)
                lst.SubItems.Add(txtFanID.Text)
                lst.SubItems.Add(txtFanWidth.Text)
                lst.SubItems.Add(txtStep.Text)
                ListView4.Items.Add(lst)
                i = i + 1
                txtFanPkID.Text = ""
                cmbFanName.Text = ""
                txtFanOD.Text = ""
                txtFanID.Text = ""
                txtFanWidth.Text = ""
                txtStep.Text = ""
                cmbFanName.Focus()
                Exit Sub
            End If
            For j = 0 To temp - 1
                If (ListView4.Items(j).SubItems(2).Text = cmbFanName.Text) Then
                    MessageBox.Show("Fan Name Already added", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtFanPkID.Text = ""
                    cmbFanName.Text = ""
                    txtFanOD.Text = ""
                    txtFanID.Text = ""
                    txtFanWidth.Text = ""
                    txtStep.Text = ""
                    cmbFanName.Focus()
                    Exit Sub
                End If
            Next j
            Dim k As Integer
            Dim lst1 As New ListViewItem(k)
            lst1.SubItems.Add(txtFanPkID.Text)
            lst1.SubItems.Add(cmbFanName.Text)
            lst1.SubItems.Add(txtFanOD.Text)
            lst1.SubItems.Add(txtFanID.Text)
            lst1.SubItems.Add(txtFanWidth.Text)
            lst1.SubItems.Add(txtStep.Text)
            ListView4.Items.Add(lst1)
            k = k + 1
            txtFanPkID.Text = ""
            cmbFanName.Text = ""
            txtFanOD.Text = ""
            txtFanID.Text = ""
            txtFanWidth.Text = ""
            txtStep.Text = ""
            cmbFanName.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnAddCopperWire_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddCopperWire.Click
        Try
            If Len(Trim(cmbCopperWire.Text)) = 0 Then
                MessageBox.Show("Please select Copper Wire", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbCopperWire.Focus()
                Exit Sub
            End If

            Dim temp As Integer
            temp = ListView5.Items.Count()
            If temp = 0 Then
                Dim i As Integer
                Dim lst As New ListViewItem(i)
                lst.SubItems.Add(txtCopperWireID.Text)
                lst.SubItems.Add(cmbCopperWire.Text)
                ListView5.Items.Add(lst)
                i = i + 1
                txtCopperWireID.Text = ""
                cmbCopperWire.Text = ""
                cmbCopperWire.Focus()
                Exit Sub
            End If
            For j = 0 To temp - 1
                If (ListView5.Items(j).SubItems(2).Text = cmbCopperWire.Text) Then
                    MessageBox.Show("Copper wire Already added", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtCopperWireID.Text = ""
                    cmbCopperWire.Text = ""
                    cmbCopperWire.Focus()
                    Exit Sub
                End If
            Next j
            Dim k As Integer
            Dim lst1 As New ListViewItem(k)
            lst1.SubItems.Add(txtCopperWireID.Text)
            lst1.SubItems.Add(cmbCopperWire.Text)
            ListView5.Items.Add(lst1)
            k = k + 1
            txtCopperWireID.Text = ""
            cmbCopperWire.Text = ""
            cmbCopperWire.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Reset()
        ListView1.Items.Clear()
        ListView2.Items.Clear()
        ListView3.Items.Clear()
        ListView4.Items.Clear()
        ListView5.Items.Clear()
        txtFanPkID.Text = ""
        cmbFanName.Text = ""
        txtFanOD.Text = ""
        txtFanID.Text = ""
        txtFanWidth.Text = ""
        txtStep.Text = ""
        txtCommutatorPKID.Text = ""
        cmbCommutatorName.Text = ""
        txtCommutatorOD.Text = ""
        txtCommutatorID.Text = ""
        txtCommutatorCopperLength.Text = ""
        txtStampPKID.Text = ""
        cmbStampingName.Text = ""
        txtStampOD.Text = ""
        txtStampID.Text = ""
        txtStampType.Text = ""
        cmbShaftName.Text = ""
        txtShaftID.Text = ""
        txtShaftLength.Text = ""
        txtCopperWireID.Text = ""
        cmbCopperWire.Text = ""
        btnRemoveCommutator.Enabled = False
        btnRemoveCopperWire.Enabled = False
        btnRemoveFan.Enabled = False
        btnRemoveShaft.Enabled = False
        btnRemoveStamping.Enabled = False
        PictureBox1.Image = Nothing
        PictureBox2.Image = Nothing
        btnSave.Enabled = True
        btnUpdate.Enabled = False
        btnDelete.Enabled = False
        txtProductName.Text = ""
        btnPrint.Enabled = False
        txtProductName.Focus()
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
            Dim ct As String = "delete from Product_Shaft where ProductID=" & txtProductID.Text & ""
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct1 As String = "delete from Product_Stamping where ProductID=" & txtProductID.Text & ""
            cmd = New OleDbCommand(ct1)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct2 As String = "delete from Product_Fan where ProductID=" & txtProductID.Text & ""
            cmd = New OleDbCommand(ct2)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct3 As String = "delete from Product_Commutator where ProductID=" & txtProductID.Text & ""
            cmd = New OleDbCommand(ct3)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct4 As String = "delete from Product_CopperWire where ProductID=" & txtProductID.Text & ""
            cmd = New OleDbCommand(ct4)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con.Open()
            Dim cq1 As String = "delete from Product where ID=" & txtProductID.Text & ""
            cmd = New OleDbCommand(cq1)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtProductName.Text)) = 0 Then
            MessageBox.Show("Please enter product name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtProductName.Focus()
            Exit Sub
        End If
        If PictureBox1.Image Is Nothing Then
            MessageBox.Show("Please browse and select image", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            btnBrowse.Focus()
            Exit Sub
        End If
        If PictureBox2.Image Is Nothing Then
            MessageBox.Show("Please browse and select image", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            btnBrowse1.Focus()
            Exit Sub
        End If
        If ListView1.Items.Count = 0 Then
            MessageBox.Show("sorry no shaft info added in a list", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbShaftName.Focus()
            Exit Sub
        End If
        If ListView2.Items.Count = 0 Then
            MessageBox.Show("sorry no stamping info added in a list", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbStampingName.Focus()
            Exit Sub
        End If
        If ListView3.Items.Count = 0 Then
            MessageBox.Show("sorry no commutator info added in a list", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbCommutatorName.Focus()
            Exit Sub
        End If
        If ListView4.Items.Count = 0 Then
            MessageBox.Show("sorry no fan info added in a list", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbFanName.Focus()
            Exit Sub
        End If
        If ListView5.Items.Count = 0 Then
            MessageBox.Show("sorry no copper wire added in a list", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbCopperWire.Focus()
            Exit Sub
        End If

        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select Product_name from Product where Product_name=@d1"
            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtProductName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Product Name Already Exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtProductName.Text = ""
                txtProductName.Focus()
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "insert Into Product(Product_Name,Image1,Image2) VALUES (@d1,@d2,@d3)"
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtProductName.Text)
            Dim ms, ms1 As New MemoryStream()
            Dim bmpImage As New Bitmap(PictureBox1.Image)
            Dim bmpImage1 As New Bitmap(PictureBox2.Image)
            bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            bmpImage1.Save(ms1, System.Drawing.Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim data1 As Byte() = ms1.GetBuffer()
            Dim p As New OleDbParameter("@d2", OleDbType.VarBinary)
            Dim p1 As New OleDbParameter("@d3", OleDbType.VarBinary)
            p.Value = data
            p1.Value = data1
            cmd.Parameters.Add(p)
            cmd.Parameters.Add(p1)
            cmd.ExecuteNonQuery()
            con.Close()
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT ID from Product WHERE Product_Name=@d1"
            cmd.Parameters.AddWithValue("@d1", txtProductName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtProductID.Text = rdr.GetValue(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            For i = 0 To ListView1.Items.Count - 1
                con = New OleDbConnection(cs)
                Dim cd As String = "insert Into Product_Shaft(ProductID,ShaftID) VALUES (@d1,@d2)"
                cmd = New OleDbCommand(cd)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", txtProductID.Text)
                cmd.Parameters.AddWithValue("@d2", ListView1.Items(i).SubItems(1).Text)
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            For i = 0 To ListView2.Items.Count - 1
                con = New OleDbConnection(cs)
                Dim cd As String = "insert Into Product_Stamping(ProductID,StampingID) VALUES (@d1,@d2)"
                cmd = New OleDbCommand(cd)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", txtProductID.Text)
                cmd.Parameters.AddWithValue("@d2", ListView2.Items(i).SubItems(1).Text)
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            For i = 0 To ListView3.Items.Count - 1
                con = New OleDbConnection(cs)
                Dim cd As String = "insert Into Product_Commutator(ProductID,CommutatorID) VALUES (@d1,@d2)"
                cmd = New OleDbCommand(cd)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", txtProductID.Text)
                cmd.Parameters.AddWithValue("@d2", ListView3.Items(i).SubItems(1).Text)
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            For i = 0 To ListView4.Items.Count - 1
                con = New OleDbConnection(cs)
                Dim cd As String = "insert Into Product_Fan(ProductID,FanID) VALUES (@d1,@d2)"
                cmd = New OleDbCommand(cd)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", txtProductID.Text)
                cmd.Parameters.AddWithValue("@d2", ListView4.Items(i).SubItems(1).Text)
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            For i = 0 To ListView5.Items.Count - 1
                con = New OleDbConnection(cs)
                Dim cd As String = "insert Into Product_CopperWire(ProductID,CopperWireID) VALUES (@d1,@d2)"
                cmd = New OleDbCommand(cd)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", txtProductID.Text)
                cmd.Parameters.AddWithValue("@d2", ListView5.Items(i).SubItems(1).Text)
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            btnSave.Enabled = False
            btnPrint.Enabled = True
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmProduct_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Me.Hide()
        frmMain.Show()
    End Sub

  
    Private Sub btnGetData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetData.Click
        frmProductRecord.ShowDialog()
    End Sub

  
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(txtProductName.Text)) = 0 Then
            MessageBox.Show("Please enter product name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtProductName.Focus()
            Exit Sub
        End If
        If PictureBox1.Image Is Nothing Then
            MessageBox.Show("Please browse and select image", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            btnBrowse.Focus()
            Exit Sub
        End If
        If PictureBox2.Image Is Nothing Then
            MessageBox.Show("Please browse and select image", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            btnBrowse1.Focus()
            Exit Sub
        End If
        If ListView1.Items.Count = 0 Then
            MessageBox.Show("sorry no shaft info added in a list", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbShaftName.Focus()
            Exit Sub
        End If
        If ListView2.Items.Count = 0 Then
            MessageBox.Show("sorry no stamping info added in a list", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbStampingName.Focus()
            Exit Sub
        End If
        If ListView3.Items.Count = 0 Then
            MessageBox.Show("sorry no commutator info added in a list", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbCommutatorName.Focus()
            Exit Sub
        End If
        If ListView4.Items.Count = 0 Then
            MessageBox.Show("sorry no fan info added in a list", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbFanName.Focus()
            Exit Sub
        End If
        If ListView5.Items.Count = 0 Then
            MessageBox.Show("sorry no copper wire added in a list", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbCopperWire.Focus()
            Exit Sub
        End If

        Try

            con = New OleDbConnection(cs)
            con.Open()
            Dim cb As String = "Update Product set Product_Name=@d1,Image1=@d2,Image2=@d3 where ID=" & txtProductID.Text & ""
            cmd = New OleDbCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtProductName.Text)
            Dim ms, ms1 As New MemoryStream()
            Dim bmpImage As New Bitmap(PictureBox1.Image)
            Dim bmpImage1 As New Bitmap(PictureBox2.Image)
            bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            bmpImage1.Save(ms1, System.Drawing.Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim data1 As Byte() = ms1.GetBuffer()
            Dim p As New OleDbParameter("@d2", OleDbType.VarBinary)
            Dim p1 As New OleDbParameter("@d3", OleDbType.VarBinary)
            p.Value = data
            p1.Value = data1
            cmd.Parameters.Add(p)
            cmd.Parameters.Add(p1)
            cmd.ExecuteNonQuery()
            con.Close()
            con = New OleDbConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT ID from Product WHERE Product_Name=@d1"
            cmd.Parameters.AddWithValue("@d1", txtProductName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtProductID.Text = rdr.GetValue(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            For i = 0 To ListView1.Items.Count - 1
                con = New OleDbConnection(cs)
                con.Open()
                Dim ct As String = "delete from Product_Shaft where ProductID=" & txtProductID.Text & ""
                cmd = New OleDbCommand(ct)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
            Next
            For i = 0 To ListView1.Items.Count - 1
                con = New OleDbConnection(cs)
                Dim cd As String = "insert Into Product_Shaft(ProductID,ShaftID) VALUES (@d1,@d2)"
                cmd = New OleDbCommand(cd)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", txtProductID.Text)
                cmd.Parameters.AddWithValue("@d2", ListView1.Items(i).SubItems(1).Text)
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            For i = 0 To ListView2.Items.Count - 1
                con = New OleDbConnection(cs)
                con.Open()
                Dim ct As String = "delete from Product_Stamping where ProductID=" & txtProductID.Text & ""
                cmd = New OleDbCommand(ct)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
            Next
            For i = 0 To ListView2.Items.Count - 1
                con = New OleDbConnection(cs)
                Dim cd As String = "insert Into Product_Stamping(ProductID,StampingID) VALUES (@d1,@d2)"
                cmd = New OleDbCommand(cd)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", txtProductID.Text)
                cmd.Parameters.AddWithValue("@d2", ListView2.Items(i).SubItems(1).Text)
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            For i = 0 To ListView3.Items.Count - 1
                con = New OleDbConnection(cs)
                con.Open()
                Dim ct As String = "delete from Product_Commutator where ProductID=" & txtProductID.Text & ""
                cmd = New OleDbCommand(ct)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
            Next
            For i = 0 To ListView3.Items.Count - 1
                con = New OleDbConnection(cs)
                Dim cd As String = "insert Into Product_Commutator(ProductID,CommutatorID) VALUES (@d1,@d2)"
                cmd = New OleDbCommand(cd)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", txtProductID.Text)
                cmd.Parameters.AddWithValue("@d2", ListView3.Items(i).SubItems(1).Text)
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            For i = 0 To ListView4.Items.Count - 1
                con = New OleDbConnection(cs)
                con.Open()
                Dim ct As String = "delete from Product_Fan where ProductID=" & txtProductID.Text & ""
                cmd = New OleDbCommand(ct)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
            Next
            For i = 0 To ListView4.Items.Count - 1
                con = New OleDbConnection(cs)
                Dim cd As String = "insert Into Product_Fan(ProductID,FanID) VALUES (@d1,@d2)"
                cmd = New OleDbCommand(cd)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", txtProductID.Text)
                cmd.Parameters.AddWithValue("@d2", ListView4.Items(i).SubItems(1).Text)
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            For i = 0 To ListView5.Items.Count - 1
                con = New OleDbConnection(cs)
                con.Open()
                Dim ct As String = "delete from Product_CopperWire where ProductID=" & txtProductID.Text & ""
                cmd = New OleDbCommand(ct)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
            Next
            For i = 0 To ListView5.Items.Count - 1
                con = New OleDbConnection(cs)
                Dim cd As String = "insert Into Product_CopperWire(ProductID,CopperWireID) VALUES (@d1,@d2)"
                cmd = New OleDbCommand(cd)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", txtProductID.Text)
                cmd.Parameters.AddWithValue("@d2", ListView5.Items(i).SubItems(1).Text)
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            btnUpdate.Enabled = False
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptProduct() 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New OleDbConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT Product_Name,Image1,Image2,Shaft_Name,Length,S_Image,Stamping_Name, Stamping_Od, Stamping_Id, Stamping_Type,Stamping_Image,Commutator_Name, C_Od, C_Id,Copper_Length,C_Image,Fan_Name, F_Od, F_Id,F_Width,F_Step,F_Image,Copper_Wire from Product,Shaft,Stamping,Commutator,Fan,CopperWire,Product_Shaft,Product_Stamping,Product_Commutator,Product_Fan,Product_CopperWire where Shaft.ID=Product_Shaft.ShaftID and Product.ID=Product_Shaft.ProductID and Stamping.ID=Product_Stamping.StampingID and Product.ID=Product_Stamping.ProductID and Commutator.ID=Product_Commutator.CommutatorID and Product.ID=Product_Commutator.ProductID and Fan.ID=Product_Fan.FanID and Product.ID=Product_Fan.ProductID and CopperWire.ID=Product_CopperWire.CopperWireID and Product.ID=Product_CopperWire.ProductID and Product.ID=" & txtProductID.Text & ""
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Shaft")
            myDA.Fill(myDS, "Product_Shaft")
            myDA.Fill(myDS, "Product")
            myDA.Fill(myDS, "Stamping")
            myDA.Fill(myDS, "Product_Stamping")
            myDA.Fill(myDS, "Commutator")
            myDA.Fill(myDS, "Product_Commutator")
            myDA.Fill(myDS, "Fan")
            myDA.Fill(myDS, "Product_Fan")
            myDA.Fill(myDS, "CopperWire")
            myDA.Fill(myDS, "Product_CopperWire")
            rpt.SetDataSource(myDS)
            frmProductReport.CrystalReportViewer1.ReportSource = rpt
            frmProductReport.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub
End Class