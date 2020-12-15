Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.ReportSource

Public Class Add_Products

    Private Sub loadList()
        DataGridView1.DataSource = Nothing

        Dim qry As String = "select * from Products order by AllocationOrder ASC"
        Dim ds As New DataSet
        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim dt As DataTable

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LoanDetails")
            If ds.Tables(0).Rows.Count > 0 Then
                dt = ds.Tables(0)
                Dim bs As New BindingSource
                bs.DataSource = dt.DefaultView
                DataGridView1.DataSource = bs
                DataGridView1.Columns(DataGridView1.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Else
                DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadList()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If decision("Are you sure to continue ?") = DialogResult.Yes Then
            If tbproductid.Text = "" Then
                If tbproductname.Text = "" Or tbproductallocationorder.Text = "" Or tbstatus.Text.Trim = "" Or tbautocalculate.Text.Trim = "" Or tbhassystem.Text.Trim = "" Or tbsystem.Text.Trim = "" Or tbuseincheques.Text.Trim = "" Or tbuseinsettlement.Text.Trim = "" Or tbuseinreceipts.Text.Trim = "" Then
                    MessageBox.Show("Enter all the details", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    saveProduct(tbproductname.Text.Trim, Integer.Parse(tbproductallocationorder.Text), tbstatus.Text.Trim, Integer.Parse(tbautocalculate.Text.Trim), Integer.Parse(tbhassystem.Text.Trim), tbsystem.Text.Trim, tbuseincheques.Text.Trim, tbuseinsettlement.Text.Trim, tbuseinreceipts.Text.Trim)
                    tbproductid.Text = ""
                    tbproductname.Text = ""
                    tbproductallocationorder.Text = ""
                    tbstatus.Text = ""
                    tbautocalculate.Text = ""
                    tbhassystem.Text = ""
                    tbsystem.Text = ""
                    tbuseincheques.Text = ""
                    tbuseinsettlement.Text = ""
                    tbuseinreceipts.Text = ""
                    loadList()
                    Button1.Enabled = True
                    Button2.Enabled = False
                    wrapper.saveAction("Add Product", Date.Now, logedinname, "")
                End If
            Else
                MessageBox.Show("Clear text first", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub Add_Products_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadList()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        tbproductid.Text = ""
        tbproductname.Text = ""
        tbproductallocationorder.Text = ""
        tbstatus.Text = ""
        tbautocalculate.Text = ""
        tbhassystem.Text = ""
        tbsystem.Text = ""
        tbuseincheques.Text = ""
        tbuseinsettlement.Text = ""
        tbuseinreceipts.Text = ""
        Button1.Enabled = True
        Button2.Enabled = False
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If decision("Are you sure to continue ?") = DialogResult.Yes Then
            If tbproductid.Text = "" Then
                MessageBox.Show("Select the record to update", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If tbproductname.Text = "" Or tbproductallocationorder.Text = "" Or tbstatus.Text.Trim = "" Or tbautocalculate.Text.Trim = "" Or tbhassystem.Text.Trim = "" Or tbsystem.Text.Trim = "" Or tbuseincheques.Text.Trim = "" Or tbuseinsettlement.Text.Trim = "" Or tbuseinreceipts.Text.Trim = "" Then
                    MessageBox.Show("Enter all the details", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    updateProduct(Integer.Parse(tbproductid.Text), tbproductname.Text.Trim, tbstatus.Text.Trim, Integer.Parse(tbproductallocationorder.Text), Integer.Parse(tbautocalculate.Text.Trim), Integer.Parse(tbhassystem.Text.Trim), tbsystem.Text.Trim, tbuseincheques.Text.Trim, tbuseinsettlement.Text.Trim, tbuseinreceipts.Text.Trim)
                    tbproductid.Text = ""
                    tbproductname.Text = ""
                    tbproductallocationorder.Text = ""
                    tbstatus.Text = ""
                    tbautocalculate.Text = ""
                    tbhassystem.Text = ""
                    tbsystem.Text = ""
                    tbuseincheques.Text = ""
                    tbuseinsettlement.Text = ""
                    tbuseinreceipts.Text = ""
                    Button1.Enabled = True
                    Button2.Enabled = False
                    wrapper.saveAction("Updated Product", Date.Now, logedinname, tbproductid.Text.Trim)
                    loadList()
                End If
            End If
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, "")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        tbstatus.Text = ComboBox1.Text
    End Sub

    Private Sub cbautocalculate_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbautocalculate.SelectedIndexChanged
        tbautocalculate.Text = cbautocalculate.Text
    End Sub

    Private Sub tbautocalculate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbautocalculate.TextChanged
        If tbautocalculate.Text.Trim = "1" Then
            Label6.Text = "Auto Allocate"
        ElseIf tbautocalculate.Text.Trim = "0" Then
            Label6.Text = "Do Not Auto Allocate"
        Else
            Label6.Text = "text here"
        End If
    End Sub

    Private Sub cbhassystem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbhassystem.SelectedIndexChanged
        tbhassystem.Text = cbhassystem.Text
    End Sub

    Private Sub tbhassystem_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbhassystem.TextChanged
        If tbhassystem.Text.Trim = "1" Then
            Label7.Text = "It Has"
        ElseIf tbhassystem.Text.Trim = "0" Then
            Label7.Text = "Do Not Have"
        Else
            Label7.Text = "text here"
        End If
    End Sub

    Private Sub cbsystem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbsystem.SelectedIndexChanged
        tbsystem.Text = cbsystem.Text
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        tbsystem.Text = tbproductname.Text.Trim
    End Sub

    Private Sub cbuseincheques_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbuseincheques.SelectedIndexChanged
        tbuseincheques.Text = cbuseincheques.Text
    End Sub

    Private Sub cbuseinsettlement_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbuseinsettlement.SelectedIndexChanged
        tbuseinsettlement.Text = cbuseinsettlement.Text
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        Try
            If DataGridView1.Rows.Count > 0 Then
                tbproductid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString
                tbproductname.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString
                tbproductallocationorder.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString
                tbstatus.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString
                ComboBox1.SelectedIndex = ComboBox1.FindStringExact(tbstatus.Text)
                If Boolean.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString) = True Then
                    tbautocalculate.Text = 1
                Else
                    tbautocalculate.Text = 0
                End If
                cbautocalculate.SelectedIndex = cbautocalculate.FindStringExact(tbautocalculate.Text)

                If Boolean.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(5).Value.ToString) = True Then
                    tbhassystem.Text = 1
                Else
                    tbhassystem.Text = 0
                End If
                cbhassystem.SelectedIndex = cbhassystem.FindStringExact(tbhassystem.Text)

                tbsystem.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(6).Value.ToString
                cbsystem.SelectedIndex = cbsystem.FindStringExact(tbsystem.Text)
                tbuseincheques.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(7).Value.ToString
                cbuseincheques.SelectedIndex = cbuseincheques.FindStringExact(tbuseincheques.Text)
                tbuseinsettlement.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(8).Value.ToString
                cbuseinsettlement.SelectedIndex = cbuseinsettlement.FindStringExact(tbuseinsettlement.Text)

                tbuseinreceipts.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(9).Value.ToString
                cbuseinreceipts.SelectedIndex = cbuseinreceipts.FindStringExact(tbuseinreceipts.Text)

                Button1.Enabled = False
                Button2.Enabled = True
            End If
        Catch ex As Exception
            Button1.Enabled = False
            Button2.Enabled = True
        End Try
    End Sub

    Private Sub cbuseinreceipts_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbuseinreceipts.SelectedIndexChanged
        tbuseinreceipts.Text = cbuseinreceipts.Text
    End Sub

End Class