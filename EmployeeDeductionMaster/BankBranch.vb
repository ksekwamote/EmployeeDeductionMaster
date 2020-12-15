Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class BankBranch

    Dim selected As Boolean = False

    Private Sub loadBankBranch()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select BankBranchID,BankBranchName,BankBranch.BankID,BankName,BranchCode from BankBranch inner join Bank on BankBranch.BankID = Bank.BankID order by BankBranchName ASC"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim dt As DataTable

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "DIM")
            If ds.Tables(0).Rows.Count > 0 Then
                dt = ds.Tables(0)
                Dim bs As New BindingSource
                bs.DataSource = dt.DefaultView
                DataGridView1.DataSource = bs
                BindingNavigator1.BindingSource = bs
                DataGridView1.Columns(DataGridView1.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Dim filterManager As New DgvFilterManager
                filterManager = New DgvFilterManager(DataGridView1)
            Else
                DataGridView1.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "loadBankBranch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadBankBranch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If tbbranchid.Text.Trim = "" Or tbbranchname.Text.Trim = "" Or tbbankname.Text.Trim = "" Or tbbranchcode.Text.Trim = "" Then
                MessageBox.Show("Enter all the details before saving", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    saveBankBranch(tbbranchid.Text.Trim, tbbranchname.Text.Trim, tbbankid.Text.Trim, "Active", tbbranchcode.Text.Trim)
                    loadBankBranch()
                    If decision("Add another branch in the bank ?") = DialogResult.Yes Then
                        clearTextExcludeBank()
                    Else
                        clearAllText()
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            selected = True
            tbbranchid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
            tbbranchname.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
            tbbankid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString()
            tbbankname.Text = getValueString("select BankName from Bank where BankID = '" & DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString() & "'")
            tbbranchcode.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString()
            Button2.Enabled = True
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If tbbranchid.Text = "" Or tbbranchname.Text.Trim = "" Or tbbankname.Text.Trim = "" Or tbbranchcode.Text.Trim = "" Then
                MessageBox.Show("Make Sure to select the country to update changes and enter the changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    updateBankBranch(tbbranchid.Text.Trim, tbbranchname.Text.Trim, tbbranchcode.Text.Trim, tbbankid.Text.Trim)
                    loadBankBranch()
                    clearAllText()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        clearAllText()
    End Sub

    Public Sub clearAllText()
        Button2.Enabled = False
        Button1.Enabled = True
        tbbranchid.Text = ""
        tbbranchname.Text = ""
        tbbankid.Text = ""
        tbbankname.Text = ""
        tbbranchcode.Text = ""
        selected = False
    End Sub

    Public Sub clearTextExcludeBank()
        Button2.Enabled = False
        Button1.Enabled = True
        tbbranchid.Text = ""
        tbbranchname.Text = ""
        tbbranchcode.Text = ""
        selected = False
    End Sub

    Private Sub tbbankname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbbranchname.TextChanged
        If Button1.Enabled = True Then
            If tbbranchname.Text.Trim.Length > 3 And tbbankname.Text.Trim.Length > 3 Then
                If selected = True Then

                Else
                    tbbranchid.Text = getNextBankBranchID(tbbankname.Text.Substring(0, 3) & tbbranchname.Text.Substring(0, 3))
                End If
            End If
        End If
    End Sub

    Private Sub BankBranch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadBankBranch()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim vrl As New BankList
        vrl.MdiParent = mform
        vrl.whichform = "BankBranch"
        vrl.bb = Me
        vrl.Show()
    End Sub

    Private Sub tbbankname_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbbankname.TextChanged
        If Button1.Enabled = True Then
            If tbbranchname.Text.Trim.Length > 3 And tbbankname.Text.Trim.Length > 3 Then
                If selected = True Then

                Else
                    tbbranchid.Text = getNextBankBranchID(tbbankname.Text.Substring(0, 3) & tbbranchname.Text.Substring(0, 3))
                End If
            End If
        End If
    End Sub
End Class