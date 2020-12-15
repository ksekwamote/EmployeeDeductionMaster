Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class ModeOfPayment

    Private Sub loadModes()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select * from ModeOfPayment"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
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
                BindingNavigator1.BindingSource = bs
                DataGridView1.Columns(DataGridView1.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Else
                DataGridView1.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "loadModes()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadModes()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Branches_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadModes()
    End Sub

    Public Sub saveMode(ByVal ModeName As String, ByVal Status As String, ByVal SameDayBanking As String, ByVal NoBankingConfirmation As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            transactioninsert = con.BeginTransaction()
            Dim qryinsert As String = "Insert into ModeOfPayment values ('" & ModeName & "','" & Status & "','" & SameDayBanking & "','" & NoBankingConfirmation & "')"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            cmdinsert.Transaction = transactioninsert
            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "saveMode()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveMode()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btsave.Click
        If tbmodename.Text.Trim = "" Then
            MessageBox.Show("Enter all the details before saving")
        Else
            saveMode(tbmodename.Text.Trim, "Active", tbproceed.Text.Trim, tbconfirmation.Text.Trim)
            loadModes()
            tbmodename.Text = ""
            btupdate.Enabled = False
            btsave.Enabled = True
            tbmodeid.Text = ""
            tbmodename.Text = ""
            tbproceed.Text = ""
            tbconfirmation.Text = ""
        End If
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            tbmodeid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
            tbmodename.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
            tbproceed.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString()
            cbproceed.SelectedIndex = cbproceed.FindStringExact(tbproceed.Text)
            tbconfirmation.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString()
            cbconfirmation.SelectedIndex = cbconfirmation.FindStringExact(tbconfirmation.Text)
            btupdate.Enabled = True
            btsave.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btupdate.Click
        If tbmodeid.Text = "" Then
            MessageBox.Show("Make Sure to select the branch to update changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            updateMode(tbmodename.Text.Trim, "Active", tbproceed.Text.Trim, tbconfirmation.Text.Trim, Integer.Parse(tbmodeid.Text.Trim))
            loadModes()
            btupdate.Enabled = False
            btsave.Enabled = True
            tbmodeid.Text = ""
            tbmodename.Text = ""
            tbproceed.Text = ""
            tbconfirmation.Text = ""
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        btupdate.Enabled = False
        btsave.Enabled = True
        tbmodeid.Text = ""
        tbmodename.Text = ""
        tbproceed.Text = ""
        tbconfirmation.Text = ""
    End Sub

    Public Sub updateMode(ByVal ModeName As String, ByVal Status As String, ByVal SameDayBanking As String, ByVal NoBankingConfirmation As String, ByVal ModeID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "update ModeOfPayment set ModeName = '" & ModeName & "',Status = '" & Status & "',SameDayBanking = '" & SameDayBanking & "',NoBankingConfirmation = '" & NoBankingConfirmation & "' where ModeID = '" & ModeID & "'"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            transactioninsert = con.BeginTransaction()
            cmdinsert.Transaction = transactioninsert

            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "updateBranch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateBranch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Private Sub cbproceed_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbproceed.SelectedIndexChanged
        tbproceed.Text = cbproceed.Text
    End Sub

    Private Sub cbconfirmation_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbconfirmation.SelectedIndexChanged
        tbconfirmation.Text = cbconfirmation.Text
    End Sub
End Class