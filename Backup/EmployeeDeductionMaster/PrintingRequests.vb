Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class PrintingRequests

    Dim username As String

    Private Sub PrintingRequests_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadRequests(username)
    End Sub

    Private Sub loadRequests(ByVal username As String)
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select AutoID,Name as PrintFor,'Sheet'+SheetNumber as SheetNumber,RequestTime,RequestedBy,Status from PrintingRequests inner join Users on PrintingRequests.Username = Users.Username where PrintingRequests.Username = '" & username & "' order by RequestTime DESC"

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
            MessageBox.Show(ex.Message, "loadRequests()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadRequests()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        loadRequests(username)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btsave.Click
        If tbname.Text.Trim = "" Or tbsheetnumber.Text.Trim = "" Then
            MessageBox.Show("Select the individual to print for and write the sheet number", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If decision("Are you to proceed ?") = DialogResult.Yes Then
                savePrintingRequest(tbusername.Text.Trim, tbsheetnumber.Text.Trim, Date.Now, logedinname, "Requested", Date.Now)
                loadRequests(username)
                cleartext()
                MessageBox.Show("Request Sent", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bclear.Click
        cleartext()
    End Sub

    Public Sub cleartext()
        tbautoid.Text = ""
        tbsheetnumber.Text = ""
        btsave.Enabled = True
        bupdate.Enabled = False
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            tbautoid.Text = DataGridView1.CurrentRow.Cells(0).Value
            tbusername.Text = getUsernameByName(DataGridView1.CurrentRow.Cells(1).Value)
            username = tbusername.Text
            tbname.Text = DataGridView1.CurrentRow.Cells(1).Value
            tbsheetnumber.Text = getSheetNumber(Integer.Parse(DataGridView1.CurrentRow.Cells(0).Value))
            If DataGridView1.CurrentRow.Cells(5).Value = "Requested" Then
                btsave.Enabled = False
                bupdate.Enabled = True
            Else
                btsave.Enabled = False
                bupdate.Enabled = False
            End If
        End If
    End Sub

    Private Sub bupdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bupdate.Click
        If tbname.Text.Trim = "" Or tbsheetnumber.Text.Trim = "" Or tbautoid.Text.Trim = "" Then
            MessageBox.Show("Select the individual to print for and write the sheet number", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If decision("Are you to proceed ?") = DialogResult.Yes Then
                updatePrintingRequest(tbusername.Text.Trim, tbsheetnumber.Text.Trim, Date.Now, logedinname, Integer.Parse(tbautoid.Text.Trim))
                loadRequests(username)
                cleartext()
                MessageBox.Show("Request Updated", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim vrl As New UserList
        vrl.MdiParent = mform
        vrl.whichform = "Print"
        vrl.pr = Me
        vrl.Show()
    End Sub

    Private Sub tbname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbname.TextChanged
        username = getUsernameByName(tbname.Text.Trim)
        loadRequests(username)
    End Sub
End Class