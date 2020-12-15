Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class ChangeCustomerEmployer

    Private Sub ChangeCustomerEmployer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DataGridView1.AllowUserToAddRows = False
        loadEmployers()
        tbemployername.Text = ""
        tbemployerid.Text = ""
    End Sub

    Public Sub loadEmployers()
        Dim qry As String = "select * from Employers"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "Emp")
            If ds.Tables(0).Rows.Count > 0 Then
                Me.ComboBox1.DataSource = ds
                Me.ComboBox1.DisplayMember = "Emp.EmployerName"
                Me.ComboBox1.ValueMember = "Emp.AutoID"
                Me.ComboBox1.SelectedIndex = 0
            Else

            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView1.AllowUserToAddRows = False
        Button1.Enabled = False
        Button2.Enabled = True

        If decision("Do you want to load data ?") = DialogResult.Yes Then
            Dim csvfile As String
            Dim OpenFileDialog As New OpenFileDialog
            OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            OpenFileDialog.Filter = "CSV Files (*.csv)|*.csv"

            If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then

                Dim fi As New FileInfo(OpenFileDialog.FileName)
                Dim FileName As String = OpenFileDialog.FileName
                csvfile = fi.FullName
                Dim counter As Integer = 1
                For Each line As String In System.IO.File.ReadAllLines(csvfile)
                    If counter = 1 Then
                    Else
                        DataGridView1.Rows.Add(line.Split(","))
                    End If
                    counter = counter + 1
                Next

                For Each row As DataGridViewRow In DataGridView1.Rows
                    If checkCustomerInDatabase(row.Cells(0).Value) = True Then

                    Else
                        If row.Cells(1).Value = "" Then
                            row.Cells(1).Value = "Customer Does Not Exist"
                        Else
                            row.Cells(1).Value = row.Cells(1).Value & "/ Customer Does Not Exist"
                        End If
                        Button2.Enabled = False
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        tbemployername.Text = ComboBox1.Text
        tbemployerid.Text = ComboBox1.SelectedValue.ToString
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If decision("Are you sure you want to proceed?") = DialogResult.Yes Then
            For Each row As DataGridViewRow In DataGridView1.Rows
                updateCustomerEmployer(Integer.Parse(tbemployerid.Text.Trim), row.Cells(0).Value)
            Next
            MessageBox.Show("Done updating", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Close()
        End If
    End Sub
End Class