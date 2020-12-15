Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class UploadingCustomerDetails

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, "Format")
        DataGridView1.AllowUserToAddRows = False
        LinkLabel1.Text = "Export"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnload.Click
        DataGridView1.AllowUserToAddRows = False
        btnload.Enabled = False
        btnsave.Enabled = True

        If decision("Do you want to load the customers ?") = DialogResult.Yes Then
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

                ProgressBar1.Maximum = DataGridView1.Rows.Count
                For Each row As DataGridViewRow In DataGridView1.Rows

                    If employerid = Integer.Parse(row.Cells(2).Value.ToString) Then

                    Else
                        If row.Cells(4).Value = "" Then
                            row.Cells(4).Value = "The Employer Name Not the Same as the One Logged Into (" & employername & ")"
                        Else
                            row.Cells(4).Value = row.Cells(4).Value & "/ The Employer Name Not the Same as the One Logged Into (" & employername & ")"
                        End If
                        btnsave.Enabled = False
                        Label2.Enabled = True
                    End If

                    If checkCustomerDetailsInDatabase(row.Cells(0).Value.ToString) = True Then
                        If checkCustomerDetails(row.Cells(0).Value.ToString, Integer.Parse(row.Cells(2).Value.ToString)) = False Then
                            If row.Cells(4).Value = "" Then
                                row.Cells(4).Value = "Customer Exists (" & getEmployerNameByCustomerID(row.Cells(0).Value.ToString) & "), Not In The Employer Specified"
                            Else
                                row.Cells(4).Value = row.Cells(4).Value & "/ Customer Exists (" & getEmployerNameByCustomerID(row.Cells(0).Value.ToString) & "), Not In The Employer Specified"
                            End If
                            btnsave.Enabled = False
                            Label2.Enabled = True
                        Else
                            If row.Cells(4).Value = "" Then
                                row.Cells(4).Value = "Customer Exists (" & getEmployerNameByCustomerID(row.Cells(0).Value.ToString) & ") In The Employer Specified"
                            Else
                                row.Cells(4).Value = row.Cells(4).Value & "/ Customer Exists (" & getEmployerNameByCustomerID(row.Cells(0).Value.ToString) & ") In The Employer Specified"
                            End If
                            btnsave.Enabled = False
                            Label2.Enabled = True
                        End If
                    Else

                    End If

                    If checkEmployer(Integer.Parse(row.Cells(2).Value.ToString)) = True Then

                    Else
                        If row.Cells(4).Value = "" Then
                            row.Cells(4).Value = "Employer Does not exist"
                        Else
                            row.Cells(4).Value = row.Cells(4).Value & "/ Employer Does not exist"
                        End If
                        btnsave.Enabled = False
                        Label2.Enabled = True
                    End If

                    If row.Cells(3).Value.ToString = "Female" Or row.Cells(3).Value.ToString = "Male" Then

                    Else
                        If row.Cells(4).Value = "" Then
                            row.Cells(4).Value = "Gender is incorrect"
                        Else
                            row.Cells(4).Value = row.Cells(4).Value & "/ Gender is incorrect"
                        End If
                        btnsave.Enabled = False
                        Label2.Enabled = True
                    End If

                    ProgressBar1.Value = row.Index
                Next
                ProgressBar1.Value = 0
            End If
        End If
    End Sub

    Public Sub ExportToExcel(ByVal DataGridView1 As DataGridView, ByVal title As String)
        'verfying the datagridview having data or not
        If ((DataGridView1.Columns.Count = 0) Or (DataGridView1.Rows.Count = 0)) Then
            Exit Sub
        End If

        'Creating dataset to export
        Dim dset As New DataSet
        'add table to dataset
        dset.Tables.Add()
        'add column to that table
        For i As Integer = 0 To DataGridView1.ColumnCount - 1
            dset.Tables(0).Columns.Add(DataGridView1.Columns(i).HeaderText)
        Next
        'add rows to the table
        Dim dr1 As DataRow
        For i As Integer = 0 To DataGridView1.RowCount - 1
            dr1 = dset.Tables(0).NewRow
            For j As Integer = 0 To DataGridView1.Columns.Count - 1
                dr1(j) = DataGridView1.Rows(i).Cells(j).Value
            Next
            dset.Tables(0).Rows.Add(dr1)
        Next

        Dim excel As New Microsoft.Office.Interop.Excel.ApplicationClass
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

        wBook = excel.Workbooks.Add()
        wSheet = wBook.ActiveSheet()

        Dim dt As System.Data.DataTable = dset.Tables(0)
        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 1

        For Each dc In dt.Columns
            colIndex = colIndex + 1
            excel.Cells(2, colIndex) = dc.ColumnName
        Next
        'put excel heading
        excel.Cells(1, 1) = title
        excel.Range("A1:D1").Merge()

        For Each dr In dt.Rows
            rowIndex = rowIndex + 1
            colIndex = 0
            For Each dc In dt.Columns
                colIndex = colIndex + 1
                excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)

            Next
        Next

        wSheet.Columns.AutoFit()
        excel.Visible = True

    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave.Click
        If decision("Are you sure you want to proceed?") = DialogResult.Yes Then
            If DataGridView1.Rows.Count > 0 Then
                ProgressBar1.Maximum = DataGridView1.Rows.Count
                For Each row As DataGridViewRow In DataGridView1.Rows
                    saveCustomerDetails(row.Cells(0).Value.ToString, row.Cells(1).Value.ToString, Integer.Parse(row.Cells(2).Value.ToString), row.Cells(3).Value.ToString)
                    wrapper.saveAction("Customer Uploaded", Date.Now, logedinname, row.Cells(0).Value.ToString)
                    ProgressBar1.Value = row.Index
                Next
                ProgressBar1.Value = 0
                MessageBox.Show("Customer details saved successfully", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Close()
            End If
        Else
            MessageBox.Show("Nothing to save", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If decision("Are you sure you want to proceed?") = DialogResult.Yes Then
            If DataGridView1.Rows.Count > 0 Then
                ProgressBar1.Maximum = DataGridView1.Rows.Count
                For Each row As DataGridViewRow In DataGridView1.Rows
                    updateCustomerDetails(row.Cells(0).Value.ToString, row.Cells(1).Value.ToString, Integer.Parse(row.Cells(2).Value.ToString), row.Cells(3).Value.ToString)
                    wrapper.saveAction("Customer Name Updated", Date.Now, logedinname, row.Cells(0).Value.ToString)
                    ProgressBar1.Value = row.Index
                Next
                ProgressBar1.Value = 0
                MessageBox.Show("Customer details updated successfully", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Close()
            End If
        Else
            MessageBox.Show("Nothing to save", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub UploadingCustomerDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class