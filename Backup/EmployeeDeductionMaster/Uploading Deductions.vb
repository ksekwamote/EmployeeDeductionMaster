Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class Uploading_Deductions

    Dim total As Decimal

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, "Format")
        DataGridView1.AllowUserToAddRows = False
        LinkLabel1.Text = "Export"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If checkExpectedDeductions(DateTimePicker1.Value, employerid) = True Then
            DataGridView1.AllowUserToAddRows = False
            Button1.Enabled = False
            Button2.Enabled = True

            If decision("Do you want to load the deductions ?") = DialogResult.Yes Then
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
                        If checkCustomerDetails(row.Cells(0).Value.ToString, employerid) = True Then

                        Else
                            If row.Cells(3).Value = "" Then
                                row.Cells(3).Value = "Customer Does Not Exists"
                            Else
                                row.Cells(3).Value = row.Cells(3).Value & "/ Customer Does Not Exists"
                            End If
                            Button2.Enabled = False
                            Label2.Enabled = True
                        End If
                    Next

                    total = 0
                    TextBox1.Text = ""
                    For Each row As DataGridViewRow In DataGridView1.Rows
                        total = total + Decimal.Parse(row.Cells(1).Value.ToString)
                    Next
                    TextBox1.Text = total

                End If
            End If
        Else
            MessageBox.Show("There are no expected deductions available", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            DataGridView1.AllowUserToAddRows = False
            Button1.Enabled = False
            Button2.Enabled = True
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


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If decision("Are you sure you want to proceed?") = DialogResult.Yes Then
            If DataGridView1.Rows.Count > 0 Then
                For Each row As DataGridViewRow In DataGridView1.Rows
                    saveActualDeduction(row.Cells(0).Value.ToString, Decimal.Parse(row.Cells(1).Value.ToString), getLastDateOfMonth(DateTimePicker1.Value), Date.Parse(row.Cells(2).Value.ToString), Date.Now, employerid)
                Next
                MessageBox.Show("All actual deductions has been saved successfully", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                wrapper.saveAction("Uploaded Deduction", Date.Now, logedinname, "")
                Close()
            End If
        Else
            MessageBox.Show("Nothing to save", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker1.Value = getLastDateOfMonth(DateTimePicker1.Value)
        If checkActualDeduction(DateTimePicker1.Value, employerid) = True Then
            Button1.Enabled = False
            MessageBox.Show("There are existing actual deductions for this month", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If decision("Do you want to load other deductions for the same month ?") = DialogResult.Yes Then
                Button1.Enabled = True
            End If
        Else
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Uploading_Deductions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = getLastDateOfMonth(addMonth(getLastActualDeductionDate(employerid), 1))
        loadReceiptsBatch(employerid)
    End Sub

    Private Sub loadReceiptsBatch(ByVal employerid As String)
        DataGridView2.DataSource = Nothing
        Dim qry As String = "select SUM(ActualDeduction),ForMonth,EmployerName from ActualDeductions inner join Employers on ActualDeductions.EmployerID = Employers.AutoID where EmployerID = '" & employerid & "' group by ForMonth,EmployerName order by ForMonth DESC"

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
                DataGridView2.DataSource = bs
                BindingNavigator1.BindingSource = bs
                DataGridView2.Columns(DataGridView2.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Else
                DataGridView2.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadReceiptsBatch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadReceiptsBatch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

End Class